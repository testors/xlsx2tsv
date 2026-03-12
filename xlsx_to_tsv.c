#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <time.h>
#include <stdbool.h>
#include <stdint.h>
#include <regex.h>
#include <sys/stat.h>

#include "miniz.h"
#include "filter.h"

#define MAX_CELL_VALUE 32768
#define BUFFER_SIZE 65536
#define MAX_SHEET_NAME 256
#define MAX_REL_ID 64
#define MAX_OUTPUT_FILENAME (MAX_SHEET_NAME + 32)
#define MAX_OUTPUT_PATH 4096
#define MAX_SHEET_STATE 16
#define MAX_WARNING_CODE 64
#define MAX_WARNING_MESSAGE 256
#define MAX_SHEET_WARNINGS 16
#define MAX_SHEET_FILTERS 128
#define TOOL_VERSION "0.1.0"

static FILE* g_log_fp = NULL;
#define printf(...) fprintf(g_log_fp ? g_log_fp : stdout, __VA_ARGS__)

typedef struct {
    char code[MAX_WARNING_CODE];
    char message[MAX_WARNING_MESSAGE];
} SheetWarning;

// 성능을 위한 공유 문자열 구조체
typedef struct {
    char** strings;
    int count;
    int capacity;
} SharedStrings;

// 시트 정보 구조체
typedef struct {
    char name[MAX_SHEET_NAME];
    char rel_id[MAX_REL_ID];
    char filename[MAX_SHEET_NAME];
    char output_name[MAX_OUTPUT_PATH];
    char state[MAX_SHEET_STATE];
    int sheet_id;
    bool hidden;
    bool selected;
    bool processed;
    bool truncated;
    int emitted_rows;
    int emitted_cols;
    int approx_rows;
    int approx_cols;
    SheetWarning warnings[MAX_SHEET_WARNINGS];
    int warning_count;
} SheetInfo;

// 워크북 구조체
typedef struct {
    SheetInfo* sheets;
    int sheet_count;
    int sheet_capacity;
} Workbook;

typedef enum {
    FORMAT_KIND_RAW,
    FORMAT_KIND_NUMBER,
    FORMAT_KIND_PERCENT,
    FORMAT_KIND_SCIENTIFIC,
    FORMAT_KIND_ZERO_PAD,
    FORMAT_KIND_DATE,
    FORMAT_KIND_TIME,
    FORMAT_KIND_DATETIME,
    FORMAT_KIND_DURATION
} FormatKind;

typedef struct {
    FormatKind kind;
    int decimals;
    int zero_width;
    bool use_thousands;
    bool show_seconds;
} CellFormat;

typedef struct {
    int num_fmt_id;
    char* format_code;
} CustomNumFormat;

typedef struct {
    CustomNumFormat* custom_formats;
    int custom_count;
    int custom_capacity;

    CellFormat* xf_formats;
    int xf_count;
    int xf_capacity;

    bool date_1904;
    bool enabled;
} Styles;

typedef struct {
    int start_col;
    int end_col;
} HiddenColumnRange;

typedef struct {
    HiddenColumnRange* ranges;
    int count;
    int capacity;
} HiddenColumns;

typedef struct {
    int start_row;
    int end_row;
    int start_col;
    int end_col;
    char value[MAX_CELL_VALUE];
    bool value_set;
} MergeRegion;

typedef struct {
    MergeRegion* regions;
    int count;
    int capacity;
} MergeRegions;

typedef struct {
    int col;
    char* value;
} RowCell;

typedef struct {
    RowCell* cells;
    int count;
    int capacity;
} RowBuffer;

typedef struct {
    bool list_sheets;
    bool json;
    bool manifest_stdout;
    const char* manifest_json_path;
    bool stdout_output;
    bool fail_if_no_sheet;
    bool fail_on_output_collision;
    int max_sheets;
    int max_rows_per_sheet;
    size_t max_output_bytes;
    bool fail_if_truncated;
    const char* output_dir;
    OutputFormat output_format;
    int start_row;
    bool game_db_fast_mode;
} RunOptions;

typedef struct {
    char messages[MAX_SHEET_WARNINGS][MAX_WARNING_MESSAGE];
    int count;
    bool truncated;
} GlobalWarnings;

void unescape_xml_entities(char* str);
void escape_tsv_value(const char* input, char* output, int max_len);

static void init_workbook(Workbook* wb) {
    wb->sheets = NULL;
    wb->sheet_count = 0;
    wb->sheet_capacity = 0;
}

static void free_workbook(Workbook* wb) {
    free(wb->sheets);
    init_workbook(wb);
}

static void add_sheet_warning(SheetInfo* sheet, const char* code, const char* message) {
    if (sheet->warning_count >= MAX_SHEET_WARNINGS) {
        return;
    }

    strncpy(sheet->warnings[sheet->warning_count].code, code, MAX_WARNING_CODE - 1);
    sheet->warnings[sheet->warning_count].code[MAX_WARNING_CODE - 1] = '\0';
    strncpy(sheet->warnings[sheet->warning_count].message, message, MAX_WARNING_MESSAGE - 1);
    sheet->warnings[sheet->warning_count].message[MAX_WARNING_MESSAGE - 1] = '\0';
    sheet->warning_count++;
}

static void add_global_warning(GlobalWarnings* warnings, const char* message) {
    if (warnings->count >= MAX_SHEET_WARNINGS) {
        return;
    }

    strncpy(warnings->messages[warnings->count], message, MAX_WARNING_MESSAGE - 1);
    warnings->messages[warnings->count][MAX_WARNING_MESSAGE - 1] = '\0';
    warnings->count++;
}

static void ensure_workbook_capacity(Workbook* wb, int needed) {
    if (needed <= wb->sheet_capacity) {
        return;
    }

    int new_capacity = wb->sheet_capacity ? wb->sheet_capacity * 2 : 16;
    while (new_capacity < needed) {
        new_capacity *= 2;
    }

    wb->sheets = realloc(wb->sheets, sizeof(SheetInfo) * new_capacity);
    wb->sheet_capacity = new_capacity;
}

// 빠른 XML 속성 찾기 - 제공된 버퍼에 쓰기
// 추출된 값의 길이를 반환, 찾지 못하면 -1 반환
static inline int find_attribute(const char* xml, const char* attr_name, int attr_len, char* out_buffer, int max_len) {
    const char* pos = strstr(xml, attr_name);
    if (!pos) return -1;

    pos += attr_len;

    // 공백 건너뛰기
    while (*pos && (*pos == ' ' || *pos == '\t')) pos++;

    // 여는 따옴표 찾기
    if (*pos != '"') return -1;
    pos++; // 여는 따옴표 건너뛰기

    // 닫는 따옴표 찾기
    const char* end = strchr(pos, '"');
    if (!end) return -1;

    int len = end - pos;
    if (len >= max_len) len = max_len - 1;

    memcpy(out_buffer, pos, len);
    out_buffer[len] = '\0';
    return len;
}

static const char* find_in_range(const char* start, const char* end,
                                 const char* needle, size_t needle_len) {
    const char* pos = start;

    if (!start || !end || start >= end || needle_len == 0) {
        return NULL;
    }

    while (pos < end) {
        size_t remaining = (size_t)(end - pos);
        const char* match;

        if (remaining < needle_len) {
            return NULL;
        }

        match = memchr(pos, needle[0], remaining - needle_len + 1);
        if (!match) {
            return NULL;
        }

        if (memcmp(match, needle, needle_len) == 0) {
            return match;
        }

        pos = match + 1;
    }

    return NULL;
}

static inline int find_attribute_in_range(const char* start, const char* end,
                                          const char* attr_name, int attr_len,
                                          char* out_buffer, int max_len) {
    const char* pos = find_in_range(start, end, attr_name, (size_t)attr_len);
    if (!pos) return -1;

    pos += attr_len;

    while (pos < end && (*pos == ' ' || *pos == '\t')) pos++;

    if (pos >= end || *pos != '"') return -1;
    pos++;

    const char* attr_end = memchr(pos, '"', (size_t)(end - pos));
    if (!attr_end) return -1;

    int len = (int)(attr_end - pos);
    if (len >= max_len) len = max_len - 1;

    memcpy(out_buffer, pos, (size_t)len);
    out_buffer[len] = '\0';
    return len;
}

// 일반적인 속성을 위한 헬퍼 매크로
#define FIND_ATTR_NAME(xml, buf, sz) find_attribute(xml, "name=", 5, buf, sz)
#define FIND_ATTR_SHEET_ID(xml, buf, sz) find_attribute(xml, "sheetId=", 8, buf, sz)
#define FIND_ATTR_REL_ID(xml, buf, sz) find_attribute(xml, "r:id=", 5, buf, sz)
#define FIND_ATTR_R(xml, buf, sz) find_attribute(xml, "r=", 2, buf, sz)
#define FIND_ATTR_S(xml, buf, sz) find_attribute(xml, "s=", 2, buf, sz)
#define FIND_ATTR_T(xml, buf, sz) find_attribute(xml, "t=", 2, buf, sz)
#define FIND_ATTR_ID(xml, buf, sz) find_attribute(xml, "Id=", 3, buf, sz)
#define FIND_ATTR_TARGET(xml, buf, sz) find_attribute(xml, "Target=", 7, buf, sz)
#define FIND_ATTR_TYPE(xml, buf, sz) find_attribute(xml, "Type=", 5, buf, sz)
#define FIND_ATTR_NUMFMT_ID(xml, buf, sz) find_attribute(xml, "numFmtId=", 9, buf, sz)
#define FIND_ATTR_FORMAT_CODE(xml, buf, sz) find_attribute(xml, "formatCode=", 11, buf, sz)
#define FIND_ATTR_MIN(xml, buf, sz) find_attribute(xml, "min=", 4, buf, sz)
#define FIND_ATTR_MAX(xml, buf, sz) find_attribute(xml, "max=", 4, buf, sz)
#define FIND_ATTR_HIDDEN(xml, buf, sz) find_attribute(xml, "hidden=", 7, buf, sz)
#define FIND_ATTR_REF(xml, buf, sz) find_attribute(xml, "ref=", 4, buf, sz)
#define FIND_ATTR_STATE(xml, buf, sz) find_attribute(xml, "state=", 6, buf, sz)

static inline bool is_xml_name_char(char c) {
    return isalnum((unsigned char)c) || c == '_' || c == ':' || c == '-' || c == '.';
}

static const char* find_tag_end_in_range(const char* tag_start, const char* end) {
    char quote = '\0';

    for (const char* pos = tag_start; pos < end; pos++) {
        if (quote) {
            if (*pos == quote) {
                quote = '\0';
            }
            continue;
        }

        if (*pos == '"' || *pos == '\'') {
            quote = *pos;
            continue;
        }

        if (*pos == '>') {
            return pos;
        }
    }

    return NULL;
}

static bool tag_has_local_name(const char* tag_start, const char* name, bool closing) {
    const char* pos = tag_start + 1;
    if (closing) {
        if (*pos != '/') {
            return false;
        }
        pos++;
    } else if (*pos == '/') {
        return false;
    }

    if (*pos == '!' || *pos == '?') {
        return false;
    }

    const char* name_start = pos;
    while (is_xml_name_char(*pos)) {
        pos++;
    }

    if (pos == name_start) {
        return false;
    }

    const char* local_name = name_start;
    for (const char* p = name_start; p < pos; p++) {
        if (*p == ':') {
            local_name = p + 1;
        }
    }

    size_t local_len = (size_t)(pos - local_name);
    size_t expected_len = strlen(name);
    if (local_len != expected_len) {
        return false;
    }

    return strncmp(local_name, name, expected_len) == 0;
}

static const char* find_next_tag_local(const char* pos, const char* end,
                                       const char* name, bool closing) {
    while (pos < end) {
        const char* tag_start = memchr(pos, '<', (size_t)(end - pos));
        if (!tag_start) {
            return NULL;
        }

        if (tag_has_local_name(tag_start, name, closing)) {
            return tag_start;
        }

        pos = tag_start + 1;
    }

    return NULL;
}

static inline const char* find_next_start_tag_local(const char* pos, const char* end, const char* name) {
    return find_next_tag_local(pos, end, name, false);
}

static inline const char* find_next_end_tag_local(const char* pos, const char* end, const char* name) {
    return find_next_tag_local(pos, end, name, true);
}

static bool is_self_closing_tag(const char* tag_start, const char* tag_end) {
    const char* pos = tag_end - 1;
    while (pos > tag_start && isspace((unsigned char)*pos)) {
        pos--;
    }
    return *pos == '/';
}

static void copy_tag_range(const char* start, const char* end, char* buffer, int max_len) {
    size_t len = (size_t)(end - start + 1);
    if (len >= (size_t)max_len) {
        len = (size_t)max_len - 1;
    }
    memcpy(buffer, start, len);
    buffer[len] = '\0';
}

static int append_text_range(const char* start, const char* end, char* out, int out_pos, int max_len) {
    while (start < end && out_pos < max_len - 1) {
        out[out_pos++] = *start++;
    }
    out[out_pos] = '\0';
    return out_pos;
}

static int extract_first_tag_text(const char* start, const char* end,
                                  const char* tag_name, char* out, int max_len) {
    const char* tag_start = find_next_start_tag_local(start, end, tag_name);
    if (!tag_start) {
        return 0;
    }

    const char* tag_end = find_tag_end_in_range(tag_start, end);
    if (!tag_end) {
        return 0;
    }

    if (is_self_closing_tag(tag_start, tag_end)) {
        out[0] = '\0';
        return 1;
    }

    const char* close_start = find_next_end_tag_local(tag_end + 1, end, tag_name);
    if (!close_start) {
        return 0;
    }

    int out_pos = append_text_range(tag_end + 1, close_start, out, 0, max_len);
    out[out_pos] = '\0';
    unescape_xml_entities(out);
    return 1;
}

static int extract_text_runs(const char* start, const char* end, char* out, int max_len) {
    const char* pos = start;
    int out_pos = 0;
    int found = 0;
    out[0] = '\0';

    while ((pos = find_next_start_tag_local(pos, end, "t")) != NULL) {
        const char* tag_end = find_tag_end_in_range(pos, end);
        if (!tag_end) {
            break;
        }

        if (is_self_closing_tag(pos, tag_end)) {
            pos = tag_end + 1;
            continue;
        }

        const char* close_start = find_next_end_tag_local(tag_end + 1, end, "t");
        if (!close_start) {
            break;
        }

        out_pos = append_text_range(tag_end + 1, close_start, out, out_pos, max_len);
        found = 1;

        const char* close_end = find_tag_end_in_range(close_start, end);
        if (!close_end) {
            break;
        }
        pos = close_end + 1;
    }

    if (found) {
        out[out_pos] = '\0';
        unescape_xml_entities(out);
    }

    return found;
}

static int extract_inline_string_text(const char* start, const char* end, char* out, int max_len) {
    const char* is_start = find_next_start_tag_local(start, end, "is");
    if (!is_start) {
        return extract_text_runs(start, end, out, max_len);
    }

    const char* is_tag_end = find_tag_end_in_range(is_start, end);
    if (!is_tag_end) {
        return 0;
    }

    if (is_self_closing_tag(is_start, is_tag_end)) {
        out[0] = '\0';
        return 1;
    }

    const char* is_close = find_next_end_tag_local(is_tag_end + 1, end, "is");
    if (!is_close) {
        return 0;
    }

    return extract_text_runs(is_tag_end + 1, is_close, out, max_len);
}

static CellFormat make_raw_cell_format(void) {
    CellFormat format;
    format.kind = FORMAT_KIND_RAW;
    format.decimals = 0;
    format.zero_width = 0;
    format.use_thousands = false;
    format.show_seconds = false;
    return format;
}

void init_styles(Styles* styles) {
    styles->custom_formats = NULL;
    styles->custom_count = 0;
    styles->custom_capacity = 0;
    styles->xf_formats = NULL;
    styles->xf_count = 0;
    styles->xf_capacity = 0;
    styles->date_1904 = false;
    styles->enabled = false;
}

void free_styles(Styles* styles) {
    for (int i = 0; i < styles->custom_count; i++) {
        free(styles->custom_formats[i].format_code);
    }
    free(styles->custom_formats);
    free(styles->xf_formats);
    init_styles(styles);
}

static bool attr_is_true(const char* attr_value) {
    return attr_value[0] == '1' || attr_value[0] == 't' || attr_value[0] == 'T';
}

static void init_hidden_columns(HiddenColumns* hidden_columns) {
    hidden_columns->ranges = NULL;
    hidden_columns->count = 0;
    hidden_columns->capacity = 0;
}

static void free_hidden_columns(HiddenColumns* hidden_columns) {
    free(hidden_columns->ranges);
    init_hidden_columns(hidden_columns);
}

static void add_hidden_column_range(HiddenColumns* hidden_columns, int start_col, int end_col) {
    if (start_col > end_col) {
        return;
    }

    if (hidden_columns->count >= hidden_columns->capacity) {
        hidden_columns->capacity = hidden_columns->capacity ? hidden_columns->capacity * 2 : 8;
        hidden_columns->ranges = realloc(hidden_columns->ranges,
                                         sizeof(HiddenColumnRange) * hidden_columns->capacity);
    }

    hidden_columns->ranges[hidden_columns->count].start_col = start_col;
    hidden_columns->ranges[hidden_columns->count].end_col = end_col;
    hidden_columns->count++;
}

static bool is_hidden_column(const HiddenColumns* hidden_columns, int col) {
    for (int i = 0; i < hidden_columns->count; i++) {
        if (col >= hidden_columns->ranges[i].start_col && col <= hidden_columns->ranges[i].end_col) {
            return true;
        }
    }
    return false;
}

static void init_merge_regions(MergeRegions* merge_regions) {
    merge_regions->regions = NULL;
    merge_regions->count = 0;
    merge_regions->capacity = 0;
}

static void free_merge_regions(MergeRegions* merge_regions) {
    free(merge_regions->regions);
    init_merge_regions(merge_regions);
}

static void add_merge_region(MergeRegions* merge_regions,
                             int start_row, int end_row, int start_col, int end_col) {
    if (start_row > end_row || start_col > end_col) {
        return;
    }

    if (merge_regions->count >= merge_regions->capacity) {
        merge_regions->capacity = merge_regions->capacity ? merge_regions->capacity * 2 : 8;
        merge_regions->regions = realloc(merge_regions->regions,
                                         sizeof(MergeRegion) * merge_regions->capacity);
    }

    MergeRegion* region = &merge_regions->regions[merge_regions->count++];
    region->start_row = start_row;
    region->end_row = end_row;
    region->start_col = start_col;
    region->end_col = end_col;
    region->value[0] = '\0';
    region->value_set = false;
}

static void init_row_buffer(RowBuffer* row_buffer) {
    row_buffer->cells = NULL;
    row_buffer->count = 0;
    row_buffer->capacity = 0;
}

static void clear_row_buffer(RowBuffer* row_buffer) {
    for (int i = 0; i < row_buffer->count; i++) {
        free(row_buffer->cells[i].value);
        row_buffer->cells[i].value = NULL;
    }
    row_buffer->count = 0;
}

static void free_row_buffer(RowBuffer* row_buffer) {
    clear_row_buffer(row_buffer);
    free(row_buffer->cells);
    init_row_buffer(row_buffer);
}

static void ensure_row_buffer_capacity(RowBuffer* row_buffer, int needed) {
    if (needed <= row_buffer->capacity) {
        return;
    }

    int new_capacity = row_buffer->capacity ? row_buffer->capacity * 2 : 16;
    while (new_capacity < needed) {
        new_capacity *= 2;
    }

    row_buffer->cells = realloc(row_buffer->cells, sizeof(RowCell) * new_capacity);
    row_buffer->capacity = new_capacity;
}

static int find_row_cell_index(const RowBuffer* row_buffer, int col) {
    for (int i = 0; i < row_buffer->count; i++) {
        if (row_buffer->cells[i].col == col) {
            return i;
        }
    }
    return -1;
}

static void add_row_cell(RowBuffer* row_buffer, int col, const char* value) {
    int existing_index = find_row_cell_index(row_buffer, col);
    if (existing_index >= 0) {
        free(row_buffer->cells[existing_index].value);
        row_buffer->cells[existing_index].value = strdup(value);
        return;
    }

    ensure_row_buffer_capacity(row_buffer, row_buffer->count + 1);
    row_buffer->cells[row_buffer->count].col = col;
    row_buffer->cells[row_buffer->count].value = strdup(value);
    row_buffer->count++;
}

static void add_row_cell_if_missing(RowBuffer* row_buffer, int col, const char* value) {
    if (find_row_cell_index(row_buffer, col) >= 0) {
        return;
    }
    add_row_cell(row_buffer, col, value);
}

static int compare_row_cells(const void* left, const void* right) {
    const RowCell* a = left;
    const RowCell* b = right;
    if (a->col < b->col) return -1;
    if (a->col > b->col) return 1;
    return 0;
}

static void sort_row_buffer(RowBuffer* row_buffer) {
    if (row_buffer->count > 1) {
        qsort(row_buffer->cells, row_buffer->count, sizeof(RowCell), compare_row_cells);
    }
}

static void add_custom_num_format(Styles* styles, int num_fmt_id, const char* format_code) {
    if (styles->custom_count >= styles->custom_capacity) {
        styles->custom_capacity = styles->custom_capacity ? styles->custom_capacity * 2 : 16;
        styles->custom_formats = realloc(styles->custom_formats,
                                         sizeof(CustomNumFormat) * styles->custom_capacity);
    }

    styles->custom_formats[styles->custom_count].num_fmt_id = num_fmt_id;
    styles->custom_formats[styles->custom_count].format_code = strdup(format_code);
    styles->custom_count++;
}

static const char* find_custom_num_format(const Styles* styles, int num_fmt_id) {
    for (int i = 0; i < styles->custom_count; i++) {
        if (styles->custom_formats[i].num_fmt_id == num_fmt_id) {
            return styles->custom_formats[i].format_code;
        }
    }
    return NULL;
}

static void add_xf_format(Styles* styles, CellFormat format) {
    if (styles->xf_count >= styles->xf_capacity) {
        styles->xf_capacity = styles->xf_capacity ? styles->xf_capacity * 2 : 32;
        styles->xf_formats = realloc(styles->xf_formats, sizeof(CellFormat) * styles->xf_capacity);
    }

    styles->xf_formats[styles->xf_count++] = format;
}

static bool workbook_uses_1904_date_system(const char* xml_data) {
    const char* workbook_pr = strstr(xml_data, "<workbookPr");
    char date_attr[16];
    if (!workbook_pr) {
        return false;
    }

    if (find_attribute(workbook_pr, "date1904=", 9, date_attr, sizeof(date_attr)) < 0) {
        return false;
    }

    return date_attr[0] == '1' || date_attr[0] == 't' || date_attr[0] == 'T';
}

static const char* builtin_number_format_code(int num_fmt_id) {
    switch (num_fmt_id) {
        case 1: return "0";
        case 2: return "0.00";
        case 3: return "#,##0";
        case 4: return "#,##0.00";
        case 9: return "0%";
        case 10: return "0.00%";
        case 11: return "0.00E+00";
        case 14: return "mm-dd-yy";
        case 15: return "d-mmm-yy";
        case 16: return "d-mmm";
        case 17: return "mmm-yy";
        case 18: return "h:mm AM/PM";
        case 19: return "h:mm:ss AM/PM";
        case 20: return "h:mm";
        case 21: return "h:mm:ss";
        case 22: return "m/d/yy h:mm";
        case 45: return "mm:ss";
        case 46: return "[h]:mm:ss";
        case 47: return "mmss.0";
        case 48: return "##0.0E+0";
        case 49: return "@";
        default: return NULL;
    }
}

static void sanitize_number_format_code(const char* format_code, char* out, int max_len, bool* elapsed_time) {
    int out_pos = 0;
    *elapsed_time = false;

    for (int i = 0; format_code[i] != '\0' && out_pos < max_len - 1; i++) {
        char c = format_code[i];

        if (c == ';') {
            break;
        }

        if (c == '"') {
            i++;
            while (format_code[i] && format_code[i] != '"') {
                i++;
            }
            continue;
        }

        if (c == '\\' || c == '_' || c == '*') {
            if (format_code[i + 1] != '\0') {
                i++;
            }
            continue;
        }

        if (c == '[') {
            int j = i + 1;
            while (format_code[j] && format_code[j] != ']') {
                j++;
            }

            if (!format_code[j]) {
                break;
            }

            if (j == i + 2 || j == i + 3) {
                char token0 = (char)tolower((unsigned char)format_code[i + 1]);
                if (token0 == 'h' || token0 == 'm' || token0 == 's') {
                    *elapsed_time = true;
                    while (++i < j && out_pos < max_len - 1) {
                        out[out_pos++] = (char)tolower((unsigned char)format_code[i]);
                    }
                    i = j;
                    continue;
                }
            }

            i = j;
            continue;
        }

        out[out_pos++] = (char)tolower((unsigned char)c);
    }

    out[out_pos] = '\0';
}

static int count_format_decimals(const char* sanitized_code) {
    const char* dot = strchr(sanitized_code, '.');
    if (!dot) {
        return 0;
    }

    int decimals = 0;
    for (const char* pos = dot + 1; *pos; pos++) {
        if (*pos == '0' || *pos == '#') {
            decimals++;
            continue;
        }
        if (*pos == '%' || *pos == 'e') {
            break;
        }
        if (!isdigit((unsigned char)*pos) && *pos != ',' && *pos != '?') {
            break;
        }
    }
    return decimals;
}

static bool is_zero_pad_format(const char* sanitized_code, int* zero_width) {
    int width = 0;
    if (sanitized_code[0] == '\0') {
        return false;
    }

    for (const char* pos = sanitized_code; *pos; pos++) {
        if (*pos != '0') {
            return false;
        }
        width++;
    }

    *zero_width = width;
    return width > 0;
}

static bool format_has_numeric_placeholders(const char* sanitized_code) {
    return strpbrk(sanitized_code, "0#?") != NULL;
}

static CellFormat analyze_number_format(int num_fmt_id, const char* format_code) {
    CellFormat format = make_raw_cell_format();
    bool elapsed_time = false;
    char sanitized[256];
    int zero_width = 0;

    if (!format_code || format_code[0] == '\0') {
        return format;
    }

    sanitize_number_format_code(format_code, sanitized, sizeof(sanitized), &elapsed_time);

    if (sanitized[0] == '@' && sanitized[1] == '\0') {
        return format;
    }

    bool has_am_pm = strstr(sanitized, "am/pm") != NULL || strstr(sanitized, "a/p") != NULL;
    bool has_hour = strchr(sanitized, 'h') != NULL;
    bool has_second = strchr(sanitized, 's') != NULL;
    bool has_year = strchr(sanitized, 'y') != NULL;
    bool has_day = strchr(sanitized, 'd') != NULL;
    bool has_month = strchr(sanitized, 'm') != NULL;
    bool has_time = has_hour || has_second || has_am_pm;
    bool has_date = has_year || has_day || (has_month && !has_time);

    if (has_month && has_time) {
        has_time = true;
    }

    if (has_date || has_time || elapsed_time ||
        (num_fmt_id >= 27 && num_fmt_id <= 36) ||
        (num_fmt_id >= 50 && num_fmt_id <= 58)) {
        if (elapsed_time) {
            format.kind = FORMAT_KIND_DURATION;
            format.show_seconds = has_second || num_fmt_id == 46;
            return format;
        }

        if (has_date && has_time) {
            format.kind = FORMAT_KIND_DATETIME;
            format.show_seconds = has_second;
            return format;
        }

        if (has_time) {
            format.kind = FORMAT_KIND_TIME;
            format.show_seconds = has_second;
            return format;
        }

        format.kind = FORMAT_KIND_DATE;
        return format;
    }

    if (strstr(sanitized, "e+") || strstr(sanitized, "e-")) {
        format.kind = FORMAT_KIND_SCIENTIFIC;
        format.decimals = count_format_decimals(sanitized);
        return format;
    }

    if (strchr(sanitized, '%')) {
        format.kind = FORMAT_KIND_PERCENT;
        format.decimals = count_format_decimals(sanitized);
        return format;
    }

    if (is_zero_pad_format(sanitized, &zero_width)) {
        format.kind = FORMAT_KIND_ZERO_PAD;
        format.zero_width = zero_width;
        return format;
    }

    if (format_has_numeric_placeholders(sanitized)) {
        format.kind = FORMAT_KIND_NUMBER;
        format.decimals = count_format_decimals(sanitized);
        format.use_thousands = strchr(sanitized, ',') != NULL;
    }

    return format;
}

static CellFormat resolve_cell_format(const Styles* styles, int num_fmt_id) {
    const char* format_code = find_custom_num_format(styles, num_fmt_id);
    if (!format_code) {
        format_code = builtin_number_format_code(num_fmt_id);
    }
    return analyze_number_format(num_fmt_id, format_code);
}

void parse_styles_xml(const char* xml_data, Styles* styles) {
    const char* xml_end = xml_data + strlen(xml_data);
    const char* numfmts_start = find_next_start_tag_local(xml_data, xml_end, "numFmts");
    const char* cellxfs_start = find_next_start_tag_local(xml_data, xml_end, "cellXfs");
    char tag_buffer[512];
    char id_attr[32];
    char code_attr[256];

    if (numfmts_start) {
        const char* numfmts_tag_end = find_tag_end_in_range(numfmts_start, xml_end);
        const char* numfmts_end = numfmts_tag_end ? find_next_end_tag_local(numfmts_tag_end + 1, xml_end, "numFmts") : NULL;
        const char* pos = numfmts_tag_end ? numfmts_tag_end + 1 : NULL;

        while (pos && numfmts_end && (pos = find_next_start_tag_local(pos, numfmts_end, "numFmt")) != NULL) {
            const char* tag_end = find_tag_end_in_range(pos, numfmts_end);
            if (!tag_end) {
                break;
            }

            copy_tag_range(pos, tag_end, tag_buffer, sizeof(tag_buffer));
            if (FIND_ATTR_NUMFMT_ID(tag_buffer, id_attr, sizeof(id_attr)) >= 0 &&
                FIND_ATTR_FORMAT_CODE(tag_buffer, code_attr, sizeof(code_attr)) >= 0) {
                unescape_xml_entities(code_attr);
                add_custom_num_format(styles, atoi(id_attr), code_attr);
            }

            pos = tag_end + 1;
        }
    }

    if (!cellxfs_start) {
        return;
    }

    const char* cellxfs_tag_end = find_tag_end_in_range(cellxfs_start, xml_end);
    const char* cellxfs_end = cellxfs_tag_end ? find_next_end_tag_local(cellxfs_tag_end + 1, xml_end, "cellXfs") : NULL;
    const char* pos = cellxfs_tag_end ? cellxfs_tag_end + 1 : NULL;

    while (pos && cellxfs_end && (pos = find_next_start_tag_local(pos, cellxfs_end, "xf")) != NULL) {
        const char* tag_end = find_tag_end_in_range(pos, cellxfs_end);
        if (!tag_end) {
            break;
        }

        copy_tag_range(pos, tag_end, tag_buffer, sizeof(tag_buffer));
        int num_fmt_id = 0;
        if (FIND_ATTR_NUMFMT_ID(tag_buffer, id_attr, sizeof(id_attr)) >= 0) {
            num_fmt_id = atoi(id_attr);
        }

        add_xf_format(styles, resolve_cell_format(styles, num_fmt_id));
        pos = tag_end + 1;
    }
}

static int64_t floor_double_to_i64(double value) {
    int64_t truncated = (int64_t)value;
    if ((double)truncated > value) {
        truncated--;
    }
    return truncated;
}

static int64_t days_from_civil(int year, unsigned month, unsigned day) {
    year -= month <= 2;
    const int era = (year >= 0 ? year : year - 399) / 400;
    const unsigned yoe = (unsigned)(year - era * 400);
    const unsigned doy = (153 * (month + (month > 2 ? -3 : 9)) + 2) / 5 + day - 1;
    const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097 + (int)doe - 719468;
}

static void civil_from_days(int64_t z, int* year, unsigned* month, unsigned* day) {
    z += 719468;
    const int era = (z >= 0 ? z : z - 146096) / 146097;
    const unsigned doe = (unsigned)(z - era * 146097);
    const unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    const int y = (int)yoe + era * 400;
    const unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    const unsigned mp = (5 * doy + 2) / 153;

    *day = doy - (153 * mp + 2) / 5 + 1;
    *month = mp + (mp < 10 ? 3 : -9);
    *year = y + (*month <= 2);
}

static bool excel_serial_to_datetime(double serial, bool date_1904,
                                     int* year, int* month, int* day,
                                     int* hour, int* minute, int* second) {
    int64_t days = floor_double_to_i64(serial);
    double fractional = serial - (double)days;

    if (fractional < 0.0) {
        fractional += 1.0;
        days--;
    }

    int total_seconds = (int)(fractional * 86400.0 + 0.5);
    if (total_seconds >= 86400) {
        total_seconds -= 86400;
        days++;
    }

    if (!date_1904 && days == 60) {
        *year = 1900;
        *month = 2;
        *day = 29;
    } else {
        int64_t adjusted_days = days;
        int64_t base_days = date_1904 ? days_from_civil(1904, 1, 1) : days_from_civil(1899, 12, 31);

        if (!date_1904 && days > 60) {
            adjusted_days--;
        }

        unsigned out_month;
        unsigned out_day;
        civil_from_days(base_days + adjusted_days, year, &out_month, &out_day);
        *month = (int)out_month;
        *day = (int)out_day;
    }

    *hour = total_seconds / 3600;
    *minute = (total_seconds % 3600) / 60;
    *second = total_seconds % 60;
    return true;
}

static void append_thousands_separators(const char* input, char* output, int max_len) {
    const char* digits = input;
    char sign = '\0';
    int out_pos = 0;

    if (*digits == '-' || *digits == '+') {
        sign = *digits++;
        if (out_pos < max_len - 1) {
            output[out_pos++] = sign;
        }
    }

    const char* dot = strchr(digits, '.');
    int int_len = dot ? (int)(dot - digits) : (int)strlen(digits);

    for (int i = 0; i < int_len && out_pos < max_len - 1; i++) {
        if (i > 0 && ((int_len - i) % 3 == 0)) {
            output[out_pos++] = ',';
        }
        output[out_pos++] = digits[i];
    }

    if (dot) {
        for (const char* pos = dot; *pos && out_pos < max_len - 1; pos++) {
            output[out_pos++] = *pos;
        }
    }

    output[out_pos] = '\0';
}

static void format_fixed_number(double value, int decimals, bool use_thousands, char* output, int max_len) {
    char buffer[128];
    snprintf(buffer, sizeof(buffer), "%.*f", decimals, value);

    if (use_thousands) {
        append_thousands_separators(buffer, output, max_len);
    } else {
        strncpy(output, buffer, max_len - 1);
        output[max_len - 1] = '\0';
    }
}

static void format_zero_padded_number(double value, int zero_width, char* output, int max_len) {
    char rounded[128];
    const char* digits = rounded;
    bool negative = false;
    int out_pos = 0;

    snprintf(rounded, sizeof(rounded), "%.0f", value);
    if (rounded[0] == '-') {
        negative = true;
        digits++;
    }

    if (negative && out_pos < max_len - 1) {
        output[out_pos++] = '-';
    }

    int digit_len = (int)strlen(digits);
    for (int i = digit_len; i < zero_width && out_pos < max_len - 1; i++) {
        output[out_pos++] = '0';
    }

    for (int i = 0; digits[i] && out_pos < max_len - 1; i++) {
        output[out_pos++] = digits[i];
    }

    output[out_pos] = '\0';
}

static void format_duration_value(double serial, bool show_seconds, char* output, int max_len) {
    int64_t total_seconds = floor_double_to_i64(serial * 86400.0 + 0.5);
    if (total_seconds < 0) {
        snprintf(output, max_len, "%.10g", serial);
        return;
    }

    int64_t hours = total_seconds / 3600;
    int minutes = (int)((total_seconds % 3600) / 60);
    int seconds = (int)(total_seconds % 60);

    if (show_seconds) {
        snprintf(output, max_len, "%lld:%02d:%02d", (long long)hours, minutes, seconds);
    } else {
        snprintf(output, max_len, "%lld:%02d", (long long)hours, minutes);
    }
}

static bool format_numeric_value(double value, CellFormat format, bool date_1904, char* output, int max_len) {
    int year, month, day, hour, minute, second;

    switch (format.kind) {
        case FORMAT_KIND_RAW:
            return false;

        case FORMAT_KIND_NUMBER:
            format_fixed_number(value, format.decimals, format.use_thousands, output, max_len);
            return true;

        case FORMAT_KIND_PERCENT:
            format_fixed_number(value * 100.0, format.decimals, false, output, max_len - 1);
            strncat(output, "%", (size_t)max_len - strlen(output) - 1);
            return true;

        case FORMAT_KIND_SCIENTIFIC:
            snprintf(output, max_len, "%.*E", format.decimals, value);
            return true;

        case FORMAT_KIND_ZERO_PAD:
            format_zero_padded_number(value, format.zero_width, output, max_len);
            return true;

        case FORMAT_KIND_DATE:
        case FORMAT_KIND_TIME:
        case FORMAT_KIND_DATETIME:
            excel_serial_to_datetime(value, date_1904, &year, &month, &day, &hour, &minute, &second);
            if (format.kind == FORMAT_KIND_DATE) {
                snprintf(output, max_len, "%04d-%02d-%02d", year, month, day);
            } else if (format.kind == FORMAT_KIND_TIME) {
                if (format.show_seconds) {
                    snprintf(output, max_len, "%02d:%02d:%02d", hour, minute, second);
                } else {
                    snprintf(output, max_len, "%02d:%02d", hour, minute);
                }
            } else if (format.show_seconds) {
                snprintf(output, max_len, "%04d-%02d-%02d %02d:%02d:%02d",
                         year, month, day, hour, minute, second);
            } else {
                snprintf(output, max_len, "%04d-%02d-%02d %02d:%02d",
                         year, month, day, hour, minute);
            }
            return true;

        case FORMAT_KIND_DURATION:
            format_duration_value(value, format.show_seconds, output, max_len);
            return true;
    }

    return false;
}

static void format_generic_scalar(const char* raw_value, bool has_type_attr, const char* type_attr,
                                  int style_index, const Styles* styles, bool formatted_output,
                                  char* output, int max_len) {
    if (formatted_output && has_type_attr && strcmp(type_attr, "b") == 0) {
        if (strcmp(raw_value, "1") == 0) {
            escape_tsv_value("TRUE", output, max_len);
        } else if (strcmp(raw_value, "0") == 0) {
            escape_tsv_value("FALSE", output, max_len);
        } else {
            escape_tsv_value(raw_value, output, max_len);
        }
        return;
    }

    if (!formatted_output || !styles || !styles->enabled ||
        (has_type_attr && (strcmp(type_attr, "e") == 0 || strcmp(type_attr, "str") == 0 || strcmp(type_attr, "d") == 0))) {
        escape_tsv_value(raw_value, output, max_len);
        return;
    }

    char* end_ptr = NULL;
    double numeric_value = strtod(raw_value, &end_ptr);
    if (!end_ptr || *end_ptr != '\0') {
        escape_tsv_value(raw_value, output, max_len);
        return;
    }

    CellFormat format = make_raw_cell_format();
    if (style_index >= 0 && style_index < styles->xf_count) {
        format = styles->xf_formats[style_index];
    }

    if (!format_numeric_value(numeric_value, format, styles->date_1904, output, max_len)) {
        escape_tsv_value(raw_value, output, max_len);
    }
}

// 공유 문자열 초기화
void init_shared_strings(SharedStrings* ss) {
    ss->capacity = 1024;  // 작게 시작, 필요에 따라 확장
    ss->strings = malloc(sizeof(char*) * ss->capacity);
    ss->count = 0;
}

static void reserve_shared_strings(SharedStrings* ss, int needed) {
    if (needed <= ss->capacity) {
        return;
    }

    ss->capacity = needed;
    ss->strings = realloc(ss->strings, sizeof(char*) * ss->capacity);
}

// XML 엔티티를 제자리에서 언이스케이프 - 최적화 버전
void unescape_xml_entities(char* str) {
    char* src = str;
    char* dst = str;

    while (*src) {
        if (*src == '&') {
            // 비교 횟수를 줄이기 위해 두 번째 문자를 먼저 확인
            switch (src[1]) {
                case 'l':  // &lt;
                    if (src[2] == 't' && src[3] == ';') {
                        *dst++ = '<';
                        src += 4;
                    } else {
                        *dst++ = *src++;
                    }
                    break;
                case 'g':  // &gt;
                    if (src[2] == 't' && src[3] == ';') {
                        *dst++ = '>';
                        src += 4;
                    } else {
                        *dst++ = *src++;
                    }
                    break;
                case 'a':  // &amp; 또는 &apos;
                    if (src[2] == 'm' && src[3] == 'p' && src[4] == ';') {
                        *dst++ = '&';
                        src += 5;
                    } else if (src[2] == 'p' && src[3] == 'o' && src[4] == 's' && src[5] == ';') {
                        *dst++ = '\'';
                        src += 6;
                    } else {
                        *dst++ = *src++;
                    }
                    break;
                case 'q':  // &quot;
                    if (src[2] == 'u' && src[3] == 'o' && src[4] == 't' && src[5] == ';') {
                        *dst++ = '"';
                        src += 6;
                    } else {
                        *dst++ = *src++;
                    }
                    break;
                default:
                    *dst++ = *src++;
                    break;
            }
        } else {
            *dst++ = *src++;
        }
    }
    *dst = '\0';
}

// 공유 문자열에 문자열 추가
void add_shared_string(SharedStrings* ss, const char* str) {
    if (ss->count >= ss->capacity) {
        ss->capacity *= 2;
        ss->strings = realloc(ss->strings, sizeof(char*) * ss->capacity);
    }

    ss->strings[ss->count] = strdup(str);

    // XML 엔티티 언이스케이프
    if (strchr(ss->strings[ss->count], '&')) {
        unescape_xml_entities(ss->strings[ss->count]);
    }

    ss->count++;

    // 디버그: 처음 30개 공유 문자열 출력
#ifdef DEBUG
    printf("DEBUG: Shared string [%d] = '%s'\n", ss->count - 1, ss->strings[ss->count - 1]);
#endif
}

// 텍스트 내용을 추출하고 모든 태그를 건너뛰며 공유 문자열 XML 파싱
void parse_shared_strings(const char* xml_data, SharedStrings* ss) {
    const char* pos = xml_data;
    const char* sst_start = strstr(xml_data, "<sst");
    char count_attr[32];

    if (sst_start) {
        const char* sst_end = strchr(sst_start, '>');
        if (sst_end) {
            if (find_attribute_in_range(sst_start, sst_end + 1, "uniqueCount=", 12,
                                        count_attr, sizeof(count_attr)) >= 0) {
                reserve_shared_strings(ss, atoi(count_attr));
            } else if (find_attribute_in_range(sst_start, sst_end + 1, "count=", 6,
                                               count_attr, sizeof(count_attr)) >= 0) {
                reserve_shared_strings(ss, atoi(count_attr));
            }
        }
    }

    // 각 <si> (공유 문자열 항목) 요소 찾기
    while ((pos = strstr(pos, "<si")) != NULL) {
        // 자기 닫힘 태그 <si/> 확인
        const char* tag_end = strchr(pos, '>');
        if (!tag_end) break;

        if (*(tag_end - 1) == '/') {
            // 자기 닫힘 태그 <si/> - 빈 문자열을 나타냄
            add_shared_string(ss, "");
            pos = tag_end + 1;
            continue;
        }

        // 일반 <si>...</si> 태그
        if (*(pos + 3) != '>') {
            // 정확히 "<si>"가 아니면 건너뛰기
            pos++;
            continue;
        }

        pos += 4; // <si> 건너뛰기
        const char* si_end = strstr(pos, "</si>");
        if (!si_end) break;

        // 모든 태그를 건너뛰고 텍스트 내용 추출
        char text_buffer[MAX_CELL_VALUE] = "";
        int buffer_pos = 0;
        const char* current = pos;
        int inside_tag = 0;

        while (current < si_end && buffer_pos < MAX_CELL_VALUE - 1) {
            if (*current == '<') {
                inside_tag = 1;  // 태그 시작
            } else if (*current == '>') {
                inside_tag = 0;  // 태그 끝
            } else if (!inside_tag) {
                // 태그 안이 아니면 버퍼에 문자 추가
                text_buffer[buffer_pos++] = *current;
            }
            current++;
        }

        text_buffer[buffer_pos] = '\0';

        // 공유 문자열에 추가 (올바른 인덱싱을 유지하기 위해 빈 문자열 포함)
        add_shared_string(ss, text_buffer);

        pos = si_end + 5; // </si>를 지나서 이동
    }
}

void parse_shared_strings_generic(const char* xml_data, SharedStrings* ss) {
    const char* xml_end = xml_data + strlen(xml_data);
    const char* pos = xml_data;

    while ((pos = find_next_start_tag_local(pos, xml_end, "si")) != NULL) {
        const char* si_tag_end = find_tag_end_in_range(pos, xml_end);
        if (!si_tag_end) {
            break;
        }

        if (is_self_closing_tag(pos, si_tag_end)) {
            add_shared_string(ss, "");
            pos = si_tag_end + 1;
            continue;
        }

        const char* si_close = find_next_end_tag_local(si_tag_end + 1, xml_end, "si");
        if (!si_close) {
            break;
        }

        char text_buffer[MAX_CELL_VALUE] = "";
        extract_text_runs(si_tag_end + 1, si_close, text_buffer, sizeof(text_buffer));
        add_shared_string(ss, text_buffer);

        const char* si_close_end = find_tag_end_in_range(si_close, xml_end);
        if (!si_close_end) {
            break;
        }
        pos = si_close_end + 1;
    }
}

// workbook.xml을 파싱하여 시트 정보 가져오기
void parse_workbook(const char* xml_data, Workbook* wb, bool export_all_sheets) {
    wb->sheet_count = 0;
    const char* pos = xml_data;
    char name_attr[MAX_SHEET_NAME];
    char sheet_id_attr[32];
    char rel_id_attr[MAX_REL_ID];
    char state_attr[MAX_SHEET_STATE];

    while ((pos = strstr(pos, "<sheet ")) != NULL) {
        // 시트 이름 추출
        if (FIND_ATTR_NAME(pos, name_attr, MAX_SHEET_NAME) < 0) {
            pos++;
            continue;
        }

        if (FIND_ATTR_REL_ID(pos, rel_id_attr, sizeof(rel_id_attr)) < 0) {
            pos++;
            continue;
        }

        // 시트 ID 추출
        int sheet_id = wb->sheet_count + 1;
        if (FIND_ATTR_SHEET_ID(pos, sheet_id_attr, sizeof(sheet_id_attr)) >= 0) {
            sheet_id = atoi(sheet_id_attr);
        }

        // 유효하지 않은 문자가 있는 시트 건너뛰기 (A-Z, a-z, 0-9, -, _, *만 허용)
        if (!export_all_sheets && !is_valid_name(name_attr)) {
            printf("Skipping sheet: '%s' (contains invalid characters - only A-Z, a-z, 0-9, -, _, * allowed)\n", name_attr);
            pos++;
            continue;
        }

        // 시트 정보 저장
        ensure_workbook_capacity(wb, wb->sheet_count + 1);
        strncpy(wb->sheets[wb->sheet_count].name, name_attr, MAX_SHEET_NAME - 1);
        wb->sheets[wb->sheet_count].name[MAX_SHEET_NAME - 1] = '\0';
        strncpy(wb->sheets[wb->sheet_count].rel_id, rel_id_attr, MAX_REL_ID - 1);
        wb->sheets[wb->sheet_count].rel_id[MAX_REL_ID - 1] = '\0';
        wb->sheets[wb->sheet_count].filename[0] = '\0';
        wb->sheets[wb->sheet_count].output_name[0] = '\0';
        if (FIND_ATTR_STATE(pos, state_attr, sizeof(state_attr)) >= 0) {
            strncpy(wb->sheets[wb->sheet_count].state, state_attr, MAX_SHEET_STATE - 1);
            wb->sheets[wb->sheet_count].state[MAX_SHEET_STATE - 1] = '\0';
        } else {
            strncpy(wb->sheets[wb->sheet_count].state, "visible", MAX_SHEET_STATE - 1);
            wb->sheets[wb->sheet_count].state[MAX_SHEET_STATE - 1] = '\0';
        }
        wb->sheets[wb->sheet_count].sheet_id = sheet_id;
        wb->sheets[wb->sheet_count].hidden = strcmp(wb->sheets[wb->sheet_count].state, "visible") != 0;
        wb->sheets[wb->sheet_count].selected = true;
        wb->sheets[wb->sheet_count].processed = false;
        wb->sheets[wb->sheet_count].truncated = false;
        wb->sheets[wb->sheet_count].emitted_rows = 0;
        wb->sheets[wb->sheet_count].emitted_cols = 0;
        wb->sheets[wb->sheet_count].approx_rows = -1;
        wb->sheets[wb->sheet_count].approx_cols = -1;
        wb->sheets[wb->sheet_count].warning_count = 0;

        wb->sheet_count++;
        pos++;
    }
}

void resolve_sheet_filenames(const char* xml_data, Workbook* wb) {
    const char* pos = xml_data;
    char id_attr[MAX_REL_ID];
    char target_attr[MAX_SHEET_NAME];
    char type_attr[128];

    while ((pos = strstr(pos, "<Relationship ")) != NULL) {
        if (FIND_ATTR_ID(pos, id_attr, sizeof(id_attr)) < 0 ||
            FIND_ATTR_TARGET(pos, target_attr, sizeof(target_attr)) < 0 ||
            FIND_ATTR_TYPE(pos, type_attr, sizeof(type_attr)) < 0) {
            pos++;
            continue;
        }

        if (!strstr(type_attr, "/worksheet")) {
            pos++;
            continue;
        }

        for (int i = 0; i < wb->sheet_count; i++) {
            if (strcmp(wb->sheets[i].rel_id, id_attr) != 0) {
                continue;
            }

            if (strncmp(target_attr, "/xl/", 4) == 0) {
                strncpy(wb->sheets[i].filename, target_attr + 1, MAX_SHEET_NAME - 1);
                wb->sheets[i].filename[MAX_SHEET_NAME - 1] = '\0';
            } else if (strncmp(target_attr, "xl/", 3) == 0) {
                strncpy(wb->sheets[i].filename, target_attr, MAX_SHEET_NAME - 1);
                wb->sheets[i].filename[MAX_SHEET_NAME - 1] = '\0';
            } else {
                snprintf(wb->sheets[i].filename, MAX_SHEET_NAME, "xl/%s", target_attr);
            }
            break;
        }

        pos++;
    }
}

// Excel 열 참조를 숫자로 변환 (A=0, B=1, 등)
static inline int col_ref_to_num(const char* ref) {
    int col = 0;
    for (int i = 0; ref[i] && isalpha(ref[i]); i++) {
        col = col * 26 + (toupper(ref[i]) - 'A' + 1);
    }
    return col - 1;
}

// 셀 참조에서 행 번호 추출
static inline int extract_row_num(const char* ref) {
    while (*ref && isalpha(*ref)) ref++;
    return atoi(ref) - 1;
}

static inline bool parse_cell_ref_parts(const char* ref, int* row, int* col) {
    int parsed_col = 0;
    int parsed_row = 0;
    const unsigned char* pos = (const unsigned char*)ref;

    if (!pos || !isalpha(*pos)) {
        return false;
    }

    while (*pos && isalpha(*pos)) {
        unsigned char c = *pos++;
        if (c >= 'a' && c <= 'z') {
            c -= 'a' - 'A';
        }
        parsed_col = parsed_col * 26 + (c - 'A' + 1);
    }

    if (!isdigit(*pos)) {
        return false;
    }

    while (*pos && isdigit(*pos)) {
        parsed_row = parsed_row * 10 + (*pos++ - '0');
    }

    if (parsed_col <= 0 || parsed_row <= 0) {
        return false;
    }

    *col = parsed_col - 1;
    *row = parsed_row - 1;
    return true;
}

static void parse_hidden_columns_xml(const char* xml_data, HiddenColumns* hidden_columns) {
    const char* xml_end = xml_data + strlen(xml_data);
    const char* pos = xml_data;
    char tag_buffer[256];
    char min_attr[32];
    char max_attr[32];
    char hidden_attr[16];

    while ((pos = find_next_start_tag_local(pos, xml_end, "col")) != NULL) {
        const char* tag_end = find_tag_end_in_range(pos, xml_end);
        if (!tag_end) {
            break;
        }

        copy_tag_range(pos, tag_end, tag_buffer, sizeof(tag_buffer));
        if (FIND_ATTR_HIDDEN(tag_buffer, hidden_attr, sizeof(hidden_attr)) >= 0 &&
            attr_is_true(hidden_attr) &&
            FIND_ATTR_MIN(tag_buffer, min_attr, sizeof(min_attr)) >= 0 &&
            FIND_ATTR_MAX(tag_buffer, max_attr, sizeof(max_attr)) >= 0) {
            add_hidden_column_range(hidden_columns, atoi(min_attr) - 1, atoi(max_attr) - 1);
        }

        pos = tag_end + 1;
    }
}

static bool parse_cell_ref(const char* ref, int* row, int* col) {
    if (!ref || !isalpha((unsigned char)ref[0])) {
        return false;
    }

    *col = col_ref_to_num(ref);
    *row = extract_row_num(ref);
    return *row >= 0 && *col >= 0;
}

static void parse_merge_regions_xml(const char* xml_data, MergeRegions* merge_regions) {
    const char* xml_end = xml_data + strlen(xml_data);
    const char* pos = xml_data;
    char tag_buffer[256];
    char ref_attr[64];

    while ((pos = find_next_start_tag_local(pos, xml_end, "mergeCell")) != NULL) {
        const char* tag_end = find_tag_end_in_range(pos, xml_end);
        if (!tag_end) {
            break;
        }

        copy_tag_range(pos, tag_end, tag_buffer, sizeof(tag_buffer));
        if (FIND_ATTR_REF(tag_buffer, ref_attr, sizeof(ref_attr)) >= 0) {
            char start_ref[32];
            char end_ref[32];
            const char* separator = strchr(ref_attr, ':');
            int start_row;
            int start_col;
            int end_row;
            int end_col;

            if (separator) {
                size_t start_len = (size_t)(separator - ref_attr);
                if (start_len >= sizeof(start_ref)) {
                    start_len = sizeof(start_ref) - 1;
                }
                memcpy(start_ref, ref_attr, start_len);
                start_ref[start_len] = '\0';
                strncpy(end_ref, separator + 1, sizeof(end_ref) - 1);
                end_ref[sizeof(end_ref) - 1] = '\0';
            } else {
                strncpy(start_ref, ref_attr, sizeof(start_ref) - 1);
                start_ref[sizeof(start_ref) - 1] = '\0';
                strncpy(end_ref, ref_attr, sizeof(end_ref) - 1);
                end_ref[sizeof(end_ref) - 1] = '\0';
            }

            if (parse_cell_ref(start_ref, &start_row, &start_col) &&
                parse_cell_ref(end_ref, &end_row, &end_col)) {
                add_merge_region(merge_regions, start_row, end_row, start_col, end_col);
            }
        }

        pos = tag_end + 1;
    }
}

static void register_merge_anchor_value(MergeRegions* merge_regions, int row, int col, const char* value) {
    for (int i = 0; i < merge_regions->count; i++) {
        MergeRegion* region = &merge_regions->regions[i];
        if (region->start_row == row && region->start_col == col) {
            strncpy(region->value, value, sizeof(region->value) - 1);
            region->value[sizeof(region->value) - 1] = '\0';
            region->value_set = true;
        }
    }
}

static void apply_merge_regions_to_row(MergeRegions* merge_regions, int row, RowBuffer* row_buffer) {
    for (int i = 0; i < merge_regions->count; i++) {
        MergeRegion* region = &merge_regions->regions[i];
        if (!region->value_set || row < region->start_row || row > region->end_row) {
            continue;
        }

        for (int col = region->start_col; col <= region->end_col; col++) {
            add_row_cell_if_missing(row_buffer, col, region->value);
        }
    }
}

static bool row_buffer_has_visible_cells(const RowBuffer* row_buffer, const HiddenColumns* hidden_columns) {
    for (int i = 0; i < row_buffer->count; i++) {
        if (!is_hidden_column(hidden_columns, row_buffer->cells[i].col)) {
            return true;
        }
    }
    return false;
}

static int find_last_visible_col_in_range(int start_col, int end_col, const HiddenColumns* hidden_columns) {
    for (int col = end_col; col >= start_col; col--) {
        if (!is_hidden_column(hidden_columns, col)) {
            return col;
        }
    }
    return -1;
}

static int scan_generic_visible_max_col(const char* xml_data, int start_row, bool skip_hidden,
                                        const HiddenColumns* hidden_columns, const MergeRegions* merge_regions) {
    const char* xml_end = xml_data + strlen(xml_data);
    const char* pos = xml_data;
    int max_visible_col = -1;

    while ((pos = find_next_start_tag_local(pos, xml_end, "row")) != NULL) {
        const char* row_tag_end = find_tag_end_in_range(pos, xml_end);
        const char* row_close;
        const char* row_close_end;
        char row_open_tag[256];
        char row_attr[32];
        char hidden_attr[16];
        int row = -1;
        bool row_hidden = false;

        if (!row_tag_end) {
            break;
        }

        copy_tag_range(pos, row_tag_end, row_open_tag, sizeof(row_open_tag));
        row_close = is_self_closing_tag(pos, row_tag_end) ? NULL : find_next_end_tag_local(row_tag_end + 1, xml_end, "row");
        row_close_end = row_close ? find_tag_end_in_range(row_close, xml_end) : row_tag_end;
        if (!row_close_end) {
            break;
        }

        if (find_attribute(row_open_tag, "r=", 2, row_attr, sizeof(row_attr)) >= 0) {
            row = atoi(row_attr) - 1;
        }

        if (row < start_row) {
            pos = row_close_end + 1;
            continue;
        }

        if (skip_hidden &&
            FIND_ATTR_HIDDEN(row_open_tag, hidden_attr, sizeof(hidden_attr)) >= 0 &&
            attr_is_true(hidden_attr)) {
            row_hidden = true;
        }

        if (!row_hidden && row_close) {
            const char* cell_pos = row_tag_end + 1;
            while ((cell_pos = find_next_start_tag_local(cell_pos, row_close, "c")) != NULL) {
                const char* cell_tag_end = find_tag_end_in_range(cell_pos, row_close);
                char cell_open_tag[512];
                char r_attr[32];
                int col;

                if (!cell_tag_end) {
                    break;
                }

                copy_tag_range(cell_pos, cell_tag_end, cell_open_tag, sizeof(cell_open_tag));
                if (FIND_ATTR_R(cell_open_tag, r_attr, sizeof(r_attr)) >= 0) {
                    col = col_ref_to_num(r_attr);
                    if (!is_hidden_column(hidden_columns, col) && col > max_visible_col) {
                        max_visible_col = col;
                    }
                }

                if (is_self_closing_tag(cell_pos, cell_tag_end)) {
                    cell_pos = cell_tag_end + 1;
                } else {
                    const char* cell_close = find_next_end_tag_local(cell_tag_end + 1, row_close, "c");
                    const char* cell_close_end = cell_close ? find_tag_end_in_range(cell_close, row_close) : NULL;
                    if (!cell_close_end) {
                        break;
                    }
                    cell_pos = cell_close_end + 1;
                }
            }
        }

        pos = row_close_end + 1;
    }

    for (int i = 0; i < merge_regions->count; i++) {
        const MergeRegion* region = &merge_regions->regions[i];
        int visible_col;

        if (region->end_row < start_row) {
            continue;
        }

        visible_col = find_last_visible_col_in_range(region->start_col, region->end_col, hidden_columns);
        if (visible_col > max_visible_col) {
            max_visible_col = visible_col;
        }
    }

    return max_visible_col;
}

static void emit_row_buffer(RowBuffer* row_buffer, const HiddenColumns* hidden_columns,
                            int target_max_col, bool allow_empty_row, Filter* output) {
    sort_row_buffer(row_buffer);

    int last_col = -1;
    for (int i = 0; i < row_buffer->count; i++) {
        int col = row_buffer->cells[i].col;
        if (is_hidden_column(hidden_columns, col)) {
            continue;
        }

        for (int gap_col = last_col + 1; gap_col < col; gap_col++) {
            if (!is_hidden_column(hidden_columns, gap_col)) {
                filter_push(output, "");
            }
        }

        filter_push(output, row_buffer->cells[i].value);
        last_col = col;
    }

    for (int gap_col = last_col + 1; gap_col <= target_max_col; gap_col++) {
        if (!is_hidden_column(hidden_columns, gap_col)) {
            filter_push(output, "");
            last_col = gap_col;
        }
    }

    if (last_col >= 0 || allow_empty_row) {
        filter_finish_line(output);
    }
}

static int max_merge_row_with_values(const MergeRegions* merge_regions) {
    int max_row = -1;
    for (int i = 0; i < merge_regions->count; i++) {
        if (merge_regions->regions[i].value_set && merge_regions->regions[i].end_row > max_row) {
            max_row = merge_regions->regions[i].end_row;
        }
    }
    return max_row;
}

static void emit_synthetic_merge_rows(int first_row, int last_row,
                                      const HiddenColumns* hidden_columns,
                                      MergeRegions* merge_regions,
                                      int target_max_col,
                                      RowBuffer* row_buffer,
                                      Filter* output) {
    for (int row = first_row; row < last_row; row++) {
        clear_row_buffer(row_buffer);
        apply_merge_regions_to_row(merge_regions, row, row_buffer);
        if (row_buffer_has_visible_cells(row_buffer, hidden_columns) || target_max_col >= 0) {
            emit_row_buffer(row_buffer, hidden_columns, target_max_col, target_max_col < 0, output);
        }
    }
}

// TSV 특수 문자 이스케이프
void escape_tsv_value(const char* input, char* output, int max_len) {
    int i = 0, j = 0;
    while (input[i] && j < max_len - 1) {
        if (input[i] == '\t') {
            output[j++] = ' ';  // 탭을 공백으로 치환
        } else if (input[i] == '\n' || input[i] == '\r') {
            output[j++] = ' ';  // 개행을 공백으로 치환
        } else {
            output[j++] = input[i];
        }
        i++;
    }
    output[j] = '\0';
}

static inline void escape_tsv_value_range(const char* start, const char* end,
                                          char* output, int max_len) {
    int j = 0;

    while (start < end && j < max_len - 1) {
        if (*start == '\t' || *start == '\n' || *start == '\r') {
            output[j++] = ' ';
        } else {
            output[j++] = *start;
        }
        start++;
    }

    output[j] = '\0';
}

static inline int parse_int_range(const char* start, const char* end) {
    int value = 0;

    while (start < end && isspace((unsigned char)*start)) {
        start++;
    }

    while (start < end && isdigit((unsigned char)*start)) {
        value = value * 10 + (*start - '0');
        start++;
    }

    return value;
}

// 시트 이름에서 안전한 파일명용 베이스 이름 생성
void create_safe_filename_base(const char* sheet_name, char* safe_name, int max_len) {
    int i = 0, j = 0;
    while (sheet_name[i] && j < max_len - 1) {
        char c = sheet_name[i];
        // 안전하지 않은 문자를 언더스코어로 치환
        if (c == '/' || c == '\\' || c == ':' ||
            c == '?' || c == '"' || c == '<' || c == '>' || c == '|') {
            safe_name[j++] = '_';
        } else if (c == ' ') {
            safe_name[j++] = '_';  // 공백을 언더스코어로 치환
        } else if (c == '*') {
            // 별표 문자 건너뛰기 (파일명에 추가하지 않음)
            // 아무것도 하지 않고 다음 문자로 이동
        } else {
            safe_name[j++] = c;
        }
        i++;
    }
    if (j == 0) {
        strncpy(safe_name, "sheet", max_len - 1);
        safe_name[max_len - 1] = '\0';
        return;
    }
    safe_name[j] = '\0';
}

// 시트 이름에서 안전한 파일명 생성
void create_safe_filename(const char* sheet_name, const char* extension, char* safe_name, int max_len) {
    char base_name[MAX_OUTPUT_FILENAME];
    create_safe_filename_base(sheet_name, base_name, sizeof(base_name));
    snprintf(safe_name, max_len, "%s.%s", base_name, extension);
}

void create_unique_output_filename(const char* sheet_name, int sheet_index,
                                   char used_names[][MAX_OUTPUT_FILENAME], int used_count,
                                   const char* extension, char* output_filename, int max_len) {
    char base_name[MAX_OUTPUT_FILENAME];
    create_safe_filename_base(sheet_name, base_name, sizeof(base_name));

    if (base_name[0] == '\0') {
        snprintf(base_name, sizeof(base_name), "sheet%d", sheet_index + 1);
    }

    snprintf(output_filename, max_len, "%s.%s", base_name, extension);

    bool is_duplicate = false;
    for (int i = 0; i < used_count; i++) {
        if (strcmp(used_names[i], output_filename) == 0) {
            is_duplicate = true;
            break;
        }
    }

    if (!is_duplicate) {
        return;
    }

    for (int suffix = 2; ; suffix++) {
        snprintf(output_filename, max_len, "%s__%d.%s", base_name, suffix, extension);

        is_duplicate = false;
        for (int i = 0; i < used_count; i++) {
            if (strcmp(used_names[i], output_filename) == 0) {
                is_duplicate = true;
                break;
            }
        }

        if (!is_duplicate) {
            return;
        }
    }
}

static void join_output_path(const char* output_dir, const char* filename,
                             char* output_path, int max_len) {
    size_t dir_len;

    if (!output_dir || output_dir[0] == '\0' || strcmp(output_dir, ".") == 0) {
        strncpy(output_path, filename, (size_t)max_len - 1);
        output_path[max_len - 1] = '\0';
        return;
    }

    dir_len = strlen(output_dir);
    if (dir_len > 0 && output_dir[dir_len - 1] == '/') {
        snprintf(output_path, (size_t)max_len, "%s%s", output_dir, filename);
    } else {
        snprintf(output_path, (size_t)max_len, "%s/%s", output_dir, filename);
    }
}

static bool output_name_exists(char used_names[][MAX_OUTPUT_FILENAME], int used_count, const char* output_filename) {
    for (int i = 0; i < used_count; i++) {
        if (strcmp(used_names[i], output_filename) == 0) {
            return true;
        }
    }
    return false;
}

static void parse_sheet_dimension_xml(const char* xml_data, int* row_count, int* col_count) {
    const char* pos = strstr(xml_data, "<dimension ");
    char ref_attr[64];
    char start_ref[32];
    char end_ref[32];
    const char* separator;
    int start_row;
    int start_col;
    int end_row;
    int end_col;

    *row_count = -1;
    *col_count = -1;

    if (!pos || FIND_ATTR_REF(pos, ref_attr, sizeof(ref_attr)) < 0) {
        return;
    }

    separator = strchr(ref_attr, ':');
    if (separator) {
        size_t start_len = (size_t)(separator - ref_attr);
        if (start_len >= sizeof(start_ref)) {
            start_len = sizeof(start_ref) - 1;
        }
        memcpy(start_ref, ref_attr, start_len);
        start_ref[start_len] = '\0';
        strncpy(end_ref, separator + 1, sizeof(end_ref) - 1);
        end_ref[sizeof(end_ref) - 1] = '\0';
    } else {
        strncpy(start_ref, ref_attr, sizeof(start_ref) - 1);
        start_ref[sizeof(start_ref) - 1] = '\0';
        strncpy(end_ref, ref_attr, sizeof(end_ref) - 1);
        end_ref[sizeof(end_ref) - 1] = '\0';
    }

    if (parse_cell_ref(start_ref, &start_row, &start_col) &&
        parse_cell_ref(end_ref, &end_row, &end_col)) {
        *row_count = end_row - start_row + 1;
        *col_count = end_col - start_col + 1;
    }
}

static bool sheet_matches_filters(const SheetInfo* sheet,
                                  char sheet_names[][MAX_SHEET_NAME], int sheet_name_count,
                                  regex_t* sheet_regex, bool has_sheet_regex) {
    if (sheet_name_count == 0 && !has_sheet_regex) {
        return true;
    }

    for (int i = 0; i < sheet_name_count; i++) {
        if (strcmp(sheet->name, sheet_names[i]) == 0) {
            return true;
        }
    }

    if (has_sheet_regex && regexec(sheet_regex, sheet->name, 0, NULL, 0) == 0) {
        return true;
    }

    return false;
}

static void json_write_string(FILE* fp, const char* value) {
    const unsigned char* pos = (const unsigned char*)(value ? value : "");
    fputc('"', fp);
    while (*pos) {
        switch (*pos) {
            case '\\':
                fputs("\\\\", fp);
                break;
            case '"':
                fputs("\\\"", fp);
                break;
            case '\b':
                fputs("\\b", fp);
                break;
            case '\f':
                fputs("\\f", fp);
                break;
            case '\n':
                fputs("\\n", fp);
                break;
            case '\r':
                fputs("\\r", fp);
                break;
            case '\t':
                fputs("\\t", fp);
                break;
            default:
                if (*pos < 0x20) {
                    fprintf(fp, "\\u%04x", *pos);
                } else {
                    fputc(*pos, fp);
                }
                break;
        }
        pos++;
    }
    fputc('"', fp);
}

static void json_write_sheet_warnings(FILE* fp, const SheetInfo* sheet) {
    fputc('[', fp);
    for (int i = 0; i < sheet->warning_count; i++) {
        if (i > 0) {
            fputc(',', fp);
        }
        fputs("{\"code\":", fp);
        json_write_string(fp, sheet->warnings[i].code);
        fputs(",\"message\":", fp);
        json_write_string(fp, sheet->warnings[i].message);
        fputc('}', fp);
    }
    fputc(']', fp);
}

static void write_manifest_json(FILE* fp, const char* input_file,
                                const RunOptions* options, const Workbook* workbook,
                                const GlobalWarnings* global_warnings) {
    int first = 1;
    int selected_count = 0;
    int processed_count = 0;

    for (int i = 0; i < workbook->sheet_count; i++) {
        if (workbook->sheets[i].selected) {
            selected_count++;
        }
        if (workbook->sheets[i].processed) {
            processed_count++;
        }
    }

    fputs("{\"tool_version\":", fp);
    json_write_string(fp, TOOL_VERSION);
    fputs(",\"input_file\":", fp);
    json_write_string(fp, input_file);
    fputs(",\"mode\":", fp);
    json_write_string(fp, options->game_db_fast_mode ? "game-db-fast" : "generic");
    fputs(",\"output_format\":", fp);
    json_write_string(fp, options->output_format == OUTPUT_FORMAT_CSV ? "csv" :
                          options->output_format == OUTPUT_FORMAT_JSONL ? "jsonl" : "tsv");
    fprintf(fp, ",\"start_row\":%d", options->start_row + 1);
    fputs(",\"output_dir\":", fp);
    json_write_string(fp, options->output_dir ? options->output_dir : "");
    fprintf(fp, ",\"selected_sheet_count\":%d", selected_count);
    fprintf(fp, ",\"processed_sheet_count\":%d", processed_count);
    fprintf(fp, ",\"truncated\":%s", global_warnings->truncated ? "true" : "false");
    fputs(",\"warnings\":[", fp);
    for (int i = 0; i < global_warnings->count; i++) {
        if (i > 0) {
            fputc(',', fp);
        }
        json_write_string(fp, global_warnings->messages[i]);
    }
    fputs("],\"sheets\":[", fp);

    for (int i = 0; i < workbook->sheet_count; i++) {
        const SheetInfo* sheet = &workbook->sheets[i];
        if (!sheet->selected && !sheet->processed && sheet->warning_count == 0) {
            continue;
        }
        if (!first) {
            fputc(',', fp);
        }
        first = 0;
        fputs("{\"sheet_name\":", fp);
        json_write_string(fp, sheet->name);
        fputs(",\"state\":", fp);
        json_write_string(fp, sheet->state);
        fprintf(fp, ",\"hidden\":%s", sheet->hidden ? "true" : "false");
        fprintf(fp, ",\"selected\":%s", sheet->selected ? "true" : "false");
        fprintf(fp, ",\"processed\":%s", sheet->processed ? "true" : "false");
        fprintf(fp, ",\"truncated\":%s", sheet->truncated ? "true" : "false");
        fputs(",\"output_path\":", fp);
        if (sheet->output_name[0] != '\0') {
            json_write_string(fp, sheet->output_name);
        } else {
            fputs("null", fp);
        }
        fprintf(fp, ",\"rows_emitted\":%d,\"cols_emitted\":%d", sheet->emitted_rows, sheet->emitted_cols);
        fprintf(fp, ",\"approx_rows\":%d,\"approx_cols\":%d", sheet->approx_rows, sheet->approx_cols);
        fputs(",\"warnings\":", fp);
        json_write_sheet_warnings(fp, sheet);
        fputc('}', fp);
    }

    fputs("]}", fp);
}

static void write_list_sheets_json(FILE* fp, const Workbook* workbook) {
    fputs("{\"tool_version\":", fp);
    json_write_string(fp, TOOL_VERSION);
    fputs(",\"sheets\":[", fp);
    for (int i = 0; i < workbook->sheet_count; i++) {
        const SheetInfo* sheet = &workbook->sheets[i];
        if (i > 0) {
            fputc(',', fp);
        }
        fputs("{\"sheet_name\":", fp);
        json_write_string(fp, sheet->name);
        fputs(",\"state\":", fp);
        json_write_string(fp, sheet->state);
        fprintf(fp, ",\"hidden\":%s", sheet->hidden ? "true" : "false");
        fprintf(fp, ",\"selected\":%s", sheet->selected ? "true" : "false");
        fprintf(fp, ",\"approx_rows\":%d,\"approx_cols\":%d", sheet->approx_rows, sheet->approx_cols);
        fputc('}', fp);
    }
    fputs("]}", fp);
}

// 고성능 워크시트 파서
void parse_worksheet(const char* xml_data, SharedStrings* ss, int start_row, Filter* output) {
    const char* pos = xml_data;
    int last_row = -1;
    int last_col = -1;

    char r_attr[32];
    char t_attr[32];
    char cell_value[MAX_CELL_VALUE];

    while (!filter_should_stop(output) && (pos = strstr(pos, "<c ")) != NULL) {
        const char* tag_close = strchr(pos, '>');
        const char* cell_content_end;
        const char* cell_end;
        bool self_closing;
        int row;
        int col;
        int has_t_attr;
        if (!tag_close) {
            pos++;
            continue;
        }

        self_closing = *(tag_close - 1) == '/';
        if (self_closing) {
            cell_end = tag_close + 1;
            cell_content_end = tag_close;
        } else {
            cell_content_end = strstr(tag_close + 1, "</c>");
            if (!cell_content_end) {
                pos++;
                continue;
            }
            cell_end = cell_content_end + 4;
        }

        if (find_attribute_in_range(pos, tag_close + 1, "r=", 2, r_attr, sizeof(r_attr)) < 0) {
            pos = cell_end;
            continue;
        }

        if (!parse_cell_ref_parts(r_attr, &row, &col)) {
            pos = cell_end;
            continue;
        }

        if (row < start_row) {
            pos = cell_end;
            continue;
        }

        if (last_row != -1 && row != last_row) {
            filter_finish_line(output);
#ifdef DEBUG
            printf("DEBUG: New row, outputting newline\n");
#endif
            last_col = -1;
            if (filter_should_stop(output)) {
                break;
            }
        }

        // 빈 열을 탭으로 채우기 (last_col과 현재 col 사이의 열들)
        int tabs_needed = col - last_col - 1;

        for (int i = 0; i < tabs_needed; i++) {
            filter_push(output, "");
        }

        has_t_attr = find_attribute_in_range(pos, tag_close + 1, "t=", 2, t_attr, sizeof(t_attr)) >= 0;
        cell_value[0] = '\0';

        if (!self_closing) {
            const char* value_start = find_in_range(tag_close + 1, cell_content_end, "<v>", 3);

            if (value_start) {
                const char* value_end;
                value_start += 3;
                value_end = find_in_range(value_start, cell_content_end, "</v>", 4);
                if (value_end) {
                    if (has_t_attr && strcmp(t_attr, "s") == 0) {
                        int str_index = parse_int_range(value_start, value_end);
                        if (str_index >= 0 && str_index < ss->count) {
                            escape_tsv_value(ss->strings[str_index], cell_value, MAX_CELL_VALUE);
                        }
#ifdef DEBUG
                        printf("DEBUG: Cell %s [+%dtabs] : shared_string[%d] : '%s'\n",
                               r_attr, tabs_needed, str_index, cell_value);
#endif
                    } else {
                        escape_tsv_value_range(value_start, value_end, cell_value, MAX_CELL_VALUE);
#ifdef DEBUG
                        printf("DEBUG: Cell %s [+%dtabs] : direct_value : '%s'\n",
                               r_attr, tabs_needed, cell_value);
#endif
                    }
                }
            }

            if (!cell_value[0]) {
                const char* inline_start = find_in_range(tag_close + 1, cell_content_end, "<is><t>", 7);
                if (inline_start) {
                    const char* inline_end;
                    inline_start += 7;
                    inline_end = find_in_range(inline_start, cell_content_end, "</t></is>", 9);
                    if (inline_end) {
                        escape_tsv_value_range(inline_start, inline_end, cell_value, MAX_CELL_VALUE);
                    }
                }
            }

            if (!cell_value[0]) {
                const char* text_start = find_in_range(tag_close + 1, cell_content_end, "<t>", 3);
                if (text_start) {
                    const char* text_end;
                    text_start += 3;
                    text_end = find_in_range(text_start, cell_content_end, "</t>", 4);
                    if (text_end) {
                        escape_tsv_value_range(text_start, text_end, cell_value, MAX_CELL_VALUE);
                    }
                }
            }
        }

        filter_push(output, cell_value);

        last_row = row;
        last_col = col;

        // 다음 셀로 이동
        pos = cell_end;
    }

    // 행을 처리한 경우 마지막 개행 출력
    if (last_row >= start_row && !filter_should_stop(output)) {
        filter_finish_line(output);
    }
}

static void extract_generic_cell_value(const char* cell_open_tag,
                                       const char* cell_content_start,
                                       const char* cell_close,
                                       bool self_closing,
                                       SharedStrings* ss,
                                       const Styles* styles,
                                       bool formatted_output,
                                       char* cell_value,
                                       int max_len) {
    char s_attr[32];
    char t_attr[32];
    char v_content[MAX_CELL_VALUE];
    bool has_type_attr = FIND_ATTR_T(cell_open_tag, t_attr, sizeof(t_attr)) >= 0;
    int style_index = -1;

    if (FIND_ATTR_S(cell_open_tag, s_attr, sizeof(s_attr)) >= 0) {
        style_index = atoi(s_attr);
    }

    v_content[0] = '\0';
    cell_value[0] = '\0';

    if (self_closing) {
        return;
    }

    if (has_type_attr && strcmp(t_attr, "s") == 0) {
        if (extract_first_tag_text(cell_content_start, cell_close, "v", v_content, sizeof(v_content))) {
            int str_index = atoi(v_content);
            if (str_index >= 0 && str_index < ss->count) {
                escape_tsv_value(ss->strings[str_index], cell_value, max_len);
            }
        }
        return;
    }

    if (has_type_attr && strcmp(t_attr, "inlineStr") == 0) {
        if (extract_inline_string_text(cell_content_start, cell_close, v_content, sizeof(v_content))) {
            escape_tsv_value(v_content, cell_value, max_len);
        }
        return;
    }

    if (extract_first_tag_text(cell_content_start, cell_close, "v", v_content, sizeof(v_content))) {
        format_generic_scalar(v_content, has_type_attr, t_attr, style_index, styles,
                              formatted_output, cell_value, max_len);
        return;
    }

    if (extract_inline_string_text(cell_content_start, cell_close, v_content, sizeof(v_content))) {
        escape_tsv_value(v_content, cell_value, max_len);
    }
}

void parse_worksheet_generic(const char* xml_data, SharedStrings* ss, const Styles* styles,
                             const HiddenColumns* hidden_columns, MergeRegions* merge_regions,
                             int start_row, bool formatted_output, bool skip_hidden,
                             Filter* output) {
    const char* xml_end = xml_data + strlen(xml_data);
    const char* pos = xml_data;
    int last_parsed_row = -1;
    int inferred_row = -1;
    int target_max_col = scan_generic_visible_max_col(xml_data, start_row, skip_hidden,
                                                      hidden_columns, merge_regions);
    RowBuffer row_buffer;

    init_row_buffer(&row_buffer);

    while (!filter_should_stop(output) &&
           (pos = find_next_start_tag_local(pos, xml_end, "row")) != NULL) {
        const char* row_tag_end = find_tag_end_in_range(pos, xml_end);
        const char* row_content_start;
        const char* row_close;
        const char* row_close_end;
        bool self_closing_row;
        char row_open_tag[256];
        char row_attr[32];
        char hidden_attr[16];
        int row;
        bool row_hidden = false;

        if (!row_tag_end) {
            break;
        }

        copy_tag_range(pos, row_tag_end, row_open_tag, sizeof(row_open_tag));
        self_closing_row = is_self_closing_tag(pos, row_tag_end);
        row_content_start = row_tag_end + 1;
        row_close = self_closing_row ? NULL : find_next_end_tag_local(row_content_start, xml_end, "row");
        if (!self_closing_row && !row_close) {
            pos = row_tag_end + 1;
            continue;
        }

        row_close_end = self_closing_row ? row_tag_end : find_tag_end_in_range(row_close, xml_end);
        if (!row_close_end) {
            break;
        }

        row = inferred_row + 1;
        if (find_attribute(row_open_tag, "r=", 2, row_attr, sizeof(row_attr)) >= 0) {
            row = atoi(row_attr) - 1;
        }
        inferred_row = row;

        if (last_parsed_row >= 0) {
            int synthetic_start = last_parsed_row + 1;
            if (synthetic_start < start_row) {
                synthetic_start = start_row;
            }
            if (synthetic_start < row) {
                emit_synthetic_merge_rows(synthetic_start, row, hidden_columns, merge_regions,
                                          target_max_col, &row_buffer, output);
                if (filter_should_stop(output)) {
                    break;
                }
            }
        }

        clear_row_buffer(&row_buffer);

        if (skip_hidden &&
            FIND_ATTR_HIDDEN(row_open_tag, hidden_attr, sizeof(hidden_attr)) >= 0 &&
            attr_is_true(hidden_attr)) {
            row_hidden = true;
        }

        if (!self_closing_row) {
            const char* cell_pos = row_content_start;

            while ((cell_pos = find_next_start_tag_local(cell_pos, row_close, "c")) != NULL) {
                const char* cell_tag_end = find_tag_end_in_range(cell_pos, row_close);
                const char* cell_content_start;
                const char* cell_close;
                const char* cell_close_end;
                bool self_closing_cell;
                char cell_open_tag[512];
                char r_attr[32];
                char cell_value[MAX_CELL_VALUE];
                int col;

                if (!cell_tag_end) {
                    break;
                }

                copy_tag_range(cell_pos, cell_tag_end, cell_open_tag, sizeof(cell_open_tag));
                self_closing_cell = is_self_closing_tag(cell_pos, cell_tag_end);
                cell_content_start = cell_tag_end + 1;
                cell_close = self_closing_cell ? NULL : find_next_end_tag_local(cell_content_start, row_close, "c");
                if (!self_closing_cell && !cell_close) {
                    cell_pos = cell_tag_end + 1;
                    continue;
                }

                if (FIND_ATTR_R(cell_open_tag, r_attr, sizeof(r_attr)) < 0) {
                    cell_pos = self_closing_cell ? cell_tag_end + 1 : cell_close + 1;
                    continue;
                }

                col = col_ref_to_num(r_attr);
                extract_generic_cell_value(cell_open_tag, cell_content_start, cell_close, self_closing_cell,
                                           ss, styles, formatted_output, cell_value, sizeof(cell_value));
                add_row_cell(&row_buffer, col, cell_value);
                register_merge_anchor_value(merge_regions, row, col, cell_value);

                if (self_closing_cell) {
                    cell_pos = cell_tag_end + 1;
                } else {
                    cell_close_end = find_tag_end_in_range(cell_close, row_close);
                    if (!cell_close_end) {
                        break;
                    }
                    cell_pos = cell_close_end + 1;
                }
            }
        }

        apply_merge_regions_to_row(merge_regions, row, &row_buffer);

        if (row >= start_row && !row_hidden) {
            emit_row_buffer(&row_buffer, hidden_columns, target_max_col, target_max_col < 0, output);
            if (filter_should_stop(output)) {
                break;
            }
        }

        last_parsed_row = row;
        pos = row_close_end + 1;
    }

    if (last_parsed_row >= 0) {
        int trailing_start = last_parsed_row + 1;
        int trailing_end = max_merge_row_with_values(merge_regions) + 1;
        if (trailing_start < start_row) {
            trailing_start = start_row;
        }
        if (trailing_start < trailing_end) {
            emit_synthetic_merge_rows(trailing_start, trailing_end, hidden_columns, merge_regions,
                                      target_max_col, &row_buffer, output);
        }
    }

    free_row_buffer(&row_buffer);
}

// 공유 문자열 메모리 해제
void free_shared_strings(SharedStrings* ss) {
    for (int i = 0; i < ss->count; i++) {
        free(ss->strings[i]);
    }
    free(ss->strings);
    ss->count = 0;
    ss->capacity = 0;
}

int main(int argc, char* argv[]) {
    const char* input_file;
    int start_row = 0;
    bool export_all_sheets;
    bool game_db_fast_mode = false;
    bool formatted_output = false;
    bool expand_merged_cells = false;
    bool skip_hidden = false;
    bool all_sheets_alias = false;
    bool no_wildcard_mode = false;
    bool start_row_set = false;
    OutputFormat output_format = OUTPUT_FORMAT_TSV;
    const char* output_dir = NULL;
    bool list_sheets = false;
    bool json_output = false;
    bool manifest_stdout = false;
    const char* manifest_json_path = NULL;
    bool stdout_output = false;
    bool fail_if_no_sheet = false;
    bool fail_on_output_collision = false;
    bool fail_if_truncated = false;
    int max_sheets = 0;
    int max_rows_per_sheet = 0;
    size_t max_output_bytes = 0;
    char sheet_names[MAX_SHEET_FILTERS][MAX_SHEET_NAME];
    int sheet_name_count = 0;
    const char* sheet_regex_pattern = NULL;
    regex_t sheet_regex;
    bool regex_compiled = false;
    bool zip_open = false;
    int processed_sheets = 0;
    int selected_sheets = 0;
    int exit_code = 1;
    size_t total_output_bytes = 0;
    char (*used_output_names)[MAX_OUTPUT_FILENAME] = NULL;
    GlobalWarnings global_warnings = {0};
    Styles styles;
    SharedStrings shared_strings;
    Workbook workbook;
    mz_zip_archive zip;
    RunOptions run_options = {0};

    g_log_fp = stdout;
    init_styles(&styles);
    styles.enabled = false;
    init_shared_strings(&shared_strings);
    init_workbook(&workbook);

    if (argc == 2 && strcmp(argv[1], "--version") == 0) {
        fprintf(stdout, "xlsx2tsv %s\n", TOOL_VERSION);
        free_shared_strings(&shared_strings);
        free_styles(&styles);
        free_workbook(&workbook);
        return 0;
    }

    if (argc < 2 || argv[1][0] == '-') {
        fprintf(stdout, "Usage: %s <input.xlsx> [start_row] [--mode generic|game-db-fast] [--output-dir dir] [--sheet name] [--sheet-regex pattern] [--list-sheets] [--json] [--manifest-json path] [--manifest-stdout] [--stdout] [--max-sheets n] [--max-rows-per-sheet n] [--max-output-bytes n] [--fail-if-truncated] [--fail-if-no-sheet] [--fail-on-output-collision] [--no-wildcard] [--formatted] [--expand-merged] [--skip-hidden] [--csv] [--jsonl]\n", argv[0]);
        fprintf(stdout, "  --version: Print a stable version string and exit\n");
        fprintf(stdout, "  --sheet name: Select one sheet by exact name (repeatable)\n");
        fprintf(stdout, "  --sheet-regex pattern: Select sheets by POSIX regex\n");
        fprintf(stdout, "  --list-sheets: Inspect workbook sheets without converting\n");
        fprintf(stdout, "  --json: Emit JSON for --list-sheets\n");
        fprintf(stdout, "  --manifest-json path / --manifest-stdout: Write conversion manifest JSON\n");
        fprintf(stdout, "  --stdout: Write a single selected sheet to stdout\n");
        fprintf(stdout, "  --max-sheets / --max-rows-per-sheet / --max-output-bytes: Resource guards\n");
        fprintf(stdout, "  --fail-if-truncated / --fail-if-no-sheet / --fail-on-output-collision: Strict modes\n");
        free_shared_strings(&shared_strings);
        free_styles(&styles);
        free_workbook(&workbook);
        return 1;
    }

    input_file = argv[1];

    for (int i = 2; i < argc; i++) {
        if (strcmp(argv[i], "--mode") == 0) {
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: --mode requires a value: generic or game-db-fast\n");
                goto cleanup;
            }
            i++;
            if (strcmp(argv[i], "generic") == 0) {
                game_db_fast_mode = false;
            } else if (strcmp(argv[i], "game-db-fast") == 0) {
                game_db_fast_mode = true;
            } else {
                fprintf(stderr, "Error: Unknown mode: %s (expected generic or game-db-fast)\n", argv[i]);
                goto cleanup;
            }
            continue;
        }

        if (strcmp(argv[i], "--output-dir") == 0) {
            struct stat st;
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: --output-dir requires a directory path\n");
                goto cleanup;
            }
            output_dir = argv[++i];
            if (stat(output_dir, &st) != 0 || !S_ISDIR(st.st_mode)) {
                fprintf(stderr, "Error: Output directory does not exist or is not a directory: %s\n", output_dir);
                goto cleanup;
            }
            continue;
        }

        if (strcmp(argv[i], "--sheet") == 0) {
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: --sheet requires a sheet name\n");
                goto cleanup;
            }
            if (sheet_name_count >= MAX_SHEET_FILTERS) {
                fprintf(stderr, "Error: Too many --sheet filters\n");
                goto cleanup;
            }
            strncpy(sheet_names[sheet_name_count], argv[++i], MAX_SHEET_NAME - 1);
            sheet_names[sheet_name_count][MAX_SHEET_NAME - 1] = '\0';
            sheet_name_count++;
            continue;
        }

        if (strcmp(argv[i], "--sheet-regex") == 0) {
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: --sheet-regex requires a pattern\n");
                goto cleanup;
            }
            sheet_regex_pattern = argv[++i];
            continue;
        }

        if (strcmp(argv[i], "--list-sheets") == 0) {
            list_sheets = true;
            continue;
        }

        if (strcmp(argv[i], "--json") == 0) {
            json_output = true;
            continue;
        }

        if (strcmp(argv[i], "--manifest-json") == 0) {
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: --manifest-json requires a file path\n");
                goto cleanup;
            }
            manifest_json_path = argv[++i];
            continue;
        }

        if (strcmp(argv[i], "--manifest-stdout") == 0) {
            manifest_stdout = true;
            continue;
        }

        if (strcmp(argv[i], "--stdout") == 0) {
            stdout_output = true;
            continue;
        }

        if (strcmp(argv[i], "--fail-if-no-sheet") == 0) {
            fail_if_no_sheet = true;
            continue;
        }

        if (strcmp(argv[i], "--fail-on-output-collision") == 0) {
            fail_on_output_collision = true;
            continue;
        }

        if (strcmp(argv[i], "--fail-if-truncated") == 0) {
            fail_if_truncated = true;
            continue;
        }

        if (strcmp(argv[i], "--max-sheets") == 0) {
            char* end_ptr = NULL;
            long parsed_value;
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: --max-sheets requires a value\n");
                goto cleanup;
            }
            parsed_value = strtol(argv[++i], &end_ptr, 10);
            if (!end_ptr || *end_ptr != '\0' || parsed_value <= 0) {
                fprintf(stderr, "Error: Invalid --max-sheets value: %s\n", argv[i]);
                goto cleanup;
            }
            max_sheets = (int)parsed_value;
            continue;
        }

        if (strcmp(argv[i], "--max-rows-per-sheet") == 0) {
            char* end_ptr = NULL;
            long parsed_value;
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: --max-rows-per-sheet requires a value\n");
                goto cleanup;
            }
            parsed_value = strtol(argv[++i], &end_ptr, 10);
            if (!end_ptr || *end_ptr != '\0' || parsed_value <= 0) {
                fprintf(stderr, "Error: Invalid --max-rows-per-sheet value: %s\n", argv[i]);
                goto cleanup;
            }
            max_rows_per_sheet = (int)parsed_value;
            continue;
        }

        if (strcmp(argv[i], "--max-output-bytes") == 0) {
            char* end_ptr = NULL;
            unsigned long long parsed_value;
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: --max-output-bytes requires a value\n");
                goto cleanup;
            }
            parsed_value = strtoull(argv[++i], &end_ptr, 10);
            if (!end_ptr || *end_ptr != '\0' || parsed_value == 0) {
                fprintf(stderr, "Error: Invalid --max-output-bytes value: %s\n", argv[i]);
                goto cleanup;
            }
            max_output_bytes = (size_t)parsed_value;
            continue;
        }

        if (strcmp(argv[i], "--no-wildcard") == 0) {
            ALLOW_WILD_CARD = false;
            no_wildcard_mode = true;
            continue;
        }

        if (strcmp(argv[i], "--all-sheets") == 0) {
            all_sheets_alias = true;
            continue;
        }

        if (strcmp(argv[i], "--formatted") == 0) {
            formatted_output = true;
            continue;
        }

        if (strcmp(argv[i], "--expand-merged") == 0) {
            expand_merged_cells = true;
            continue;
        }

        if (strcmp(argv[i], "--skip-hidden") == 0) {
            skip_hidden = true;
            continue;
        }

        if (strcmp(argv[i], "--csv") == 0) {
            output_format = OUTPUT_FORMAT_CSV;
            continue;
        }

        if (strcmp(argv[i], "--jsonl") == 0) {
            output_format = OUTPUT_FORMAT_JSONL;
            continue;
        }

        if (strcmp(argv[i], "--version") == 0) {
            fprintf(stdout, "xlsx2tsv %s\n", TOOL_VERSION);
            exit_code = 0;
            goto cleanup;
        }

        char* end_ptr = NULL;
        long parsed_row = strtol(argv[i], &end_ptr, 10);
        if (!start_row_set && argv[i][0] != '\0' && end_ptr && *end_ptr == '\0') {
            start_row = (int)parsed_row - 1;
            start_row_set = true;
            continue;
        }

        fprintf(stderr, "Error: Unknown argument: %s\n", argv[i]);
        goto cleanup;
    }

    if (start_row < 0) {
        start_row = 0;
    }

    if (sheet_regex_pattern) {
        if (regcomp(&sheet_regex, sheet_regex_pattern, REG_EXTENDED | REG_NOSUB) != 0) {
            fprintf(stderr, "Error: Invalid --sheet-regex pattern: %s\n", sheet_regex_pattern);
            goto cleanup;
        }
        regex_compiled = true;
    }

    if (json_output && !list_sheets) {
        fprintf(stderr, "Error: --json is only supported with --list-sheets\n");
        goto cleanup;
    }

    if (manifest_stdout && manifest_json_path) {
        fprintf(stderr, "Error: --manifest-stdout cannot be combined with --manifest-json\n");
        goto cleanup;
    }

    if (manifest_stdout && stdout_output) {
        fprintf(stderr, "Error: --manifest-stdout cannot be combined with --stdout\n");
        goto cleanup;
    }

    if (stdout_output && output_dir) {
        fprintf(stderr, "Error: --stdout cannot be combined with --output-dir\n");
        goto cleanup;
    }

    if (stdout_output && fail_on_output_collision) {
        fprintf(stderr, "Error: --stdout cannot be combined with --fail-on-output-collision\n");
        goto cleanup;
    }

    if (list_sheets &&
        (formatted_output || expand_merged_cells || skip_hidden || manifest_stdout ||
         manifest_json_path || stdout_output || output_dir || output_format != OUTPUT_FORMAT_TSV ||
         max_sheets > 0 || max_rows_per_sheet > 0 || max_output_bytes > 0 ||
         fail_if_truncated || fail_on_output_collision || no_wildcard_mode || all_sheets_alias)) {
        fprintf(stderr, "Error: --list-sheets cannot be combined with conversion-only options\n");
        goto cleanup;
    }

    export_all_sheets = !game_db_fast_mode;
    if (game_db_fast_mode &&
        !list_sheets &&
        (formatted_output || expand_merged_cells || skip_hidden ||
         output_format != OUTPUT_FORMAT_TSV || all_sheets_alias)) {
        fprintf(stderr, "Error: --mode game-db-fast cannot be combined with generic-mode options\n");
        goto cleanup;
    }

    if (!game_db_fast_mode && no_wildcard_mode) {
        fprintf(stderr, "Error: --no-wildcard requires --mode game-db-fast\n");
        goto cleanup;
    }

    run_options.list_sheets = list_sheets;
    run_options.json = json_output;
    run_options.manifest_stdout = manifest_stdout;
    run_options.manifest_json_path = manifest_json_path;
    run_options.stdout_output = stdout_output;
    run_options.fail_if_no_sheet = fail_if_no_sheet;
    run_options.fail_on_output_collision = fail_on_output_collision;
    run_options.max_sheets = max_sheets;
    run_options.max_rows_per_sheet = max_rows_per_sheet;
    run_options.max_output_bytes = max_output_bytes;
    run_options.fail_if_truncated = fail_if_truncated;
    run_options.output_dir = output_dir;
    run_options.output_format = output_format;
    run_options.start_row = start_row;
    run_options.game_db_fast_mode = game_db_fast_mode;

    if (stdout_output || manifest_stdout || (list_sheets && json_output)) {
        g_log_fp = stderr;
    } else {
        g_log_fp = stdout;
    }

    clock_t start_time = clock();
    styles.enabled = formatted_output;

    if (!mz_zip_reader_init_file(&zip, input_file)) {
        fprintf(stderr, "Error: Could not open XLSX file: %s\n", input_file);
        goto cleanup;
    }
    zip_open = true;

    int workbook_index;
    if (!mz_zip_reader_locate_file(&zip, "xl/workbook.xml", &workbook_index)) {
        fprintf(stderr, "Error: Could not find workbook.xml in XLSX file\n");
        goto cleanup;
    }

    size_t workbook_size = mz_zip_reader_get_file_size(&zip, workbook_index);
    char* workbook_data = malloc(workbook_size + 1);
    if (!mz_zip_reader_extract_to_mem(&zip, workbook_index, workbook_data, workbook_size)) {
        fprintf(stderr, "Error: Could not extract workbook.xml\n");
        free(workbook_data);
        goto cleanup;
    }
    workbook_data[workbook_size] = '\0';
    styles.date_1904 = workbook_uses_1904_date_system(workbook_data);
    parse_workbook(workbook_data, &workbook, export_all_sheets);
    free(workbook_data);

    int workbook_rels_index;
    if (!mz_zip_reader_locate_file(&zip, "xl/_rels/workbook.xml.rels", &workbook_rels_index)) {
        fprintf(stderr, "Error: Could not find workbook.xml.rels in XLSX file\n");
        goto cleanup;
    }

    size_t workbook_rels_size = mz_zip_reader_get_file_size(&zip, workbook_rels_index);
    char* workbook_rels_data = malloc(workbook_rels_size + 1);
    if (!mz_zip_reader_extract_to_mem(&zip, workbook_rels_index, workbook_rels_data, workbook_rels_size)) {
        fprintf(stderr, "Error: Could not extract workbook.xml.rels\n");
        free(workbook_rels_data);
        goto cleanup;
    }
    workbook_rels_data[workbook_rels_size] = '\0';
    resolve_sheet_filenames(workbook_rels_data, &workbook);
    free(workbook_rels_data);

    if (workbook.sheet_count == 0) {
        fprintf(stderr, export_all_sheets ?
                "No sheets found in workbook\n" :
                "No valid sheets found (sheets must contain only A-Z, a-z, 0-9, -, _, *)\n");
        goto cleanup;
    }

    for (int i = 0; i < workbook.sheet_count; i++) {
        workbook.sheets[i].selected = sheet_matches_filters(&workbook.sheets[i],
                                                            sheet_names, sheet_name_count,
                                                            regex_compiled ? &sheet_regex : NULL,
                                                            regex_compiled);
        if (workbook.sheets[i].selected) {
            selected_sheets++;
        }
    }

    if (selected_sheets == 0 && fail_if_no_sheet) {
        fprintf(stderr, "Error: No sheets matched the provided sheet filters\n");
        goto cleanup;
    }

    if (max_sheets > 0 && selected_sheets > max_sheets) {
        int kept = 0;
        char warning_message[MAX_WARNING_MESSAGE];
        snprintf(warning_message, sizeof(warning_message),
                 "Selected sheets truncated to the first %d due to --max-sheets", max_sheets);
        add_global_warning(&global_warnings, warning_message);
        global_warnings.truncated = true;

        for (int i = 0; i < workbook.sheet_count; i++) {
            if (!workbook.sheets[i].selected) {
                continue;
            }
            if (kept < max_sheets) {
                kept++;
            } else {
                workbook.sheets[i].selected = false;
            }
        }
        selected_sheets = max_sheets;
    }

    if (stdout_output && selected_sheets != 1) {
        fprintf(stderr, "Error: --stdout requires exactly one selected sheet\n");
        goto cleanup;
    }

    if (list_sheets) {
        for (int i = 0; i < workbook.sheet_count; i++) {
            int worksheet_index;
            if (workbook.sheets[i].filename[0] == '\0') {
                continue;
            }
            if (!mz_zip_reader_locate_file(&zip, workbook.sheets[i].filename, &worksheet_index)) {
                continue;
            }

            size_t worksheet_size = mz_zip_reader_get_file_size(&zip, worksheet_index);
            char* worksheet_data = malloc(worksheet_size + 1);
            if (!mz_zip_reader_extract_to_mem(&zip, worksheet_index, worksheet_data, worksheet_size)) {
                free(worksheet_data);
                continue;
            }
            worksheet_data[worksheet_size] = '\0';
            parse_sheet_dimension_xml(worksheet_data, &workbook.sheets[i].approx_rows, &workbook.sheets[i].approx_cols);
            free(worksheet_data);
        }

        if (json_output) {
            write_list_sheets_json(stdout, &workbook);
            fputc('\n', stdout);
        } else {
            for (int i = 0; i < workbook.sheet_count; i++) {
                fprintf(stdout, "%s\t%s\t%s\t%d\t%d\n",
                        workbook.sheets[i].name,
                        workbook.sheets[i].state,
                        workbook.sheets[i].selected ? "selected" : "skipped",
                        workbook.sheets[i].approx_rows,
                        workbook.sheets[i].approx_cols);
            }
        }
        exit_code = 0;
        goto cleanup;
    }

    printf("Converting XLSX to multiple TSV files...\n");
    printf("Input: %s\n", input_file);
    printf("Starting from row: %d\n", start_row + 1);
    if (output_dir && output_dir[0] != '\0') {
        printf("Output directory: %s\n", output_dir);
    }
    printf("Mode: %s\n", export_all_sheets ? "generic" : "game-db-fast");
    if (selected_sheets > 0) {
        printf("Selected sheets: %d\n", selected_sheets);
    }
    if (formatted_output) {
        printf("Formatting: enabled\n");
    }
    if (expand_merged_cells) {
        printf("Merged cells: expanded\n");
    }
    if (skip_hidden) {
        printf("Hidden rows/columns: skipped\n");
    }
    if (output_format == OUTPUT_FORMAT_CSV) {
        printf("Output format: csv\n");
    } else if (output_format == OUTPUT_FORMAT_JSONL) {
        printf("Output format: jsonl\n");
    }
    if (stdout_output) {
        printf("Output target: stdout\n");
    }
    if (max_rows_per_sheet > 0) {
        printf("Max rows per sheet: %d\n", max_rows_per_sheet);
    }
    if (max_output_bytes > 0) {
        printf("Max output bytes: %zu\n", max_output_bytes);
    }
    printf("Found %d sheet(s) to process\n\n", selected_sheets);

    int shared_strings_index;
    if (mz_zip_reader_locate_file(&zip, "xl/sharedStrings.xml", &shared_strings_index)) {
        printf("Loading shared strings...\n");
        size_t shared_strings_size = mz_zip_reader_get_file_size(&zip, shared_strings_index);
        char* shared_strings_data = malloc(shared_strings_size + 1);

        if (mz_zip_reader_extract_to_mem(&zip, shared_strings_index, shared_strings_data, shared_strings_size)) {
            shared_strings_data[shared_strings_size] = '\0';
            if (export_all_sheets) {
                parse_shared_strings_generic(shared_strings_data, &shared_strings);
            } else {
                parse_shared_strings(shared_strings_data, &shared_strings);
            }
            printf("Loaded %d shared strings\n\n", shared_strings.count);
        }

        free(shared_strings_data);
    }

    if (formatted_output) {
        int styles_index;
        if (mz_zip_reader_locate_file(&zip, "xl/styles.xml", &styles_index)) {
            size_t styles_size = mz_zip_reader_get_file_size(&zip, styles_index);
            char* styles_data = malloc(styles_size + 1);
            if (mz_zip_reader_extract_to_mem(&zip, styles_index, styles_data, styles_size)) {
                styles_data[styles_size] = '\0';
                parse_styles_xml(styles_data, &styles);
            }
            free(styles_data);
        }
    }

    used_output_names = calloc((size_t)workbook.sheet_count, sizeof(*used_output_names));
    int used_output_count = 0;
    const char* output_extension = "tsv";
    if (output_format == OUTPUT_FORMAT_CSV) {
        output_extension = "csv";
    } else if (output_format == OUTPUT_FORMAT_JSONL) {
        output_extension = "jsonl";
    }

    int selected_index = 0;
    bool output_limit_reached = false;
    for (int i = 0; i < workbook.sheet_count; i++) {
        SheetInfo* sheet = &workbook.sheets[i];
        if (!sheet->selected) {
            continue;
        }

        if (output_limit_reached) {
            add_sheet_warning(sheet, "max-output-bytes", "Sheet was skipped after --max-output-bytes was reached");
            continue;
        }

        selected_index++;
        printf("Processing sheet %d/%d: '%s'\n", selected_index, selected_sheets, sheet->name);

        if (sheet->filename[0] == '\0') {
            add_sheet_warning(sheet, "missing-worksheet", "Could not resolve worksheet file for sheet");
            printf("Warning: Could not resolve worksheet file for sheet: %s - skipping\n\n", sheet->name);
            continue;
        }

        int worksheet_index;
        if (!mz_zip_reader_locate_file(&zip, sheet->filename, &worksheet_index)) {
            add_sheet_warning(sheet, "missing-worksheet", "Could not find worksheet file in archive");
            printf("Warning: Could not find worksheet file: %s - skipping\n\n", sheet->filename);
            continue;
        }

        size_t worksheet_size = mz_zip_reader_get_file_size(&zip, worksheet_index);
        char* worksheet_data = malloc(worksheet_size + 1);
        if (!mz_zip_reader_extract_to_mem(&zip, worksheet_index, worksheet_data, worksheet_size)) {
            add_sheet_warning(sheet, "extract-failed", "Could not extract worksheet data");
            printf("Warning: Could not extract worksheet data for: %s - skipping\n\n", sheet->name);
            free(worksheet_data);
            continue;
        }

        worksheet_data[worksheet_size] = '\0';
        parse_sheet_dimension_xml(worksheet_data, &sheet->approx_rows, &sheet->approx_cols);

        HiddenColumns hidden_columns;
        MergeRegions merge_regions;
        init_hidden_columns(&hidden_columns);
        init_merge_regions(&merge_regions);

        if (export_all_sheets) {
            if (skip_hidden) {
                parse_hidden_columns_xml(worksheet_data, &hidden_columns);
            }
            if (expand_merged_cells) {
                parse_merge_regions_xml(worksheet_data, &merge_regions);
            }
        }

        Filter* output = NULL;
        if (stdout_output) {
            output = filter_init_stream(stdout,
                                        export_all_sheets ? FILTER_MODE_RAW : FILTER_MODE_GAME_DB,
                                        output_format);
        } else {
            char output_filename[MAX_OUTPUT_FILENAME];
            char output_path[MAX_OUTPUT_PATH];

            create_safe_filename(sheet->name, output_extension, output_filename, sizeof(output_filename));
            bool collision = output_name_exists(used_output_names, used_output_count, output_filename);
            if (collision && fail_on_output_collision) {
                fprintf(stderr, "Error: Output filename collision detected for sheet: %s\n", sheet->name);
                free_hidden_columns(&hidden_columns);
                free_merge_regions(&merge_regions);
                free(worksheet_data);
                goto cleanup;
            }

            if (collision) {
                add_sheet_warning(sheet, "output-collision", "Output filename collision resolved with suffix");
            }

            create_unique_output_filename(sheet->name, i, used_output_names, used_output_count,
                                          output_extension, output_filename, sizeof(output_filename));
            join_output_path(output_dir, output_filename, output_path, sizeof(output_path));

            output = filter_init(output_path,
                                 export_all_sheets ? FILTER_MODE_RAW : FILTER_MODE_GAME_DB,
                                 output_format);
            if (output) {
                strncpy(sheet->output_name, output_path, MAX_OUTPUT_PATH - 1);
                sheet->output_name[MAX_OUTPUT_PATH - 1] = '\0';
                strncpy(used_output_names[used_output_count], output_filename, MAX_OUTPUT_FILENAME - 1);
                used_output_names[used_output_count][MAX_OUTPUT_FILENAME - 1] = '\0';
                used_output_count++;
                printf("  Output file: %s\n", output_path);
            }
        }

        if (!output) {
            add_sheet_warning(sheet, "output-open-failed", "Could not create output target");
            printf("Warning: Could not create output target for sheet: %s - skipping\n\n", sheet->name);
            free_hidden_columns(&hidden_columns);
            free_merge_regions(&merge_regions);
            free(worksheet_data);
            continue;
        }

        filter_set_limits(output, max_rows_per_sheet, max_output_bytes,
                          max_output_bytes > 0 ? &total_output_bytes : NULL);

        if (export_all_sheets) {
            parse_worksheet_generic(worksheet_data, &shared_strings, &styles,
                                    &hidden_columns, &merge_regions, start_row,
                                    formatted_output, skip_hidden, output);
        } else {
            parse_worksheet(worksheet_data, &shared_strings, start_row, output);
        }

        sheet->emitted_rows = filter_row_count(output);
        sheet->emitted_cols = filter_max_cols(output);
        sheet->truncated = filter_was_truncated(output);
        bool output_limit_hit = filter_hit_output_limit(output);
        if (sheet->truncated) {
            global_warnings.truncated = true;
            if (max_rows_per_sheet > 0 && sheet->emitted_rows >= max_rows_per_sheet) {
                add_sheet_warning(sheet, "max-rows-per-sheet", "Sheet output was truncated by --max-rows-per-sheet");
            }
            if (output_limit_hit) {
                add_sheet_warning(sheet, "max-output-bytes", "Output was truncated by --max-output-bytes");
                add_global_warning(&global_warnings, "Output was truncated by --max-output-bytes");
                output_limit_reached = true;
            }
        }

        filter_close(output);
        free_hidden_columns(&hidden_columns);
        free_merge_regions(&merge_regions);
        free(worksheet_data);
        sheet->processed = true;
        processed_sheets++;

        printf("  Sheet '%s' processed successfully!\n\n", sheet->name);

        if (!output_limit_reached &&
            max_output_bytes > 0 &&
            total_output_bytes >= max_output_bytes &&
            selected_index < selected_sheets) {
            global_warnings.truncated = true;
            add_global_warning(&global_warnings, "Further sheets were skipped after --max-output-bytes was reached");
            output_limit_reached = true;
        }

        if (output_limit_reached) {
            printf("Stopping further sheet conversion because --max-output-bytes was reached.\n\n");
        }
    }

    if (manifest_json_path) {
        FILE* manifest_fp = fopen(manifest_json_path, "w");
        if (!manifest_fp) {
            fprintf(stderr, "Error: Could not create manifest file: %s\n", manifest_json_path);
            goto cleanup;
        }
        write_manifest_json(manifest_fp, input_file, &run_options, &workbook, &global_warnings);
        fputc('\n', manifest_fp);
        fclose(manifest_fp);
    }

    if (manifest_stdout) {
        write_manifest_json(stdout, input_file, &run_options, &workbook, &global_warnings);
        fputc('\n', stdout);
    }

    clock_t end_time = clock();
    double elapsed = ((double)(end_time - start_time)) / CLOCKS_PER_SEC;

    printf("=== Conversion Summary ===\n");
    printf("Total sheets processed: %d out of %d\n", processed_sheets, selected_sheets);
    printf("Processing time: %.2f seconds\n", elapsed);

    if (processed_sheets > 0) {
        printf("Conversion completed successfully!\n");
        if (!stdout_output) {
            printf("Output files created:\n");
            for (int i = 0; i < workbook.sheet_count; i++) {
                if (workbook.sheets[i].output_name[0] != '\0') {
                    printf("  - %s (from sheet: %s)\n", workbook.sheets[i].output_name, workbook.sheets[i].name);
                }
            }
        }
        exit_code = fail_if_truncated && global_warnings.truncated ? 1 : 0;
    } else {
        printf("No sheets were processed successfully.\n");
        exit_code = 1;
    }

cleanup:
    if (regex_compiled) {
        regfree(&sheet_regex);
    }
    free(used_output_names);
    if (zip_open) {
        mz_zip_reader_end(&zip);
    }
    free_shared_strings(&shared_strings);
    free_styles(&styles);
    free_workbook(&workbook);
    return exit_code;
}
