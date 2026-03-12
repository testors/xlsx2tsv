#include "filter.h"

bool ALLOW_WILD_CARD = true;

static void free_string_list(char** items, int count) {
    for (int i = 0; i < count; i++) {
        free(items[i]);
    }
    free(items);
}

static void ensure_raw_field_capacity(Filter* filter, int needed) {
    if (needed <= filter->raw_field_capacity) {
        return;
    }

    int new_capacity = filter->raw_field_capacity ? filter->raw_field_capacity : 16;
    while (new_capacity < needed) {
        new_capacity *= 2;
    }

    filter->raw_fields = realloc(filter->raw_fields, sizeof(char*) * new_capacity);
    filter->raw_field_capacity = new_capacity;
}

static void ensure_json_header_capacity(Filter* filter, int needed) {
    if (needed <= filter->json_header_capacity) {
        return;
    }

    int new_capacity = filter->json_header_capacity ? filter->json_header_capacity : 16;
    while (new_capacity < needed) {
        new_capacity *= 2;
    }

    filter->json_headers = realloc(filter->json_headers, sizeof(char*) * new_capacity);
    for (int i = filter->json_header_capacity; i < new_capacity; i++) {
        filter->json_headers[i] = NULL;
    }
    filter->json_header_capacity = new_capacity;
}

static void reset_raw_fields(Filter* filter) {
    for (int i = 0; i < filter->raw_field_count; i++) {
        free(filter->raw_fields[i]);
        filter->raw_fields[i] = NULL;
    }
    filter->raw_field_count = 0;
}

static int header_name_exists(Filter* filter, const char* name, int limit) {
    for (int i = 0; i < limit; i++) {
        if (filter->json_headers[i] && strcmp(filter->json_headers[i], name) == 0) {
            return 1;
        }
    }
    return 0;
}

static void assign_unique_json_header(Filter* filter, int index, const char* source, int fallback_col) {
    char base_name[256];
    char unique_name[256];

    if (source && source[0] != '\0') {
        strncpy(base_name, source, sizeof(base_name) - 1);
        base_name[sizeof(base_name) - 1] = '\0';
    } else {
        snprintf(base_name, sizeof(base_name), "col_%d", fallback_col + 1);
    }

    strncpy(unique_name, base_name, sizeof(unique_name) - 1);
    unique_name[sizeof(unique_name) - 1] = '\0';

    for (int suffix = 2; header_name_exists(filter, unique_name, index); suffix++) {
        snprintf(unique_name, sizeof(unique_name), "%s__%d", base_name, suffix);
    }

    free(filter->json_headers[index]);
    filter->json_headers[index] = strdup(unique_name);
}

static void ensure_json_header_names(Filter* filter, int count) {
    ensure_json_header_capacity(filter, count);
    for (int i = filter->json_header_count; i < count; i++) {
        assign_unique_json_header(filter, i, NULL, i);
    }
    if (filter->json_header_count < count) {
        filter->json_header_count = count;
    }
}

static void capture_json_headers(Filter* filter) {
    ensure_json_header_capacity(filter, filter->raw_field_count);
    for (int i = 0; i < filter->raw_field_count; i++) {
        assign_unique_json_header(filter, i, filter->raw_fields[i], i);
    }
    filter->json_header_count = filter->raw_field_count;
}

Filter* filter_init(const char* filename, FilterMode mode, OutputFormat output_format) {
    Filter* filter = (Filter*)malloc(sizeof(Filter));
    if (!filter) {
        return NULL;
    }

    filter->mode = mode;
    filter->output_format = output_format;
    filter->col_count = 0;
    filter->row_count = 0;
    filter->valid_col_count = 0;
    filter->buffer_pos = 0;
    filter->raw_fields = NULL;
    filter->raw_field_count = 0;
    filter->raw_field_capacity = 0;
    filter->json_headers = NULL;
    filter->json_header_count = 0;
    filter->json_header_capacity = 0;

    // 명시적으로 모든 포인터를 NULL로 초기화
    for (int i = 0; i < MAX_COLUMNS; i++) {
        filter->headers[i].name = NULL;
        filter->headers[i].is_valid = 0;
    }

    filter->fp = fopen(filename, "w");
    if (!filter->fp) {
        free(filter);
        return NULL;
    }

    return filter;
}

// 시트명이 유효한 문자만 포함하는지 확인 (A-Z, a-z, 0-9, -, _, *)
int is_valid_name(const char* name) {
    for (int i = 0; name[i] != '\0'; i++) {
        char c = name[i];
        if (!((c >= 'A' && c <= 'Z') ||
              (c >= 'a' && c <= 'z') ||
              (c >= '0' && c <= '9') ||
              c == '-' || c == '_' || (ALLOW_WILD_CARD && c == '*'))) {
            return 0;  // 유효하지 않은 문자 발견
        }
    }

    return name[0] != '\0';  // 모든 문자가 유효함
}

void filter_close(Filter* filter) {
    // 버퍼에 남은 데이터 플러시
    if (filter->buffer_pos > 0) {
        fwrite(filter->line_buffer, 1, filter->buffer_pos, filter->fp);
        filter->buffer_pos = 0;
    }

    fclose(filter->fp);
    // 헤더 이름들 해제
    for (int i = 0; i < MAX_COLUMNS; i++) {
        if (filter->headers[i].name) {
            free((char*)filter->headers[i].name);
        }
    }
    free_string_list(filter->raw_fields, filter->raw_field_count);
    free_string_list(filter->json_headers, filter->json_header_count);
    free(filter);
}

// 문자열에서 * 문자 제거
void remove_wildcards(const char* input, char* output, int max_len) {
    int j = 0;
    for (int i = 0; input[i] && j < max_len - 1; i++) {
        if (input[i] != '*') {
            output[j++] = input[i];
        }
    }
    output[j] = '\0';
}

// 버퍼에 문자열을 추가하는 헬퍼 함수
static inline void append_to_buffer(Filter* filter, const char* str, int len) {
    // 버퍼가 넘칠 경우 먼저 플러시
    if (filter->buffer_pos + len >= LINE_BUFFER_SIZE) {
        fwrite(filter->line_buffer, 1, filter->buffer_pos, filter->fp);
        filter->buffer_pos = 0;

        // 문자열이 버퍼보다 큰 경우 직접 쓰기
        if (len >= LINE_BUFFER_SIZE) {
            fwrite(str, 1, len, filter->fp);
            return;
        }
    }

    // 버퍼에 추가
    memcpy(filter->line_buffer + filter->buffer_pos, str, len);
    filter->buffer_pos += len;
}

static void append_csv_field(Filter* filter, const char* data) {
    int needs_quotes = 0;
    for (const char* pos = data; *pos; pos++) {
        if (*pos == '"' || *pos == ',' || *pos == '\n' || *pos == '\r') {
            needs_quotes = 1;
            break;
        }
    }

    if (!needs_quotes) {
        append_to_buffer(filter, data, strlen(data));
        return;
    }

    append_to_buffer(filter, "\"", 1);
    for (const char* pos = data; *pos; pos++) {
        if (*pos == '"') {
            append_to_buffer(filter, "\"\"", 2);
        } else {
            append_to_buffer(filter, pos, 1);
        }
    }
    append_to_buffer(filter, "\"", 1);
}

static void append_json_string(Filter* filter, const char* data) {
    append_to_buffer(filter, "\"", 1);
    for (const unsigned char* pos = (const unsigned char*)data; *pos; pos++) {
        switch (*pos) {
            case '\\':
                append_to_buffer(filter, "\\\\", 2);
                break;
            case '"':
                append_to_buffer(filter, "\\\"", 2);
                break;
            case '\b':
                append_to_buffer(filter, "\\b", 2);
                break;
            case '\f':
                append_to_buffer(filter, "\\f", 2);
                break;
            case '\n':
                append_to_buffer(filter, "\\n", 2);
                break;
            case '\r':
                append_to_buffer(filter, "\\r", 2);
                break;
            case '\t':
                append_to_buffer(filter, "\\t", 2);
                break;
            default:
                if (*pos < 0x20) {
                    char escaped[7];
                    snprintf(escaped, sizeof(escaped), "\\u%04x", *pos);
                    append_to_buffer(filter, escaped, strlen(escaped));
                } else {
                    append_to_buffer(filter, (const char*)pos, 1);
                }
                break;
        }
    }
    append_to_buffer(filter, "\"", 1);
}

static inline void filter_push_raw(Filter* filter, const char* data) {
    ensure_raw_field_capacity(filter, filter->raw_field_count + 1);
    filter->raw_fields[filter->raw_field_count] = strdup(data);
    if (!filter->raw_fields[filter->raw_field_count]) {
        printf("Error: Memory allocation failed\n");
        exit(1);
    }
    filter->raw_field_count++;
    filter->col_count++;
}

void filter_push(Filter* filter, const char* data) {
    if (!filter || !data) {
        return;
    }

    if (filter->mode == FILTER_MODE_RAW) {
        filter_push_raw(filter, data);
        return;
    }

    if (filter->col_count >= MAX_COLUMNS) {
        printf("Error: Too many columns\n");
        exit(1);
    }

    if (filter->row_count == 0) {
        filter->headers[filter->col_count].name = strdup(data);
        if (!filter->headers[filter->col_count].name) {
            printf("Error: Memory allocation failed\n");
            exit(1);
        }
        filter->headers[filter->col_count].is_valid = is_valid_name(data);
    }

    if (filter->headers[filter->col_count].is_valid) {
        if (filter->valid_col_count > 0) {
            append_to_buffer(filter, "\t", 1);
        }

        // 헤더 행(첫 번째 행)에서만 * 문자 제거
        if (filter->row_count == 0) {
            char cleaned_data[MAX_COLUMNS * 10];  // 충분한 버퍼 크기
            remove_wildcards(data, cleaned_data, sizeof(cleaned_data));
            int len = strlen(cleaned_data);
            append_to_buffer(filter, cleaned_data, len);
        } else {
            int len = strlen(data);
            append_to_buffer(filter, data, len);
        }

        filter->valid_col_count++;
    }

    filter->col_count++;
}

void filter_finish_line(Filter* filter) {
    if (filter->mode == FILTER_MODE_RAW) {
        if (filter->output_format == OUTPUT_FORMAT_JSONL) {
            if (filter->row_count == 0) {
                capture_json_headers(filter);
            } else {
                int field_count = filter->raw_field_count;
                if (field_count < filter->json_header_count) {
                    field_count = filter->json_header_count;
                } else {
                    ensure_json_header_names(filter, field_count);
                }

                append_to_buffer(filter, "{", 1);
                for (int i = 0; i < field_count; i++) {
                    const char* value = i < filter->raw_field_count ? filter->raw_fields[i] : "";
                    if (i > 0) {
                        append_to_buffer(filter, ",", 1);
                    }
                    append_json_string(filter, filter->json_headers[i]);
                    append_to_buffer(filter, ":", 1);
                    append_json_string(filter, value);
                }
                append_to_buffer(filter, "}\n", 2);
            }
        } else {
            for (int i = 0; i < filter->raw_field_count; i++) {
                if (i > 0) {
                    if (filter->output_format == OUTPUT_FORMAT_CSV) {
                        append_to_buffer(filter, ",", 1);
                    } else {
                        append_to_buffer(filter, "\t", 1);
                    }
                }

                if (filter->output_format == OUTPUT_FORMAT_CSV) {
                    append_csv_field(filter, filter->raw_fields[i]);
                } else {
                    append_to_buffer(filter, filter->raw_fields[i], strlen(filter->raw_fields[i]));
                }
            }
            append_to_buffer(filter, "\n", 1);
        }

        reset_raw_fields(filter);
        filter->row_count++;
        filter->col_count = 0;
        filter->valid_col_count = 0;
        return;
    }

    append_to_buffer(filter, "\n", 1);
    filter->row_count++;
    filter->col_count = 0;
    filter->valid_col_count = 0;
}
