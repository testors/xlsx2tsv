#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <time.h>
#include <stdbool.h>
#include <stdint.h>

#include "miniz.h"
#include "filter.h"

// *** xlsx_to_tsv

#define MAX_CELL_VALUE 32768
#define BUFFER_SIZE 65536
#define MAX_SHEET_NAME 256
#define MAX_SHEETS 50

// Shared strings structure for performance
typedef struct {
    char** strings;
    int count;
    int capacity;
} SharedStrings;

// Sheet information structure
typedef struct {
    char name[MAX_SHEET_NAME];
    char filename[MAX_SHEET_NAME];
    int sheet_id;
} SheetInfo;

// Workbook structure
typedef struct {
    SheetInfo sheets[MAX_SHEETS];
    int sheet_count;
} Workbook;

// Fast XML attribute finder
char* find_attribute(const char* xml, const char* attr_name) {
    char* pos = strstr(xml, attr_name);
    if (!pos) return NULL;
    
    pos += strlen(attr_name);
    
    // Skip whitespace
    while (*pos && (*pos == ' ' || *pos == '\t')) pos++;
    
    // Find opening quote
    if (*pos != '"') return NULL;
    pos++; // Skip opening quote
    
    // Find closing quote
    char* end = strchr(pos, '"');
    if (!end) return NULL;
    
    int len = end - pos;
    char* result = malloc(len + 1);
    strncpy(result, pos, len);
    result[len] = '\0';
    return result;
}

// Fast XML content extractor
char* extract_xml_content(const char* xml, const char* tag) {
    char start_tag[256];
    char end_tag[256];
    snprintf(start_tag, sizeof(start_tag), "<%s>", tag);
    snprintf(end_tag, sizeof(end_tag), "</%s>", tag);
    
    char* start = strstr(xml, start_tag);
    if (!start) return NULL;
    start += strlen(start_tag);
    
    char* end = strstr(start, end_tag);
    if (!end) return NULL;
    
    int len = end - start;
    char* result = malloc(len + 1);
    strncpy(result, start, len);
    result[len] = '\0';
    return result;
}

// Initialize shared strings
void init_shared_strings(SharedStrings* ss) {
    ss->capacity = 65535;
    ss->strings = malloc(sizeof(char*) * ss->capacity);
    ss->count = 0;
}

// Unescape XML entities in-place
void unescape_xml_entities(char* str) {
    char* src = str;
    char* dst = str;
    
    while (*src) {
        if (*src == '&') {
            if (strncmp(src, "&lt;", 4) == 0) {
                *dst++ = '<';
                src += 4;
            } else if (strncmp(src, "&gt;", 4) == 0) {
                *dst++ = '>';
                src += 4;
            } else if (strncmp(src, "&amp;", 5) == 0) {
                *dst++ = '&';
                src += 5;
            } else if (strncmp(src, "&quot;", 6) == 0) {
                *dst++ = '"';
                src += 6;
            } else if (strncmp(src, "&apos;", 6) == 0) {
                *dst++ = '\'';
                src += 6;
            } else {
                *dst++ = *src++;
            }
        } else {
            *dst++ = *src++;
        }
    }
    *dst = '\0';
}

// Add string to shared strings
void add_shared_string(SharedStrings* ss, const char* str) {
    if (ss->count >= ss->capacity) {
        ss->capacity *= 2;
        ss->strings = realloc(ss->strings, sizeof(char*) * ss->capacity);
    }
    
    ss->strings[ss->count] = malloc(strlen(str) + 1);
    strcpy(ss->strings[ss->count], str);
    
    // Unescape XML entities
    unescape_xml_entities(ss->strings[ss->count]);
    
    ss->count++;

    // Debug: Print first 30 shared strings
#ifdef DEBUG
    printf("DEBUG: Shared string [%d] = '%s'\n", ss->count - 1, ss->strings[ss->count - 1]);
#endif
}

// Parse shared strings XML by extracting text content and skipping all tags
void parse_shared_strings(const char* xml_data, SharedStrings* ss) {
    const char* pos = xml_data;
    
    // Find each <si> (shared string item) element
    while ((pos = strstr(pos, "<si")) != NULL) {
        // Check for self-closing tag <si/>
        const char* tag_end = strchr(pos, '>');
        if (!tag_end) break;
        
        if (*(tag_end - 1) == '/') {
            // Self-closing tag <si/> - represents empty string
            add_shared_string(ss, "");
            pos = tag_end + 1;
            continue;
        }
        
        // Regular <si>...</si> tag
        if (*(pos + 3) != '>') {
            // Skip if not exactly "<si>"
            pos++;
            continue;
        }
        
        pos += 4; // Skip <si>
        const char* si_end = strstr(pos, "</si>");
        if (!si_end) break;
        
        // Extract text content, skipping all tags
        char text_buffer[MAX_CELL_VALUE] = "";
        int buffer_pos = 0;
        const char* current = pos;
        int inside_tag = 0;
        
        while (current < si_end && buffer_pos < MAX_CELL_VALUE - 1) {
            if (*current == '<') {
                inside_tag = 1;  // Start of a tag
            } else if (*current == '>') {
                inside_tag = 0;  // End of a tag
            } else if (!inside_tag) {
                // Not inside a tag, add character to buffer
                text_buffer[buffer_pos++] = *current;
            }
            current++;
        }
        
        text_buffer[buffer_pos] = '\0';
        
        // Add to shared strings (including empty strings to maintain correct indexing)
        add_shared_string(ss, text_buffer);
        
        pos = si_end + 5; // Move past </si>
    }
}

// Parse workbook.xml to get sheet information
void parse_workbook(const char* xml_data, Workbook* wb) {
    wb->sheet_count = 0;
    const char* pos = xml_data;
    
    while ((pos = strstr(pos, "<sheet ")) != NULL && wb->sheet_count < MAX_SHEETS) {
        // Extract sheet name
        char* name_attr = find_attribute(pos, "name=");
        if (!name_attr) {
            pos++;
            continue;
        }
        
        // Extract sheet ID
        char* sheet_id_attr = find_attribute(pos, "sheetId=");
        int sheet_id = sheet_id_attr ? atoi(sheet_id_attr) : wb->sheet_count + 1;
        
        // Skip sheets with invalid characters (only allow A-Z, a-z, 0-9, -, _, *)
        if (!is_valid_name(name_attr)) {
            printf("Skipping sheet: '%s' (contains invalid characters - only A-Z, a-z, 0-9, -, _, * allowed)\n", name_attr);
            free(name_attr);
            if (sheet_id_attr) free(sheet_id_attr);
            pos++;
            continue;
        }
        
        // Store sheet information
        strncpy(wb->sheets[wb->sheet_count].name, name_attr, MAX_SHEET_NAME - 1);
        wb->sheets[wb->sheet_count].name[MAX_SHEET_NAME - 1] = '\0';
        wb->sheets[wb->sheet_count].sheet_id = sheet_id;
        
        // Generate worksheet filename using sequential order (not sheetId)
        // Excel file structure uses sheet1.xml, sheet2.xml, etc. in document order
        snprintf(wb->sheets[wb->sheet_count].filename, MAX_SHEET_NAME,
                "xl/worksheets/sheet%d.xml", wb->sheet_count + 1);

        wb->sheet_count++;
        
        free(name_attr);
        if (sheet_id_attr) free(sheet_id_attr);
        pos++;
    }
}

// Convert Excel column reference to number (A=0, B=1, etc.)
int col_ref_to_num(const char* ref) {
    int col = 0;
    for (int i = 0; ref[i] && isalpha(ref[i]); i++) {
        col = col * 26 + (toupper(ref[i]) - 'A' + 1);
    }
    return col - 1;
}

// Extract row number from cell reference
int extract_row_num(const char* ref) {
    while (*ref && isalpha(*ref)) ref++;
    return atoi(ref) - 1;
}

// Escape TSV special characters
void escape_tsv_value(const char* input, char* output, int max_len) {
    int i = 0, j = 0;
    while (input[i] && j < max_len - 1) {
        if (input[i] == '\t') {
            output[j++] = ' ';  // Replace tab with space
        } else if (input[i] == '\n' || input[i] == '\r') {
            output[j++] = ' ';  // Replace newlines with space
        } else {
            output[j++] = input[i];
        }
        i++;
    }
    output[j] = '\0';
}

// Create safe filename from sheet name
void create_safe_filename(const char* sheet_name, char* safe_name, int max_len) {
    int i = 0, j = 0;
    while (sheet_name[i] && j < max_len - 5) { // Reserve space for .tsv
        char c = sheet_name[i];
        // Replace unsafe characters with underscore
        if (c == '/' || c == '\\' || c == ':' || 
            c == '?' || c == '"' || c == '<' || c == '>' || c == '|') {
            safe_name[j++] = '_';
        } else if (c == ' ') {
            safe_name[j++] = '_';  // Replace spaces with underscores
        } else if (c == '*') {
            // Skip asterisk characters (don't add to filename)
            // Do nothing, just move to next character
        } else {
            safe_name[j++] = c;
        }
        i++;
    }
    // Add .tsv extension
    strcpy(safe_name + j, ".tsv");
}

// High-performance worksheet parser
void parse_worksheet(const char* xml_data, SharedStrings* ss, int start_row, Filter* output) {
    const char* pos = xml_data;
    int last_row = -1;
    int last_col = -1;
    
    while ((pos = strstr(pos, "<c ")) != NULL) {
        // Find the end of this cell tag to limit our search scope
        const char* cell_end = NULL;
        size_t cell_length = 0;
        
        // Check if it's a self-closing tag
        const char* tag_close = strchr(pos, '>');
        if (!tag_close) {
            pos++;
            continue;
        }
        
        if (*(tag_close - 1) == '/') {
            // Self-closing tag: <c ... />
            cell_end = tag_close + 1;
            cell_length = cell_end - pos;
        } else {
            // Regular tag: <c ...>...</c>
            cell_end = strstr(pos, "</c>");
            if (!cell_end) {
                pos++;
                continue;
            }
            cell_length = cell_end - pos;
            cell_end += 4; // Move past </c>
        }
        
        // Extract cell reference (search within this cell only)
        char* cell_content = malloc(cell_length + 1);
        strncpy(cell_content, pos, cell_length);
        cell_content[cell_length] = '\0';
        
        char* r_attr = find_attribute(cell_content, "r=");
        if (!r_attr) {
            free(cell_content);
            pos = cell_end;
            continue;
        }
        
                int row = extract_row_num(r_attr);
        int col = col_ref_to_num(r_attr);

        // Skip rows before start_row
        if (row < start_row) {
            free(cell_content);
            free(r_attr);
            pos = cell_end;
            continue;
        }
        
        // If we moved to a new row, output newline and reset column tracking
        if (last_row != -1 && row != last_row) {
            filter_finish_line(output);
#ifdef DEBUG
            printf("DEBUG: New row, outputting newline\n");
#endif
            last_col = -1;
        }
        
        // Fill empty columns with tabs (for columns between last_col and current col)
        int tabs_needed = col - last_col - 1;
        if (last_col >= 0) tabs_needed++; // Add one more tab to separate from previous cell
        
        for (int i = 0; i < tabs_needed; i++) {
            filter_push(output, "");
        }
        
        // Extract cell type from the same cell content
        char* t_attr = NULL;
        char* temp_t_attr = find_attribute(cell_content, "t=");
        if (temp_t_attr) {
            t_attr = malloc(strlen(temp_t_attr) + 1);
            strcpy(t_attr, temp_t_attr);
            free(temp_t_attr);
        }
        
        char cell_value[MAX_CELL_VALUE] = "";
        
        // Find cell value - handle different cell value formats (search within cell_content only)
        char* v_content = NULL;
        
        // Method 1: Look for <v> tag (for numeric values and shared string references)
        char* v_start = strstr(cell_content, "<v>");
        if (v_start) {
            v_start += 3; // Skip <v>
            char* v_end = strstr(v_start, "</v>");
            if (v_end) {
                int len = v_end - v_start;
                v_content = malloc(len + 1);
                strncpy(v_content, v_start, len);
                v_content[len] = '\0';
            }
        }
        
        // Method 2: Look for inline string <is><t> tag (for inline text)
        if (!v_content) {
            char* is_start = strstr(cell_content, "<is><t>");
            if (is_start) {
                is_start += 7; // Skip <is><t>
                char* is_end = strstr(is_start, "</t></is>");
                if (is_end) {
                    int len = is_end - is_start;
                    v_content = malloc(len + 1);
                    strncpy(v_content, is_start, len);
                    v_content[len] = '\0';
                }
            }
        }
        
        // Method 3: Look for simple <t> tag (for some text values)
        if (!v_content) {
            char* t_start = strstr(cell_content, "<t>");
            if (t_start) {
                t_start += 3; // Skip <t>
                char* t_end = strstr(t_start, "</t>");
                if (t_end) {
                    int len = t_end - t_start;
                    v_content = malloc(len + 1);
                    strncpy(v_content, t_start, len);
                    v_content[len] = '\0';
                }
            }
        }
        
        if (v_content) {
            // Handle different cell types
            if (t_attr && strcmp(t_attr, "s") == 0) {
                // Shared string reference
                int str_index = atoi(v_content);
                if (str_index >= 0 && str_index < ss->count) {
                    escape_tsv_value(ss->strings[str_index], cell_value, MAX_CELL_VALUE);
#ifdef DEBUG
                    printf("DEBUG: Cell %s [+%dtabs] '%s' : shared_string[%d] : '%s'\n", 
                           r_attr, tabs_needed, v_content, str_index, cell_value);
#endif
                }
            } else {
                // Numeric, inline string, or other value
                escape_tsv_value(v_content, cell_value, MAX_CELL_VALUE);
#ifdef DEBUG
                printf("DEBUG: Cell %s [+%dtabs] '%s' : direct_value : '%s'\n", 
                       r_attr, tabs_needed, v_content, cell_value);
#endif
            }
        } else {
#ifdef DEBUG
            printf("DEBUG: Cell %s [+%dtabs] : empty_cell : ''\n", r_attr, tabs_needed);
#endif
        }
        
        // Output cell value
        filter_push(output, cell_value);
        
        last_row = row;
        last_col = col;
        
        free(cell_content);
        if (r_attr) free(r_attr);
        if (t_attr) free(t_attr);
        if (v_content) free(v_content);
        
        // Move to next cell
        pos = cell_end;
    }
    
    // Output final newline if we processed any rows
    if (last_row >= start_row) {
        filter_finish_line(output);
    }
}

// Free shared strings memory
void free_shared_strings(SharedStrings* ss) {
    for (int i = 0; i < ss->count; i++) {
        free(ss->strings[i]);
    }
    free(ss->strings);
    ss->count = 0;
    ss->capacity = 0;
}

int main(int argc, char* argv[]) {
    if (argc < 2) {
        printf("Usage: %s <input.xlsx> [start_row] [--no-wildcard]\n", argv[0]);
        printf("  start_row: 1-based row number to start conversion (default: 1)\n");
        printf("\n");
        printf("Wildcard (*) character behavior:\n");
        printf("  Default mode:\n");
        printf("    - * characters are removed from sheet/column names in output\n");
        printf("    - Example: '*Sales' -> 'Sales.tsv', '*ID' column -> 'ID'\n");
        printf("  --no-wildcard mode:\n");
        printf("    - Sheets containing * will be skipped entirely\n");
        printf("    - Columns containing * will be excluded from output\n");
        printf("\n");
        printf("Note: Only A-Z, a-z, 0-9, -, _, * characters are valid in sheet/column names\n");
        printf("      Names with spaces, special characters, or non-ASCII will be skipped\n");
        return 1;
    }
    
    const char* input_file = argv[1];
    int start_row = (argc > 2) ? atoi(argv[2]) - 1 : 0;  // Convert to 0-based
    if (argc > 3 && strcmp(argv[3], "--no-wildcard") == 0) {
        ALLOW_WILD_CARD = false;
    }
    if (start_row < 0) start_row = 0;
    
    printf("Converting XLSX to multiple TSV files...\n");
    printf("Input: %s\n", input_file);
    printf("Starting from row: %d\n", start_row + 1);
    
    clock_t start_time = clock();
    
    // Open XLSX file
    mz_zip_archive zip;
    if (!mz_zip_reader_init_file(&zip, input_file)) {
        printf("Error: Could not open XLSX file: %s\n", input_file);
        return 1;
    }
    
    // Initialize workbook and parse sheet information
    Workbook workbook;
    int workbook_index;
    if (!mz_zip_reader_locate_file(&zip, "xl/workbook.xml", &workbook_index)) {
        printf("Error: Could not find workbook.xml in XLSX file\n");
        mz_zip_reader_end(&zip);
        return 1;
    }
    
    size_t workbook_size = mz_zip_reader_get_file_size(&zip, workbook_index);
    char* workbook_data = malloc(workbook_size + 1);
    
    if (!mz_zip_reader_extract_to_mem(&zip, workbook_index, workbook_data, workbook_size)) {
        printf("Error: Could not extract workbook.xml\n");
        free(workbook_data);
        mz_zip_reader_end(&zip);
        return 1;
    }
    
    workbook_data[workbook_size] = '\0';
    parse_workbook(workbook_data, &workbook);
    free(workbook_data);
    
    if (workbook.sheet_count == 0) {
        printf("No valid sheets found (sheets must contain only A-Z, a-z, 0-9, -, _, *)\n");
        mz_zip_reader_end(&zip);
        return 1;
    }
    
    printf("Found %d sheet(s) to process\n\n", workbook.sheet_count);
    
    // Initialize shared strings
    SharedStrings shared_strings;
    init_shared_strings(&shared_strings);
    
    // Extract and parse shared strings
    int shared_strings_index;
    if (mz_zip_reader_locate_file(&zip, "xl/sharedStrings.xml", &shared_strings_index)) {
        printf("Loading shared strings...\n");
        size_t shared_strings_size = mz_zip_reader_get_file_size(&zip, shared_strings_index);
        char* shared_strings_data = malloc(shared_strings_size + 1);
        
        if (mz_zip_reader_extract_to_mem(&zip, shared_strings_index, shared_strings_data, shared_strings_size)) {
            shared_strings_data[shared_strings_size] = '\0';
            parse_shared_strings(shared_strings_data, &shared_strings);
            printf("Loaded %d shared strings\n\n", shared_strings.count);
        }
        
        free(shared_strings_data);
    }
    
    // Process each sheet
    int processed_sheets = 0;
    for (int i = 0; i < workbook.sheet_count; i++) {
        printf("Processing sheet %d/%d: '%s'\n", i + 1, workbook.sheet_count, workbook.sheets[i].name);
        
        // Extract and parse worksheet
        int worksheet_index;
        if (!mz_zip_reader_locate_file(&zip, workbook.sheets[i].filename, &worksheet_index)) {
            printf("Warning: Could not find worksheet file: %s - skipping\n\n", workbook.sheets[i].filename);
            continue;
        }
        
        size_t worksheet_size = mz_zip_reader_get_file_size(&zip, worksheet_index);
        char* worksheet_data = malloc(worksheet_size + 1);
        
        if (!mz_zip_reader_extract_to_mem(&zip, worksheet_index, worksheet_data, worksheet_size)) {
            printf("Warning: Could not extract worksheet data for: %s - skipping\n\n", workbook.sheets[i].name);
            free(worksheet_data);
            continue;
        }
        
        worksheet_data[worksheet_size] = '\0';
        
        // Create safe output filename
        char output_filename[MAX_SHEET_NAME + 10];
        create_safe_filename(workbook.sheets[i].name, output_filename, sizeof(output_filename));
        
        // Open output file
        Filter* output = filter_init(output_filename);
        if (!output) {
            printf("Warning: Could not create output file: %s - skipping\n\n", output_filename);
            free(worksheet_data);
            continue;
        }
        
        printf("  Output file: %s\n", output_filename);
        
        // Parse worksheet and generate TSV
        parse_worksheet(worksheet_data, &shared_strings, start_row, output);
        
        // Cleanup for this sheet
        filter_close(output);
        free(worksheet_data);
        processed_sheets++;
        
        printf("  Sheet '%s' processed successfully!\n\n", workbook.sheets[i].name);
    }
    mz_zip_reader_end(&zip);
    free_shared_strings(&shared_strings);
    
    clock_t end_time = clock();
    double elapsed = ((double)(end_time - start_time)) / CLOCKS_PER_SEC;
    
    printf("=== Conversion Summary ===\n");
    printf("Total sheets processed: %d out of %d\n", processed_sheets, workbook.sheet_count);
    printf("Processing time: %.2f seconds\n", elapsed);
    
    if (processed_sheets > 0) {
        printf("Conversion completed successfully!\n");
        printf("Output files created:\n");
        for (int i = 0; i < workbook.sheet_count; i++) {
            char output_filename[MAX_SHEET_NAME + 10];
            create_safe_filename(workbook.sheets[i].name, output_filename, sizeof(output_filename));
            printf("  - %s (from sheet: %s)\n", output_filename, workbook.sheets[i].name);
        }
        return 0;
    } else {
        printf("No sheets were processed successfully.\n");
        return 1;
    }
} 

// *** xlsx_to_tsv END
