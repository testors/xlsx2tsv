// *** FILTER
//#define _GNU_SOURCE  // GNU 확장 기능 활성화
//#define _POSIX_C_SOURCE 200809L  // POSIX.1-2008 기능 활성화

#include "filter.h"

// strdup 함수 프로토타입 명시적 선언
//char* strdup(const char* s);

/*
Filter* filter_init(const char* filename);
void filter_close(Filter* filter);
void filter_push(Filter* filter, const char* data);
void filter_finish_line(Filter* filter);
int is_valid_name(const char* name);
*/

bool ALLOW_WILD_CARD = true;

Filter* filter_init(const char* filename) {
    Filter* filter = (Filter*)malloc(sizeof(Filter));
    if (!filter) {
        return NULL;
    }
    
    filter->col_count = 0;
    filter->row_count = 0;
    filter->valid_col_count = 0;
    
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

// Check if sheet name contains only valid characters (A-Z, a-z, 0-9, -, _, *)
int is_valid_name(const char* name) {
    for (int i = 0; name[i] != '\0'; i++) {
        char c = name[i];
        if (!((c >= 'A' && c <= 'Z') || 
              (c >= 'a' && c <= 'z') || 
              (c >= '0' && c <= '9') || 
              c == '-' || c == '_' || (ALLOW_WILD_CARD && c == '*'))) {
            return 0;  // Invalid character found
        }
    }

    return name[0] != '\0';  // All characters are valid
}

void filter_close(Filter* filter) {
    fclose(filter->fp);
    // 헤더 이름들 해제
    for (int i = 0; i < MAX_COLUMNS; i++) {
        if (filter->headers[i].name) {
            free((char*)filter->headers[i].name);
        }
    }
    free(filter);
}

// Remove * characters from string
void remove_wildcards(const char* input, char* output, int max_len) {
    int j = 0;
    for (int i = 0; input[i] && j < max_len - 1; i++) {
        if (input[i] != '*') {
            output[j++] = input[i];
        }
    }
    output[j] = '\0';
}

void filter_push(Filter* filter, const char* data) {
    if (!filter || !data) {
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
            fputc('\t', filter->fp);
        }

        // Remove * characters only from header row (first row)
        if (filter->row_count == 0) {
            char cleaned_data[MAX_COLUMNS * 10];  // Sufficient buffer size
            remove_wildcards(data, cleaned_data, sizeof(cleaned_data));
            fputs(cleaned_data, filter->fp);
        } else {
            fputs(data, filter->fp);
        }

        filter->valid_col_count++;
    }

    filter->col_count++;
}

void filter_finish_line(Filter* filter) {
    fputc('\n', filter->fp);
    filter->row_count++;
    filter->col_count = 0;
    filter->valid_col_count = 0;
}
// *** FILTER END