#pragma once

#include <stdbool.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#define MAX_COLUMNS 1000
#define LINE_BUFFER_SIZE 65536

typedef enum {
    FILTER_MODE_GAME_DB,
    FILTER_MODE_RAW
} FilterMode;

typedef enum {
    OUTPUT_FORMAT_TSV,
    OUTPUT_FORMAT_CSV,
    OUTPUT_FORMAT_JSONL
} OutputFormat;

typedef struct {

    struct {
        char* name;
        bool is_valid;
    } headers[MAX_COLUMNS];

    FILE *fp;
    FilterMode mode;
    OutputFormat output_format;
    int col_count;
    int valid_col_count;
    int row_count;

    char** raw_fields;
    int raw_field_count;
    int raw_field_capacity;

    char** json_headers;
    int json_header_count;
    int json_header_capacity;

    // I/O 성능 향상을 위한 라인 버퍼링
    char line_buffer[LINE_BUFFER_SIZE];
    int buffer_pos;
} Filter;

Filter* filter_init(const char* filename, FilterMode mode, OutputFormat output_format);
void filter_close(Filter* filter);
void filter_push(Filter* filter, const char* data);
void filter_finish_line(Filter* filter);
int is_valid_name(const char* name);

extern bool ALLOW_WILD_CARD;
