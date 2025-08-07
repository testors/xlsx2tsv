#pragma once

//#define _GNU_SOURCE        // GNU 확장 기능 (strdup 포함)
//#define _POSIX_C_SOURCE 200809L  // POSIX.1-2008

#include <stdbool.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#define MAX_COLUMNS 1000

typedef struct {
    
    struct {
        char* name;
        bool is_valid;
    } headers[MAX_COLUMNS];

    FILE *fp;
    int col_count;
    int valid_col_count;
    int row_count;
} Filter;

Filter* filter_init(const char* filename);
void filter_close(Filter* filter);
void filter_push(Filter* filter, const char* data);
void filter_finish_line(Filter* filter);
int is_valid_name(const char* name);

extern bool ALLOW_WILD_CARD;