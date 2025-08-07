#pragma once

// *** MINIZ
#include <zlib.h>

// ZIP file structures
#pragma pack(push, 1)
typedef struct {
    uint32_t signature;
    uint16_t version;
    uint16_t flags;
    uint16_t method;
    uint16_t time;
    uint16_t date;
    uint32_t crc32;
    uint32_t comp_size;
    uint32_t uncomp_size;
    uint16_t name_len;
    uint16_t extra_len;
} mz_zip_local_file_header;

typedef struct {
    uint32_t signature;
    uint16_t version_made_by;
    uint16_t version_needed;
    uint16_t flags;
    uint16_t method;
    uint16_t time;
    uint16_t date;
    uint32_t crc32;
    uint32_t comp_size;
    uint32_t uncomp_size;
    uint16_t name_len;
    uint16_t extra_len;
    uint16_t comment_len;
    uint16_t disk_start;
    uint16_t internal_attr;
    uint32_t external_attr;
    uint32_t local_header_offset;
} mz_zip_central_dir_entry;

typedef struct {
    uint32_t signature;
    uint16_t disk_num;
    uint16_t central_dir_disk;
    uint16_t entries_this_disk;
    uint16_t total_entries;
    uint32_t central_dir_size;
    uint32_t central_dir_offset;
    uint16_t comment_len;
} mz_zip_end_central_dir;
#pragma pack(pop)

typedef struct {
    FILE* file;
    uint32_t total_entries;
    uint32_t central_dir_offset;
    mz_zip_central_dir_entry* entries;
} mz_zip_archive;

// Function declarations
int mz_zip_reader_init_file(mz_zip_archive* zip, const char* filename);
void mz_zip_reader_end(mz_zip_archive* zip);
int mz_zip_reader_locate_file(mz_zip_archive* zip, const char* name, int* file_index);
int mz_zip_reader_extract_to_mem(mz_zip_archive* zip, int file_index, void* buf, size_t buf_size);
size_t mz_zip_reader_get_file_size(mz_zip_archive* zip, int file_index);

// Implementation
int mz_zip_reader_init_file(mz_zip_archive* zip, const char* filename) {
    zip->file = fopen(filename, "rb");
    if (!zip->file) return 0;
    
    // Find end of central directory
    fseek(zip->file, -22, SEEK_END);
    mz_zip_end_central_dir end_dir;
    fread(&end_dir, sizeof(end_dir), 1, zip->file);
    
    if (end_dir.signature != 0x06054b50) {
        fclose(zip->file);
        return 0;
    }
    
    zip->total_entries = end_dir.total_entries;
    zip->central_dir_offset = end_dir.central_dir_offset;
    
    // Allocate space for central directory entries
    zip->entries = malloc(sizeof(mz_zip_central_dir_entry) * zip->total_entries);
    
    // Read central directory entries with proper parsing
    long current_pos = zip->central_dir_offset;
    for (uint32_t i = 0; i < zip->total_entries; i++) {
        fseek(zip->file, current_pos, SEEK_SET);
        fread(&zip->entries[i], sizeof(mz_zip_central_dir_entry), 1, zip->file);
        
        // Verify signature
        if (zip->entries[i].signature != 0x02014b50) {
            printf("Warning: Invalid central directory entry signature at index %d\n", i);
        }
        
        current_pos += sizeof(mz_zip_central_dir_entry) + 
                       zip->entries[i].name_len + 
                       zip->entries[i].extra_len + 
                       zip->entries[i].comment_len;
    }
    
    return 1;
}

void mz_zip_reader_end(mz_zip_archive* zip) {
    if (zip->file) fclose(zip->file);
    if (zip->entries) free(zip->entries);
    memset(zip, 0, sizeof(*zip));
}

int mz_zip_reader_locate_file(mz_zip_archive* zip, const char* name, int* file_index) {
    long current_pos = zip->central_dir_offset;
    
    for (uint32_t i = 0; i < zip->total_entries; i++) {
        fseek(zip->file, current_pos, SEEK_SET);
        
        mz_zip_central_dir_entry entry;
        fread(&entry, sizeof(entry), 1, zip->file);
        
        if (entry.signature != 0x02014b50) {
            return 0; // Invalid central directory signature
        }
        
        char* filename = malloc(entry.name_len + 1);
        fread(filename, entry.name_len, 1, zip->file);
        filename[entry.name_len] = '\0';
        
        if (strcmp(filename, name) == 0) {
            *file_index = i;
            free(filename);
            return 1;
        }
        
        free(filename);
        current_pos += sizeof(mz_zip_central_dir_entry) + entry.name_len + entry.extra_len + entry.comment_len;
    }
    return 0;
}

size_t mz_zip_reader_get_file_size(mz_zip_archive* zip, int file_index) {
    return zip->entries[file_index].uncomp_size;
}

int mz_zip_reader_extract_to_mem(mz_zip_archive* zip, int file_index, void* buf, size_t buf_size) {
    mz_zip_central_dir_entry* entry = &zip->entries[file_index];
    
    fseek(zip->file, entry->local_header_offset, SEEK_SET);
    mz_zip_local_file_header local_header;
    fread(&local_header, sizeof(local_header), 1, zip->file);
    
    // Skip filename and extra field
    fseek(zip->file, local_header.name_len + local_header.extra_len, SEEK_CUR);
    
    if (entry->method == 0) {
        // Stored (no compression)
        fread(buf, entry->uncomp_size, 1, zip->file);
        return 1;
    } else if (entry->method == 8) {
        // Deflate compression - use zlib
        char* comp_data = malloc(entry->comp_size);
        fread(comp_data, entry->comp_size, 1, zip->file);
        
        z_stream strm = {0};
        strm.next_in = (Bytef*)comp_data;
        strm.avail_in = entry->comp_size;
        strm.next_out = (Bytef*)buf;
        strm.avail_out = buf_size;
        
        // Initialize inflateInit2 with negative window bits for raw deflate
        if (inflateInit2(&strm, -MAX_WBITS) != Z_OK) {
            free(comp_data);
            return 0;
        }
        
        int result = inflate(&strm, Z_FINISH);
        inflateEnd(&strm);
        free(comp_data);
        
        return (result == Z_STREAM_END) ? 1 : 0;
    } else {
        // Unsupported compression method
        return 0;
    }
}
//*** MINIZ END