CC = gcc
CFLAGS = -O3 -DNDEBUG -Wall -Wextra -march=native -flto
LDFLAGS = -lz
TARGET = xlsx_to_tsv
SOURCES = xlsx_to_tsv.c filter.c

.PHONY: all clean test

all: $(TARGET) miniz.h filter.h

$(TARGET): $(SOURCES)
	$(CC) $(CFLAGS) -o $(TARGET) $(SOURCES) $(LDFLAGS)

clean:
	rm -f $(TARGET)

test: $(TARGET)
	@echo "Build completed successfully!"
	@echo "Usage: ./$(TARGET) input.xlsx [start_row] [--mode generic|game-db-fast] [--output-dir dir] [--sheet name] [--sheet-regex pattern] [--list-sheets] [--json] [--manifest-json path] [--manifest-stdout] [--stdout] [--max-sheets n] [--max-rows-per-sheet n] [--max-output-bytes n] [--fail-if-truncated] [--fail-if-no-sheet] [--fail-on-output-collision] [--no-wildcard] [--formatted] [--expand-merged] [--skip-hidden] [--csv] [--jsonl]"

install: $(TARGET)
	cp $(TARGET) /usr/local/bin/

.PHONY: help
help:
	@echo "Available targets:"
	@echo "  all     - Build the xlsx_to_tsv converter"
	@echo "  clean   - Remove built files"
	@echo "  test    - Build and show usage"
	@echo "  install - Install to /usr/local/bin"
	@echo "  help    - Show this help message" 
