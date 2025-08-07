CC = gcc
#CFLAGS = -O3 -Wall -Wextra -march=native -flto
CFLAGS = -Wall -Wextra -march=native -flto -g
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
	@echo "Usage: ./$(TARGET) input.xlsx output.tsv [start_row]"

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
