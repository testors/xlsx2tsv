# xlsx2tsv

XLSX 파일을 TSV/CSV/JSONL로 빠르게 변환합니다.

- 기본값은 범용 `generic` 모드입니다.
- 게임 DB 전용 고속 경로는 `--mode game-db-fast`로 분리되어 있습니다.
- LLM 에이전트나 wrapper에서 쓰기 좋은 `--list-sheets`, `--manifest-*`, `--stdout`, `--quiet`를 지원합니다.

## Install

```bash
./install.sh
```

- 기본 설치 위치는 `$HOME/.local/bin/xlsx_to_tsv`
- `--prefix /path` 또는 `--bin-dir /path/to/bin`으로 설치 위치를 바꿀 수 있습니다.

예:

```bash
./install.sh --bin-dir "$HOME/bin"
```

## Quick Start

### 1. 모든 시트를 TSV로 변환

```bash
./xlsx_to_tsv report.xlsx
```

- 기본 모드인 `generic`으로 모든 worksheet를 출력합니다.
- `start_row` 이전 행은 건너뛰고, 그 행부터 그대로 출력합니다.

### 2. 구조만 먼저 확인

```bash
./xlsx_to_tsv report.xlsx --list-sheets --json --quiet
```

- workbook 구조를 JSON으로 빠르게 확인합니다.
- 에이전트나 자동화에서 시트 선택 전에 쓰기 좋습니다.

### 3. 특정 시트만 포맷 적용 후 내보내기

```bash
mkdir -p out && ./xlsx_to_tsv report.xlsx 2 \
  --sheet Summary \
  --formatted \
  --output-dir out \
  --manifest-json out/manifest.json
```

- `Summary` 시트만 선택합니다.
- 날짜/시간/퍼센트 같은 숫자를 사람이 읽는 문자열로 바꿉니다.
- 결과 파일과 manifest JSON을 함께 생성합니다.

### 4. 단일 시트를 stdout으로 스트리밍

```bash
./xlsx_to_tsv report.xlsx 2 --sheet Summary --stdout --quiet
```

- stdout에는 시트 데이터만 출력됩니다.
- progress 로그는 숨기고, fatal error는 stderr에 유지합니다.

### 5. 게임 DB 전용 고속 변환

```bash
mkdir -p out && ./xlsx_to_tsv data.xlsx 4 --mode game-db-fast --output-dir out
```

- 기존 게임 DB 규칙에 맞는 시트와 컬럼만 처리합니다.
- `start_row` 행을 헤더 행으로 사용합니다.

## Usage

```bash
./xlsx_to_tsv <input.xlsx> [start_row] [options]
./xlsx_to_tsv --version
```

- `start_row` 기본값은 `1`입니다.
- `generic` 모드에서는 `start_row` 이전 행을 버리고 그 행부터 출력합니다.
- `game-db-fast` 모드에서는 `start_row` 행을 헤더 행으로 사용합니다.
- 전체 CLI usage는 `./xlsx_to_tsv`만 실행하면 볼 수 있습니다.

## Modes

### 기본값: `generic`

- 모든 worksheet를 export합니다.
- 시트명/헤더명 유효성 검사로 시트를 스킵하지 않습니다.
- 헤더의 `*`를 포함한 텍스트를 그대로 유지합니다.
- 파일명만 안전한 형태로 정리합니다.
- 업무용 XLSX, 범용 변환, LLM 입력 준비에 권장합니다.

### `--mode game-db-fast`

- 기존 게임 DB 전용 고속 경로를 사용합니다.
- 유효한 시트명/헤더명만 처리합니다.
- 기본 동작에서는 헤더/시트명의 `*`를 제거합니다.
- `--no-wildcard`를 주면 `*`가 포함된 시트/컬럼을 아예 제외합니다.

## Common Options

### Core

- `--version`: 안정적인 버전 문자열을 출력하고 종료합니다.
- `--quiet`: progress/info/success summary를 숨기고 데이터 출력과 fatal error만 남깁니다.
- `--output-dir dir`: 기존 디렉터리 아래에 출력 파일을 생성합니다.
  - 디렉터리는 미리 존재해야 합니다.

### Sheet Selection

- `--sheet name`: exact match로 시트를 선택합니다. 반복 가능합니다.
- `--sheet-regex pattern`: POSIX regex로 시트를 선택합니다.
- `--fail-if-no-sheet`: 선택 결과가 비면 즉시 실패합니다.
- `--list-sheets`: workbook 구조만 확인하고 변환하지 않습니다.
- `--json`: `--list-sheets`와 함께 JSON으로 출력합니다.

### Output

- `--stdout`: 선택된 단일 시트의 데이터만 stdout으로 출력합니다.
- `--manifest-json path`: 변환 후 manifest JSON을 파일로 기록합니다.
- `--manifest-stdout`: 변환 후 manifest JSON을 stdout으로 출력합니다.
- `--csv`: generic 모드에서 `.csv`로 출력합니다.
- `--jsonl`: generic 모드에서 `.jsonl`로 출력합니다.
  - 첫 번째 출력 행을 key로 사용합니다.
- `--fail-on-output-collision`: 파일명 충돌 시 자동 suffix 대신 즉시 실패합니다.

### Resource Guards

- `--max-sheets n`: 선택된 시트를 앞에서부터 `n`개까지만 처리합니다.
- `--max-rows-per-sheet n`: 각 시트의 출력 행 수를 제한합니다.
- `--max-output-bytes n`: 전체 출력 바이트 수 상한을 둡니다.
- `--fail-if-truncated`: 상한에 걸려 출력이 잘리면 exit code를 `1`로 만듭니다.

### Generic-only

- `--formatted`: `styles.xml` 숫자 포맷을 적용합니다.
- `--expand-merged`: merged cell을 top-left 값으로 확장합니다.
- `--skip-hidden`: hidden row/column을 제외합니다.
- `--all-sheets`: generic 모드의 legacy alias입니다.

### Game-db-fast-only

- `--no-wildcard`: `*`가 포함된 시트와 컬럼을 제외합니다.

## More Examples

### generic + merged/hidden 보정

```bash
./xlsx_to_tsv report.xlsx 2 --formatted --expand-merged --skip-hidden
```

### generic + CSV

```bash
mkdir -p out && ./xlsx_to_tsv report.xlsx 2 --formatted --csv --output-dir out
```

### generic + JSONL

```bash
./xlsx_to_tsv report.xlsx 2 --formatted --jsonl
```

### bounded export

```bash
mkdir -p out && ./xlsx_to_tsv report.xlsx 2 \
  --output-dir out \
  --max-sheets 3 \
  --max-rows-per-sheet 1000 \
  --max-output-bytes 10485760 \
  --fail-if-truncated
```

### game-db-fast + strict wildcard

```bash
./xlsx_to_tsv data.xlsx 4 --mode game-db-fast --no-wildcard
```

## Advanced Notes

### Output Files

- 기본 출력 확장자는 `.tsv`입니다.
- `--csv`는 `.csv`, `--jsonl`은 `.jsonl`을 사용합니다.
- `--output-dir`를 주지 않으면 현재 작업 디렉터리에 출력합니다.
- 파일명은 시트명 기반이며, 안전하지 않은 문자와 공백은 `_`로 정리합니다.
- 정리된 파일명이 겹치면 기본적으로 `__2`, `__3` 같은 suffix를 붙입니다.
- `--fail-on-output-collision`을 주면 suffix를 붙이지 않고 즉시 실패합니다.

### Generic Formatting

- 날짜 예: `45293` -> `2024-01-02`
- 시간 예: `0.5` -> `12:00`
- 날짜시간 예: `45293.6278...` -> `2024-01-02 15:04:05`
- 퍼센트 예: `0.125` -> `12.50%`
- zero-padding 예: `42` + `00000` 형식 -> `00042`
- boolean 예: `1`, `0` -> `TRUE`, `FALSE`

### Wildcard Behavior

- `generic` 모드는 `*`를 시트/헤더 텍스트에서 그대로 유지합니다.
- `game-db-fast` 모드는 기본적으로 `*`를 출력명에서 제거합니다.
- `game-db-fast`에서 `--no-wildcard`를 주면 `*`가 포함된 시트와 컬럼을 제외합니다.

### Valid Characters in `game-db-fast`

- 시트명 허용 문자: `A-Z`, `a-z`, `0-9`, `-`, `_`, `*`
- 헤더명 허용 문자: `A-Z`, `a-z`, `0-9`, `-`, `_`, `*`
- 공백, 한글, 기타 특수문자가 포함되면 해당 시트/컬럼은 제외될 수 있습니다.
- `generic` 모드에서는 이 제한을 적용하지 않습니다.

### Option Constraints

- `--manifest-stdout`는 `--manifest-json`, `--stdout`과 함께 쓸 수 없습니다.
- `--stdout`는 `--output-dir`, `--fail-on-output-collision`과 함께 쓸 수 없습니다.
- `--list-sheets`는 조회 전용 모드이며 변환 전용 옵션과 함께 쓸 수 없습니다.
- `--no-wildcard`는 `--mode game-db-fast`와 함께 써야 합니다.
- `--stdout`는 정확히 1개 시트가 선택될 때만 허용됩니다.

### Truncation Behavior

- `--max-output-bytes`는 현재 시트를 중간에서 자를 수 있습니다.
- 출력 바이트 상한에 도달하면 이후 선택된 시트는 skip될 수 있습니다.
- 이런 정보는 manifest의 top-level `truncated`, top-level `warnings`, per-sheet `warnings`에 기록됩니다.

## Agent Integration

에이전트 통합용 실행 계약, JSON schema, stdout/stderr 규약은 [docs/agent-integration.md](/Users/testors/Repos/xlsx2tsv/docs/agent-integration.md)에 따로 정리되어 있습니다.
