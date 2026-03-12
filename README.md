# xlsx2tsv
xlsx 파일을 TSV/CSV/JSONL로 고속 변환합니다.

## Usage
```bash
./xlsx_to_tsv <input.xlsx> [start_row] [--mode generic|game-db-fast] [--output-dir dir] [--sheet name] [--sheet-regex pattern] [--list-sheets] [--json] [--manifest-json path] [--manifest-stdout] [--stdout] [--max-sheets n] [--max-rows-per-sheet n] [--max-output-bytes n] [--fail-if-truncated] [--fail-if-no-sheet] [--fail-on-output-collision] [--no-wildcard] [--formatted] [--expand-merged] [--skip-hidden] [--csv] [--jsonl]
```

## Install
```bash
./install.sh
```

- 기본 설치 위치는 `$HOME/.local/bin/xlsx_to_tsv`
- `--prefix /path` 또는 `--bin-dir /path/to/bin`으로 변경 가능
- 예시:
```bash
./install.sh --bin-dir "$HOME/bin"
```

## Modes

### 기본값: generic
- 모든 worksheet를 export합니다.
- 시트명/헤더명 유효성 검사로 시트를 스킵하지 않습니다.
- `start_row`는 "출력을 시작할 행" 의미입니다.
- 헤더의 `*`를 포함한 텍스트를 그대로 유지합니다.
- 파일명만 안전한 형태로 정리합니다.

### `--mode game-db-fast`
- 기존 게임 DB 전용 고속 경로를 사용합니다.
- 유효한 시트명/헤더명만 처리합니다.
- `start_row` 행을 헤더 행으로 사용합니다.
- 기본 동작에서는 헤더/시트명의 `*`를 제거합니다.
- `--no-wildcard`와 함께 쓰면 `*`가 포함된 시트/컬럼을 아예 제외합니다.

## Parameters
- `input.xlsx`: 변환할 XLSX 파일 경로
- `start_row`: 1-based 시작 행 번호, 기본값은 `1`
  - generic 모드: `start_row` 이전 행을 무시하고 그 행부터 그대로 출력
  - `game-db-fast` 모드: `start_row` 행을 헤더 행으로 사용
- `--mode generic|game-db-fast`: export 모드 선택, 기본값은 `generic`
- `--output-dir dir`: 기존 디렉터리 아래에 출력 파일 생성
  - 디렉터리는 미리 존재해야 합니다.
- `--sheet name`: 정확한 sheet name으로 선택, 반복 가능
- `--sheet-regex pattern`: POSIX regex로 sheet 선택
- `--list-sheets`: workbook 구조만 빠르게 확인하고 변환하지 않음
- `--json`: `--list-sheets`와 함께 JSON 출력
- `--manifest-json path`: 변환 결과 manifest JSON을 파일로 저장
- `--manifest-stdout`: 변환 결과 manifest JSON을 stdout으로 출력
- `--stdout`: 선택된 단일 시트의 데이터만 stdout으로 출력
- `--all-sheets`: generic 모드의 legacy alias
- `--max-sheets n`: 선택된 시트 수를 앞에서부터 `n`개로 제한
- `--max-rows-per-sheet n`: 시트별 출력 행 수 제한
- `--max-output-bytes n`: 전체 출력 바이트 수 제한
- `--fail-if-truncated`: 리소스 가드로 출력이 잘리면 비정상 종료
- `--fail-if-no-sheet`: 시트 필터 결과가 비면 비정상 종료
- `--fail-on-output-collision`: 파일명 충돌 시 suffix 없이 바로 실패
- `--no-wildcard`: `game-db-fast` 전용 옵션
- `--formatted`: generic 전용, `styles.xml` 숫자 포맷 적용
- `--expand-merged`: generic 전용, merged cell을 top-left 값으로 확장
- `--skip-hidden`: generic 전용, hidden row/column 제외
- `--csv`: generic 전용, `.csv` 출력
- `--jsonl`: generic 전용, `.jsonl` 출력
  - 첫 번째 출력 행을 key로 사용하고 이후 행을 object로 출력
- `--version`: 안정적인 버전 문자열 출력 후 종료

## Generic Features

### `--formatted`
- 날짜 예: `45293` -> `2024-01-02`
- 시간 예: `0.5` -> `12:00`
- 날짜시간 예: `45293.6278...` -> `2024-01-02 15:04:05`
- 퍼센트 예: `0.125` -> `12.50%`
- zero-padding 예: `42` + `00000` 형식 -> `00042`
- boolean 예: `1`, `0` -> `TRUE`, `FALSE`

### `--expand-merged`
- merge 영역을 top-left 값으로 채웁니다.
- row XML이 없는 병합 하단 행도 필요하면 synthetic row로 출력할 수 있습니다.

### `--skip-hidden`
- `<row hidden="1">` 행을 제외합니다.
- `<col hidden="1">` 범위의 열을 제외합니다.

### `--csv` / `--jsonl`
- `--csv`: 각 시트를 `.csv`로 출력
- `--jsonl`: 각 시트를 `.jsonl`로 출력
- JSONL은 첫 번째 출력 행을 key로 사용하므로 보통 header 행을 `start_row`에 맞추는 편이 적절합니다.

## Agent Features

### `--version`
- health check와 wrapper 호환성 판별용 버전 문자열을 출력합니다.
- 출력 형식: `xlsx2tsv 0.1.0`

### `--list-sheets --json`
- 변환 전에 workbook 구조를 빠르게 확인합니다.
- 각 시트의 `sheet_name`, `state`, `hidden`, `selected`, `approx_rows`, `approx_cols`를 JSON으로 출력합니다.
- `--sheet` / `--sheet-regex`와 함께 쓰면 어떤 시트가 선택되는지 미리 볼 수 있습니다.

### `--manifest-json` / `--manifest-stdout`
- 변환 후 실행 결과를 JSON으로 기록합니다.
- 포함 정보:
  - 원본 파일
  - mode / output format / start row
  - selected / processed sheet count
  - 실제 출력 파일 경로
  - 시트별 emitted row/col 수
  - 시트별 warnings / truncated 여부
- `--manifest-stdout`를 쓰면 manifest는 stdout으로, 진행 로그는 stderr로 출력합니다.

### `--sheet` / `--sheet-regex`
- 기본값은 모든 sheet를 처리합니다.
- `--sheet`는 exact match이며 반복 가능합니다.
- `--sheet-regex`는 POSIX regex입니다.
- `--fail-if-no-sheet`를 함께 쓰면 아무 시트도 선택되지 않았을 때 바로 실패합니다.

### Resource Guards
- `--max-sheets n`: 선택된 시트를 앞에서부터 `n`개까지만 처리합니다.
- `--max-rows-per-sheet n`: 각 시트의 출력 행 수를 `n`개로 제한합니다.
- `--max-output-bytes n`: 전체 출력 바이트 수 상한을 둡니다.
- `--fail-if-truncated`: 위 제한에 걸려 일부만 출력됐을 때 exit code를 `1`로 만듭니다.

### `--stdout`
- 선택된 단일 시트만 stdout으로 출력합니다.
- `--sheet` 또는 `--sheet-regex`로 결과가 정확히 1개 sheet가 되도록 지정해야 합니다.
- `--output-dir`와 함께 쓸 수 없습니다.

## Wildcard Behavior

### generic 모드
- `*` 문자를 시트/헤더 텍스트에서 그대로 유지합니다.
- 단, 출력 파일명에서는 파일 시스템에 안전한 형태로 정리합니다.

### `game-db-fast` 모드
- 기본 동작: `*`를 출력명에서 제거합니다.
- `--no-wildcard`: `*`가 포함된 시트는 스킵하고, `*`가 포함된 헤더 컬럼은 제외합니다.

## Valid Characters
아래 규칙은 `--mode game-db-fast`에서만 적용됩니다.

- 시트명 허용 문자: `A-Z`, `a-z`, `0-9`, `-`, `_`, `*`
- 헤더명 허용 문자: `A-Z`, `a-z`, `0-9`, `-`, `_`, `*`
- 공백, 한글, 기타 특수문자가 포함되면 해당 시트/컬럼은 제외될 수 있습니다.

generic 모드에서는 위 제한을 적용하지 않습니다.

## Examples

### 기본 generic export
```bash
./xlsx_to_tsv report.xlsx 1
```
- 모든 worksheet를 export합니다.
- 1행부터 그대로 TSV로 출력합니다.

### generic + formatting
```bash
./xlsx_to_tsv report.xlsx 2 --formatted
```
- 1행은 건너뛰고 2행부터 출력합니다.
- 날짜/시간/퍼센트 같은 숫자를 사람이 읽는 값으로 바꿉니다.

### generic + merged/hidden 보정
```bash
./xlsx_to_tsv report.xlsx 2 --formatted --expand-merged --skip-hidden
```
- 병합 영역을 채우고 숨김 row/column을 제외합니다.

### generic + CSV
```bash
mkdir -p out && ./xlsx_to_tsv report.xlsx 2 --formatted --csv --output-dir out
```
- 각 시트를 `.csv`로 출력합니다.
- 결과 파일은 `out/` 아래에 생성됩니다.

### generic + JSONL
```bash
./xlsx_to_tsv report.xlsx 2 --formatted --jsonl
```
- 첫 번째 출력 행을 key로 사용합니다.

### list sheets as JSON
```bash
./xlsx_to_tsv report.xlsx --list-sheets --json
```
- workbook 구조만 빠르게 확인합니다.

### select one sheet and stream to stdout
```bash
./xlsx_to_tsv report.xlsx 2 --sheet Summary --stdout
```
- `Summary` 시트만 stdout으로 출력합니다.

### export with manifest
```bash
mkdir -p out && ./xlsx_to_tsv report.xlsx 2 --sheet Summary --output-dir out --manifest-json out/manifest.json
```
- 데이터 파일과 manifest JSON을 함께 생성합니다.

### strict bounded export
```bash
mkdir -p out && ./xlsx_to_tsv report.xlsx 2 --output-dir out --max-sheets 3 --max-rows-per-sheet 1000 --max-output-bytes 10485760 --fail-if-truncated
```
- 시트 수, 시트별 행 수, 전체 출력 바이트 수를 제한합니다.

### game DB fast export
```bash
mkdir -p out && ./xlsx_to_tsv data.xlsx 4 --mode game-db-fast --output-dir out
```
- 4행을 헤더 행으로 사용합니다.
- 게임 DB 규칙에 맞지 않는 시트/컬럼은 제외합니다.
- 결과 파일은 `out/` 아래에 생성됩니다.

### game DB fast + strict wildcard
```bash
./xlsx_to_tsv data.xlsx 4 --mode game-db-fast --no-wildcard
```
- `*`가 들어간 시트/컬럼을 완전히 제외합니다.

## Output
- 기본 출력 확장자는 `.tsv`
- `--csv`는 `.csv`, `--jsonl`은 `.jsonl`
- `--output-dir`를 주지 않으면 현재 작업 디렉터리에 출력합니다.
- 파일명은 시트명 기반이며, 안전하지 않은 문자와 공백은 `_`로 정리합니다.
- 정리된 파일명이 겹치면 `__2`, `__3` 같은 suffix를 붙입니다.
- TSV/CSV는 delimiter 기반 행 출력, JSONL은 line-delimited JSON object 출력입니다.
- `--fail-on-output-collision`을 주면 suffix를 붙이지 않고 즉시 실패합니다.
