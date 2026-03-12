# xlsx2tsv
xlsx 파일을 TSV/CSV/JSONL로 고속 변환합니다.

## Usage
```bash
./xlsx_to_tsv <input.xlsx> [start_row] [--mode generic|game-db-fast] [--output-dir dir] [--no-wildcard] [--formatted] [--expand-merged] [--skip-hidden] [--csv] [--jsonl]
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
- `--all-sheets`: generic 모드의 legacy alias
- `--no-wildcard`: `game-db-fast` 전용 옵션
- `--formatted`: generic 전용, `styles.xml` 숫자 포맷 적용
- `--expand-merged`: generic 전용, merged cell을 top-left 값으로 확장
- `--skip-hidden`: generic 전용, hidden row/column 제외
- `--csv`: generic 전용, `.csv` 출력
- `--jsonl`: generic 전용, `.jsonl` 출력
  - 첫 번째 출력 행을 key로 사용하고 이후 행을 object로 출력

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
./xlsx_to_tsv report.xlsx 2 --formatted --csv --output-dir out
```
- 각 시트를 `.csv`로 출력합니다.
- 결과 파일은 `out/` 아래에 생성됩니다.

### generic + JSONL
```bash
./xlsx_to_tsv report.xlsx 2 --formatted --jsonl
```
- 첫 번째 출력 행을 key로 사용합니다.

### game DB fast export
```bash
./xlsx_to_tsv data.xlsx 4 --mode game-db-fast --output-dir out
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
