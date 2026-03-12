# xlsx2tsv
xlsx 파일을 tsv 파일로 고속으로 변환해 줌

## Usage
```bash
./xlsx_to_tsv <input.xlsx> [start_row] [--no-wildcard] [--all-sheets] [--formatted] [--expand-merged] [--skip-hidden] [--csv] [--jsonl]
```

### Parameters
- `input.xlsx`: 변환할 XLSX 파일 경로 (필수)
- `start_row`: 변환을 시작할 행 번호 (1부터 시작, 기본값: 1)
  - `start_row` 이전 행은 모두 무시됨
  - 기본 모드에서는 `start_row` 행이 출력 TSV의 첫 행이 되며, 컬럼 유효성 검사와 와일드카드 처리의 기준 헤더 행으로 사용됨
- `--no-wildcard`: 와일드카드(*) 문자 필터링 모드 활성화
- `--all-sheets`: 모든 worksheet를 export하는 범용 모드
  - 시트명/헤더명 유효성 검사를 하지 않음
  - 헤더 기준 컬럼 제거와 `*` 제거를 하지 않음
  - 출력 파일명만 안전한 형태로 정리함
- `--formatted`: `styles.xml`의 숫자 포맷을 적용하는 LLM 친화 출력 모드
  - `--all-sheets`를 자동으로 활성화함
  - 날짜/시간/날짜시간을 ISO 형태에 가깝게 정규화함
  - 퍼센트, 천 단위 구분, zero-padding, duration 형식을 반영함
- `--expand-merged`: merge 영역을 top-left 값으로 채워서 export함
  - `--all-sheets`를 자동으로 활성화함
  - 병합된 셀 아래/옆의 빈 셀도 사람이 보는 형태에 가깝게 채워짐
- `--skip-hidden`: 숨김 row/column을 export에서 제외함
  - `--all-sheets`를 자동으로 활성화함
- `--csv`: generic output을 `.csv`로 직접 기록함
  - `--all-sheets`를 자동으로 활성화함
- `--jsonl`: generic output을 `.jsonl`로 직접 기록함
  - `--all-sheets`를 자동으로 활성화함
  - 첫 번째로 export되는 행을 JSON key로 사용하고, 그 다음 행부터 JSON object를 출력함

## Wildcard (*) Character Behavior

### 기본 모드 (Default)
`*` 문자는 **출력에서 자동으로 제거**됩니다.

- **시트(Sheet)**: `*` 문자가 제거된 이름으로 파일 생성
  - 예: `*Sales_2024` 시트 → `Sales_2024.tsv` 파일 생성
  - 예: `Data*2024` 시트 → `Data2024.tsv` 파일 생성

- **컬럼(Column)**: 헤더 행의 `*` 문자가 제거된 이름으로 TSV에 출력
  - 예: 헤더 행이 `*ID`, `Name`, `*Price`, `Amount`인 경우
  - TSV 헤더 출력: `ID`, `Name`, `Price`, `Amount`
  - 모든 데이터 행도 해당 컬럼에 정상 출력됨

### --no-wildcard 모드
`*` 문자가 포함된 시트/컬럼을 **완전히 배제**합니다.

- **시트(Sheet)**: `*` 문자가 **어디든** 포함된 시트는 전체 스킵
  - 예: `*Sales`, `Data*`, `S*ales` → 모두 처리되지 않음

- **컬럼(Column)**: 헤더 행에서 `*` 문자가 **어디든** 포함된 컬럼은 출력에서 제외
  - 예: 헤더 행이 `*ID`, `Name`, `*Price`, `Amount`인 경우
  - TSV 출력: `Name`, `Amount` 컬럼만 포함
  - `*ID`, `*Price` 컬럼과 해당 컬럼의 모든 데이터 제외

### --all-sheets 모드
`*` 문자를 포함한 시트/컬럼도 그대로 export됩니다.

- 시트 스킵 없음
- 컬럼 스킵 없음
- 헤더의 `*`도 그대로 유지됨
- 단, 출력 파일명에서만 파일 시스템에 문제가 되는 문자를 정리함

### --formatted 모드
업무용 XLSX를 LLM에 넘기기 쉽게 숫자 값을 사람이 읽는 형태로 바꿉니다.

- 날짜 예: `45293` → `2024-01-02`
- 시간 예: `0.5` → `12:00`
- 날짜시간 예: `45293.6278...` → `2024-01-02 15:04:05`
- 퍼센트 예: `0.125` → `12.50%`
- zero-padding 예: `42` + `00000` 형식 → `00042`
- boolean 예: `1`, `0` → `TRUE`, `FALSE`

### --expand-merged 모드
병합된 셀을 raw XML 그대로 비워 두지 않고 top-left 값으로 확장합니다.

- 예: `A3:C4`가 병합되어 있고 `A3 = Merged`면 `A3`, `B3`, `C3`, `A4`, `B4`, `C4`가 모두 `Merged`로 export됨
- row XML이 없는 병합 하단 행도 필요한 경우 synthetic row로 보정되어 출력될 수 있음

### --skip-hidden 모드
숨김 처리된 row/column을 generic export에서 제외합니다.

- `<row hidden="1">`인 행은 출력하지 않음
- `<col hidden="1">` 범위의 열은 출력하지 않음
- 숨김 열을 건너뛰기 때문에 출력 열 인덱스는 원본 Excel 열 문자와 1:1로 맞지 않을 수 있음

### CSV / JSONL 출력
generic 모드에서 TSV 대신 CSV 또는 JSONL로 바로 쓸 수 있습니다.

- `--csv`: 각 시트를 `.csv`로 출력함
- `--jsonl`: 각 시트를 `.jsonl`로 출력함
- JSONL은 첫 번째 출력 행을 key로 사용하므로, 보통 header 행을 `start_row`로 맞추는 편이 적절함

## Valid Characters
시트명과 헤더 행의 컬럼명에 허용되는 문자:
- 영문 대소문자: `A-Z`, `a-z`
- 숫자: `0-9`
- 특수문자: `-` (하이픈), `_` (언더스코어), `*` (애스터리스크, 기본 모드)

**제외되는 경우:**
- 공백, `@`, `#`, `$`, 한글, 기타 특수문자가 포함된 시트명은 시트 전체가 스킵됨
- 공백, `@`, `#`, `$`, 한글, 기타 특수문자가 포함된 헤더 컬럼은 해당 컬럼 전체가 출력에서 제외됨

`--all-sheets`에서는 위 제한을 적용하지 않습니다.

## Examples

### 기본 모드 (Default)
```bash
./xlsx_to_tsv data.xlsx
```
- `*Internal_Data` 시트 → `Internal_Data.tsv` 생성
- `*ID`, `Name`, `Amount` 컬럼 → `ID`, `Name`, `Amount`로 TSV에 출력 (`*` 제거)

### 4행부터 변환
```bash
./xlsx_to_tsv data.xlsx 4
```
- 1~3행은 무시됨
- 4행이 출력 TSV의 헤더 행으로 사용됨
- 5행부터 실제 데이터가 이어짐

### --no-wildcard 모드
```bash
./xlsx_to_tsv data.xlsx 4 --no-wildcard
```
- `*Internal_Data` 시트 → 전체 시트 스킵됨 (처리되지 않음)
- 4행 헤더가 `*ID`, `Name`, `Amount`이면 → `Name`, `Amount`만 출력 (`*ID` 컬럼 전체 제외)

### 모든 시트 export
```bash
./xlsx_to_tsv data.xlsx 4 --all-sheets
```
- 유효성 검사 때문에 스킵되던 시트도 함께 export됨
- 4행부터 그대로 TSV로 기록됨
- `*start_date` 같은 헤더도 그대로 유지됨

### 포맷 적용 export
```bash
./xlsx_to_tsv report.xlsx 1 --formatted
```
- 모든 worksheet를 export함
- 날짜/시간/퍼센트 같은 숫자 셀을 사람이 읽는 값으로 변환함
- LLM 입력용 TSV를 만들 때 raw serial number보다 읽기 쉬운 값을 얻을 수 있음

### merged cell 보정 + hidden 제외
```bash
./xlsx_to_tsv report.xlsx 1 --formatted --expand-merged --skip-hidden
```
- 병합 영역을 top-left 값으로 채움
- 숨김 row/column을 제외함
- 사람이 보는 표 구조에 더 가까운 TSV를 얻을 수 있음

### CSV 직접 출력
```bash
./xlsx_to_tsv report.xlsx 1 --formatted --expand-merged --skip-hidden --csv
```
- `.tsv` 대신 `.csv` 파일을 생성함

### JSONL 직접 출력
```bash
./xlsx_to_tsv report.xlsx 1 --formatted --expand-merged --skip-hidden --jsonl
```
- 첫 번째 출력 행을 key로 사용함
- 그 다음 행부터 sheet별 JSON object를 한 줄씩 출력함

## Output
- 기본값은 각 시트마다 별도의 TSV 파일 생성
- 기본 모드에서는 유효한 시트명만 처리됨
- `--all-sheets`에서는 workbook의 모든 worksheet를 처리함
- `--formatted`에서는 `styles.xml` 숫자 포맷을 반영한 정규화 문자열을 출력함
- `--expand-merged`에서는 병합 영역을 top-left 값으로 채움
- `--skip-hidden`에서는 숨김 row/column을 제외함
- `--csv`는 `<SheetName>.csv`, `--jsonl`은 `<SheetName>.jsonl`로 출력함
- 파일명: `<SheetName>.<ext>` (`*`는 제거되고, 파일 시스템에서 문제가 되는 문자와 공백은 `_`로 변환)
- 정리된 파일명이 겹치면 자동으로 `__2`, `__3` 같은 suffix를 붙여 충돌을 피함
- TSV/CSV는 delimiter 기반 행 출력, JSONL은 line-delimited JSON object 출력
