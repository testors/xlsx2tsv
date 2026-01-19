# xlsx2tsv
xlsx 파일을 tsv 파일로 고속으로 변환해 줌

## Usage
```bash
./xlsx_to_tsv <input.xlsx> [start_row] [--no-wildcard]
```

### Parameters
- `input.xlsx`: 변환할 XLSX 파일 경로 (필수)
- `start_row`: 변환을 시작할 행 번호 (1부터 시작, 기본값: 1)
- `--no-wildcard`: 와일드카드(*) 문자 필터링 모드 활성화

## Wildcard (*) Character Behavior

### 기본 모드 (Default)
`*` 문자는 **출력에서 자동으로 제거**됩니다.

- **시트(Sheet)**: `*` 문자가 제거된 이름으로 파일 생성
  - 예: `*Sales_2024` 시트 → `Sales_2024.tsv` 파일 생성
  - 예: `Data*2024` 시트 → `Data2024.tsv` 파일 생성

- **컬럼(Column)**: `*` 문자가 제거된 이름으로 TSV에 출력
  - 예: 원본 헤더가 `*ID`, `Name`, `*Price`, `Amount`인 경우
  - TSV 헤더 출력: `ID`, `Name`, `Price`, `Amount`
  - 모든 데이터 행도 해당 컬럼에 정상 출력됨

### --no-wildcard 모드
`*` 문자가 포함된 시트/컬럼을 **완전히 배제**합니다.

- **시트(Sheet)**: `*` 문자가 **어디든** 포함된 시트는 전체 스킵
  - 예: `*Sales`, `Data*`, `S*ales` → 모두 처리되지 않음

- **컬럼(Column)**: `*` 문자가 **어디든** 포함된 컬럼은 출력에서 제외
  - 예: 원본 헤더가 `*ID`, `Name`, `*Price`, `Amount`인 경우
  - TSV 출력: `Name`, `Amount` 컬럼만 포함
  - `*ID`, `*Price` 컬럼과 해당 컬럼의 모든 데이터 제외

## Valid Characters
시트명과 컬럼명에 허용되는 문자:
- 영문 대소문자: `A-Z`, `a-z`
- 숫자: `0-9`
- 특수문자: `-` (하이픈), `_` (언더스코어), `*` (애스터리스크, 기본 모드)

**제외되는 경우:**
- 공백, `@`, `#`, `$`, 한글, 기타 특수문자가 포함된 시트/컬럼은 자동으로 스킵됨

## Examples

### 기본 모드 (Default)
```bash
./xlsx_to_tsv data.xlsx
```
- `*Internal_Data` 시트 → `Internal_Data.tsv` 생성
- `*ID`, `Name`, `Amount` 컬럼 → `ID`, `Name`, `Amount`로 TSV에 출력 (`*` 제거)

### 2행부터 변환
```bash
./xlsx_to_tsv data.xlsx 2
```
- 1행(헤더)을 건너뛰고 2행부터 변환 시작

### --no-wildcard 모드
```bash
./xlsx_to_tsv data.xlsx 1 --no-wildcard
```
- `*Internal_Data` 시트 → 전체 시트 스킵됨 (처리되지 않음)
- `*ID`, `Name`, `Amount` 컬럼 → `Name`, `Amount`만 출력 (`*ID` 컬럼 전체 제외)

## Output
- 각 시트마다 별도의 TSV 파일 생성
- 파일명: `<SheetName>.tsv` (안전하지 않은 문자는 `_`로 변환, `*`는 제거)
- 구분자: 탭(Tab) 문자
