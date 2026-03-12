# Agent Integration Guide

이 문서는 LLM 에이전트나 wrapper가 `xlsx_to_tsv`를 안정적으로 호출하기 위한 실행 계약을 정리합니다.

## Recommended Flow

1. Health check
```bash
./xlsx_to_tsv --version
```

2. Inspect workbook structure
```bash
./xlsx_to_tsv workbook.xlsx --list-sheets --json --quiet
```

3. Export selected sheets with manifest
```bash
./xlsx_to_tsv workbook.xlsx 2 \
  --sheet Summary \
  --sheet Details \
  --output-dir out \
  --manifest-stdout \
  --quiet
```

## Mode Selection

- 기본값은 `generic`입니다.
- 업무용 XLSX, LLM 입력용 TSV/CSV/JSONL, manifest 기반 자동화에는 `generic`을 권장합니다.
- `--mode game-db-fast`는 내부 게임 DB 규칙에 맞는 특수 경로입니다.
- wrapper가 일반 문서와 게임 DB를 구분하지 못하면 `generic`을 기본값으로 두는 편이 안전합니다.

## stdout / stderr Contract

- 데이터 출력은 stdout으로 나갑니다.
  - `--stdout`: 시트 데이터
  - `--manifest-stdout`: manifest JSON
  - `--list-sheets --json`: workbook structure JSON
- fatal error와 인자 오류는 stderr로 나갑니다.
- `--quiet`는 progress/info/success summary만 숨깁니다.
- `--quiet`를 써도 stderr의 fatal error는 유지됩니다.

실무 권장:

- 구조 조회: `--list-sheets --json --quiet`
- 실제 변환: `--manifest-stdout --quiet`
- 단일 시트 스트리밍: `--stdout --quiet`

## JSON Schemas

기계가 읽는 JSON 루트에는 항상 아래 두 필드가 있습니다.

- `schema`
- `schema_version`

현재 값:

- list-sheets JSON: `schema = "list-sheets"`, `schema_version = 1`
- manifest JSON: `schema = "manifest"`, `schema_version = 1`

`tool_version`은 바이너리 버전이며, `schema_version`과 별개입니다.

### list-sheets JSON

예:

```json
{
  "schema": "list-sheets",
  "schema_version": 1,
  "tool_version": "0.1.0",
  "sheets": [
    {
      "sheet_name": "Summary",
      "state": "visible",
      "hidden": false,
      "selected": true,
      "approx_rows": 120,
      "approx_cols": 8
    }
  ]
}
```

필드 의미:

- `sheet_name`: workbook sheet name
- `state`: workbook state string (`visible`, `hidden`, `veryHidden` 등)
- `hidden`: hidden 여부
- `selected`: 현재 `--sheet` / `--sheet-regex` 필터 기준 선택 여부
- `approx_rows`, `approx_cols`: worksheet dimension 기반 추정치, 모를 때는 `-1`

### manifest JSON

예:

```json
{
  "schema": "manifest",
  "schema_version": 1,
  "tool_version": "0.1.0",
  "input_file": "workbook.xlsx",
  "mode": "generic",
  "output_format": "tsv",
  "start_row": 2,
  "output_dir": "out",
  "selected_sheet_count": 2,
  "processed_sheet_count": 2,
  "truncated": false,
  "warnings": [],
  "sheets": [
    {
      "sheet_name": "Summary",
      "state": "visible",
      "hidden": false,
      "selected": true,
      "processed": true,
      "truncated": false,
      "output_path": "out/Summary.tsv",
      "rows_emitted": 120,
      "cols_emitted": 8,
      "approx_rows": 120,
      "approx_cols": 8,
      "warnings": []
    }
  ]
}
```

wrapper가 우선적으로 봐야 하는 필드:

- top-level: `schema`, `schema_version`, `tool_version`, `truncated`, `warnings`
- per-sheet: `sheet_name`, `processed`, `truncated`, `output_path`, `rows_emitted`, `cols_emitted`, `warnings`

## Sheet Selection

- `--sheet <name>`: exact match, 반복 가능
- `--sheet-regex <pattern>`: POSIX regex
- 둘 다 없으면 모든 시트를 대상으로 처리합니다.
- 필터 결과가 비면 `--fail-if-no-sheet` 없이도 최종적으로 처리된 시트가 없어 exit code `1`이 됩니다.
- wrapper에서는 보통 `--fail-if-no-sheet`를 같이 쓰는 편이 안전합니다.

권장:

```bash
./xlsx_to_tsv workbook.xlsx --list-sheets --json --quiet
./xlsx_to_tsv workbook.xlsx 2 --sheet Summary --fail-if-no-sheet --manifest-stdout --quiet
```

## Resource Guards

대형 파일을 다룰 때는 아래 옵션을 권장합니다.

- `--max-sheets n`
- `--max-rows-per-sheet n`
- `--max-output-bytes n`
- `--fail-if-truncated`

주의:

- `--max-output-bytes`는 현재 시트를 중간에서 자를 수 있습니다.
- `--max-output-bytes`에 도달하면 이후 선택된 시트는 skip될 수 있습니다.
- 이런 상황은 manifest의 top-level `truncated`, top-level `warnings`, per-sheet `warnings`에 기록됩니다.

권장 예:

```bash
./xlsx_to_tsv workbook.xlsx 2 \
  --sheet Summary \
  --sheet Details \
  --output-dir out \
  --manifest-stdout \
  --quiet \
  --max-sheets 5 \
  --max-rows-per-sheet 5000 \
  --max-output-bytes 10485760 \
  --fail-if-truncated
```

## Option Constraints

wrapper는 아래 조합을 피해야 합니다.

- `--manifest-stdout` + `--manifest-json`
- `--manifest-stdout` + `--stdout`
- `--stdout` + `--output-dir`
- `--stdout` + `--fail-on-output-collision`
- `--list-sheets` + 변환 전용 옵션 (`--manifest-*`, `--stdout`, `--formatted`, `--expand-merged`, `--skip-hidden`, `--csv`, `--jsonl`, 리소스 가드 등)
- `--no-wildcard` without `--mode game-db-fast`

단일 시트를 stdout으로 받고 싶다면 정확히 1개 시트만 선택되어야 합니다.

## Failure Handling

wrapper는 아래를 실패로 처리하는 편이 좋습니다.

- non-zero exit code
- manifest top-level `truncated = true` when `--fail-if-truncated` is required
- per-sheet `processed = false`
- per-sheet warning 존재

즉시 실패시키고 싶은 경우:

- 시트 필터가 비면 `--fail-if-no-sheet`
- 출력 파일명 충돌을 허용하지 않으려면 `--fail-on-output-collision`
- 리소스 상한에 걸리면 `--fail-if-truncated`

## Recommended Wrapper Defaults

범용 LLM 연동 기준 권장 기본값:

- `--quiet`
- `--manifest-stdout`
- `--fail-if-no-sheet`
- 필요 시 `--max-sheets`, `--max-rows-per-sheet`, `--max-output-bytes`, `--fail-if-truncated`

권장 기본 모드:

- `generic`

권장 호출 순서:

1. `--version`
2. `--list-sheets --json --quiet`
3. 시트 선택 결정
4. `--manifest-stdout --quiet`로 실제 변환
