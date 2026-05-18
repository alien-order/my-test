# Oracle 테이블 현황 조사 스크립트

## 개요

Oracle DB 테이블 현황을 자동으로 조사하여 기존 Excel 폼 파일에 결과를 기록하는 Python 스크립트입니다.  
두 가지 버전으로 구성되며, 각 번호 셀 클릭 시 해당 테이블의 샘플 데이터 시트로 이동하는 하이퍼링크가 자동 생성됩니다.

---

## 파일 구성

| 파일 | 설명 |
|---|---|
| `survey_v1.py` | PROD 단일 스키마 조사 |
| `survey_v1.ipynb` | survey_v1 Jupyter 노트북 (단계별 실행·체크포인트) |
| `survey_v2.py` | PROD + ARCHIVE + INFA 3개 스키마 조사 |
| `survey_v2.ipynb` | survey_v2 Jupyter 노트북 (단계별 실행·체크포인트) |
| `relink_by_tablename.ipynb` | 완성본 후처리: 시트명·하이퍼링크를 번호→테이블명 전환 + 최초일자→테이블 생성일 교체 |
| `fast_first_date.ipynb` | 완성본 후처리: `데이터보관최초일자` 를 통계 기반으로 빠르게 재계산 |
| `fast_first_date_test.sql` | 최초일자 3가지 산출 방법 비교용 수동 테스트 쿼리 |
| `README.md` | 이 문서 |

> `survey_*` 로 완성한 Excel 을 입력으로 받아 추가 가공하는 **후처리 노트북** 2종이 있습니다.
> 둘 다 **원본을 수정하지 않고** 새 파일을 만들며, 헤더의 `소재` 컬럼 유무로 v1/v2 를 자동 판별합니다.

---

## 노트북 버전 공통 기능

각 스크립트와 대응하는 Jupyter 노트북으로, 셀 단위 실행과 체크포인트로 긴 작업을 단계별로 나눌 수 있습니다.

| 기능 | 설명 |
|---|---|
| 단계별 실행 | DB 접속 / 사용여부 조회 / 테이블 수집 / Excel 출력 셀 분리 |
| 체크포인트 | 수집 완료 후 `.pkl` 저장 → 재실행 시 DB 재조회 없이 Excel만 재출력 |
| 제어문자 제거 | `IllegalCharacterError` 방지 — Oracle 데이터의 제어문자 자동 제거 |
| 샘플 시트 | `pandas DataFrame.to_excel()` 사용 (openpyxl 직접 기록 대신) |

> **v2 노트북**: PROD / ARCHIVE / INFA 수집 셀이 분리(셀 7·8·9)되어 DB별로 개별 실행·재실행 가능.  
> 체크포인트에 원시 데이터(`prod_data`, `arch_data`, `infa_data`)도 함께 저장되므로 로드 후 합산 로직만 재실행 가능.

---

## 사전 조건

```bash
pip install oracledb pandas openpyxl python-dateutil
```

- Oracle Instant Client 설치 필요 (`ORACLE_CLIENT_LIB` 경로 설정)
- `v$sql_plan`, `v$sqlarea` 접근: **SELECT ANY DICTIONARY** 권한
- `dba_hist_active_sess_history` 접근: **DBA 권한** + **Diagnostics Pack 라이선스**

---

## 공통 출력 컬럼

| 컬럼 | 설명 |
|---|---|
| 번호 | 순번 (클릭 시 샘플 시트로 이동) |
| 시스템명 | 고정값 (설정에서 지정) |
| 테이블명 | Oracle 테이블명 |
| 테이블설명 | `user_tab_comments` |
| 사용여부 | 아래 판단 로직 참고 |
| 데이터보관최초일자 | 아래 판단 로직 참고 |
| 1개월평균데이터생성건수 | 최근 3개월 INSERT 수 ÷ 3 (또는 전체 ÷ 경과 개월) |
| 칼럼수(개) | `user_tab_columns` COUNT |
| 전체테이블용량 | 테이블 + LOB 세그먼트 합산 (단위: GB, 최소 0.1GB) |

---

## 사용여부 판단 로직 (우선순위 순)

| 순위 | 조건 | 판단값 |
|---|---|---|
| 1 | ASH에 SELECT 샘플 있음 | `Y (ASH-SELECT확인)` |
| 2 | ASH에 DML 샘플 있음 | `Y (ASH-DML확인)` |
| 3 | v$sqlarea Library Cache에 SELECT 흔적 | `Y (Q3-SELECT흔적)` |
| 4 | user_tab_modifications에 INSERT/UPDATE/DELETE 이력 | `Y (DML이력)` |
| 5 | 건수 = 0 (통계 기준) | `N` |
| 6 | 통계 미수집 (num_rows IS NULL) | `확인필요 (통계미수집)` |
| 7 | 건수 > 0 + 테이블명 마스터 패턴 | `확인필요 (마스터추정)` |
| 8 | 건수 > 0 + 패턴 없음 | `확인필요` |

> **마스터 패턴**: CODE, MST, MASTER, CMM, COMMON, TYPE, KIND, STAT, STATUS, GRP, GROUP, CLASS, CONFIG, PARAM, META, MENU, ROLE, AUTH 등

---

## 데이터보관최초일자 판단 로직

```
건수 = 0                  → "-"
건수 > 0 + 날짜 컬럼 있음  → MIN(날짜컬럼)
건수 > 0 + 날짜 컬럼 없음  → 테이블 생성일 (user_objects.created)
```

**날짜 컬럼 탐지 우선순위**
1. DATE/TIMESTAMP 타입 + 생성일 패턴 컬럼명 (CRE, REG, INS, FIRST, OPEN 포함)
2. DATE/TIMESTAMP 타입 (패턴 없어도)
3. VARCHAR이지만 샘플 데이터가 날짜 형태 (YYYYMMDD, YYYYMMDDHHMMSS 등)

---

## 샘플 데이터 시트

- 시트명: 번호 숫자 (예: `1`, `2`, `3`)
- 최초일자 판단에 사용된 날짜 컬럼 기준 **최근 10건** (`ORDER BY 날짜컬럼 DESC`)
- 날짜 컬럼 없는 경우: 임의 10건 (정렬 보장 없음, 시트 상단에 안내 표시)
- 모든 컬럼 출력

---

## Excel 설정

```python
# 기존 폼 파일 경로
TEMPLATE_EXCEL = r"C:\path\to\template.xlsx"

# 번호가 시작할 셀 (이 셀부터 오른쪽 방향으로 컬럼 채움)
# 예) B4 → 번호=B4, 시스템명=C4, 테이블명=D4, ...
START_CELL = "B4"

# 시스템명 고정값
SYSTEM_NAME = "TEST"
```

---

## Version 1: `survey_v1.py` — PROD 단일 스키마

### 대상
- PROD 스키마 테이블만 조사

### 입력 설정
```python
DB_CONFIGS = {
    "prod": {
        "host": "PROD_HOST",
        "port": 1521,
        "service_name": "PROD_SERVICE",
        "user": "PROD_USER",
        "password": "PROD_PASSWORD",
    }
}
```

### 출력
- 공통 컬럼만 출력 (소재 컬럼 없음)
- 지정한 기존 Excel 파일에 덮어쓰기 후 샘플 시트 추가

---

## Version 2: `survey_v2.py` — PROD + ARCHIVE + INFA

### 대상
- PROD, ARCHIVE, INFA 3개 스키마 테이블 전체

### 입력 설정
```python
DB_CONFIGS = {
    "prod": {
        "host": "PROD_HOST", "port": 1521,
        "service_name": "PROD_SERVICE",
        "user": "PROD_USER", "password": "PROD_PASSWORD",
    },
    "archive": {
        "host": "ARCHIVE_HOST", "port": 1521,
        "service_name": "ARCHIVE_SERVICE",
        "user": "ARCHIVE_USER", "password": "ARCHIVE_PASSWORD",
    },
    "infa": {
        "host": "INFA_HOST", "port": 1521,
        "service_name": "INFA_SERVICE",
        "user": "INFA_USER", "password": "INFA_PASSWORD",
    },
}
```

### 추가 컬럼: 소재

| 소재값 | 의미 |
|---|---|
| `Combined` | PROD + ARCHIVE 양쪽에 모두 존재 |
| `Prod Only` | PROD에만 존재 |
| `Archive Only` | ARCHIVE에만 존재 |
| `INFA` | INFA 스키마 테이블 |

### 합산 계산 대상 컬럼

| 컬럼 | 합산 방식 |
|---|---|
| 데이터보관최초일자 | PROD와 ARCHIVE 중 더 이른 날짜 |
| 1개월평균데이터생성건수 | PROD + ARCHIVE 합산 |
| 전체테이블용량 | PROD + ARCHIVE 세그먼트 합산 |

> `Archive Only` 테이블의 사용여부는 PROD에 없으므로 기본 `N` 처리

---

## 실행 방법

```bash
# Version 1
python survey_v1.py

# Version 2
python survey_v2.py
```

실행 시 프롬프트:
```
기존 Excel 파일 경로: C:\reports\template.xlsx
번호 시작 셀 (기본 B4): B4
```

---

## 주의사항

- `DBMS_STATS.FLUSH_DATABASE_MONITORING_INFO` 실행 후 조회해야 DML 이력이 최신화됨
- ASH 조회 기간은 AWR 보관 설정에 따라 다름 (기본 8일)
  - 확인: `SELECT retention FROM dba_hist_wr_control;`
- v$sqlarea는 현재 메모리 기준이므로 Q3 미노출 = 미사용 확정이 아님
- 기존 Excel 파일은 덮어쓰기 전 자동 백업 생성 (`_backup_HHMMSS.xlsx`)

---

# 후처리 노트북

`survey_*` 로 완성한 Excel 을 입력으로 받아 추가 가공합니다.
공통 동작: **원본 미수정 → 새 파일 생성**, 헤더의 `소재` 유무로 **v1/v2 자동 판별**(셀 2에서 강제 지정 가능), 셀 단위 실행·체크포인트 지원.
설정은 각 노트북 **셀 2** 만 수정하면 됩니다 (DB 접속정보는 survey 실행 때와 동일하게, v1 은 PROD 만).

---

## `relink_by_tablename.ipynb` — 시트명·링크 전환 + 최초일자→생성일

출력: `원본명_테이블명버전.xlsx`

| 변경 | 내용 |
|---|---|
| 시트명 | 번호(`1`,`2`…) → **테이블명** (Excel 금지문자 `\ / ? * [ ] :` 치환, 31자 제한, 중복 시 `_2`) |
| 하이퍼링크 | `번호` 셀에서 제거 → **`테이블명` 셀**에 연결 |
| 데이터보관최초일자 | 전부 **테이블 생성일**(`user_objects.created`)로 교체 (소재에 맞는 DB 선택, 없으면 폴백) |

> 생성일은 현재 스키마 기준이라 테이블이 재생성됐다면 실제 적재 시점과 다를 수 있음 (의도된 동작).

---

## `fast_first_date.ipynb` — 최초일자 빠른 재계산

3억 row 테이블의 `SELECT MIN(날짜)` 풀스캔이 느린 문제 해결. `데이터보관최초일자` 컬럼만 다시 계산.
출력: `원본명_최초일자갱신.xlsx`

**계산 우선순위 (테이블당)**

| 순위 | 조건 | 산출 |
|---|---|---|
| 1 | `num_rows = 0` | `-` |
| 2 | 날짜컬럼 없음 | 테이블 생성일 |
| 3 | DATE 날짜컬럼 | `user_tab_col_statistics.low_value` 디코딩 (스캔 0, 즉시) |
| 4 | VARCHAR 날짜컬럼 | 통계 low_value 문자열 디코딩 → 날짜 정규화 |
| 5 | 통계없음·TIMESTAMP·디코딩 실패 | `MIN()` 폴백 (`MIN_FALLBACK=False` 면 생성일 대체, 완전 무스캔) |

- 통계 기반이라 마지막 통계수집 이후 더 과거 데이터가 유입됐다면 부정확할 수 있음
- 셀 9 리포트: 방식별 건수 + `MIN 폴백`(통계수집 권장) / `STALE`(통계 `STATS_WARN_AGE_DAYS`일 초과) / `테이블없음` 집계
- TIMESTAMP 컬럼은 `DBMS_STATS.CONVERT_RAW_VALUE` 미지원 → 통계 경로 건너뛰고 폴백

---

## `fast_first_date_test.sql` — 최초일자 산출 3방법 비교 (수동)

SQL*Plus / SQL Developer 에서 한 테이블에 `&TBL`/`&COL` 만 넣고 직접 비교.

| 방법 | 내용 | 특징 |
|---|---|---|
| 1 | 통계 `low_value` 디코딩 (+ 신선도 `last_analyzed`) | 스캔 0, 즉시. 통계 기준이라 근사 |
| 2 | 날짜 RANGE 파티션 경계값 / 첫 파티션만 `MIN` | 파티션 구조 의존, 스캔 최소 |
| 3 | 직접 `MIN()` + `EXPLAIN PLAN` + 인덱스 존재 확인 | 정확. 인덱스 없으면 풀스캔(느림) |

> 노트북 적용 전, SQL 로 한두 테이블 검증해 방법1 값 신뢰도를 확인하는 흐름 권장.
> 파일 하단의 해석 가이드에 채택 기준 정리되어 있음.
