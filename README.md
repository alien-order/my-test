# Oracle 테이블 현황 조사 스크립트

## 개요

Oracle DB 테이블 현황을 자동으로 조사하여 기존 Excel 폼 파일에 결과를 기록하는 Python 스크립트입니다.  
두 가지 버전으로 구성되며, 각 번호 셀 클릭 시 해당 테이블의 샘플 데이터 시트로 이동하는 하이퍼링크가 자동 생성됩니다.

---

## 파일 구성

| 파일 | 설명 |
|---|---|
| `survey_v1.py` | PROD 단일 스키마 조사 |
| `survey_v2.py` | PROD + ARCHIVE + INFA 3개 스키마 조사 |
| `README.md` | 이 문서 |

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
