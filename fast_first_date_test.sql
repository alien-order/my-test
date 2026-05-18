-- =============================================================================
-- fast_first_date_test.sql
--   "데이터보관최초일자" 를 빠르게 구하는 3가지 방법 비교 테스트
--   SQL*Plus / SQL Developer 에서 한 테이블에 대해 직접 돌려보세요.
--
--   사용 전, 아래 두 값을 실제 테이블/날짜컬럼으로 치환:
--     &TBL  = 테이블명 (대문자, 따옴표 없이)   예) ORDER_HIST
--     &COL  = 최초일자 판단에 쓰는 날짜컬럼     예) REG_DT
--   (SQL*Plus 는 실행 시 값을 물어봅니다. SQL Developer 는 바인드 입력창)
-- =============================================================================
SET SERVEROUTPUT ON
SET TIMING ON
SET LINESIZE 200
SET PAGESIZE 100


-- -----------------------------------------------------------------------------
-- [방법 1] 옵티마이저 통계의 low_value  ← 테이블 스캔 0, 즉시 응답 (가장 빠름)
--   * user_tab_col_statistics.low_value 는 마지막 통계수집 시점의 컬럼 최솟값
--   * DATE/TIMESTAMP 는 RAW 인코딩 → DBMS_STATS.CONVERT_RAW_VALUE 로 디코딩
-- -----------------------------------------------------------------------------

-- 1-a) 통계 신선도 먼저 확인 (last_analyzed 가 오래됐으면 부정확할 수 있음)
SELECT column_name,
       num_distinct,
       num_nulls,
       last_analyzed,
       ROUND(SYSDATE - last_analyzed) AS stats_age_days
FROM   user_tab_col_statistics
WHERE  table_name  = '&TBL'
  AND  column_name = '&COL';

-- 1-b) low_value / high_value 디코딩 (DATE 컬럼)
DECLARE
  r_low   user_tab_col_statistics.low_value%TYPE;
  r_high  user_tab_col_statistics.high_value%TYPE;
  d_low   DATE;
  d_high  DATE;
BEGIN
  SELECT low_value, high_value
  INTO   r_low, r_high
  FROM   user_tab_col_statistics
  WHERE  table_name = '&TBL' AND column_name = '&COL';

  DBMS_STATS.CONVERT_RAW_VALUE(r_low,  d_low);
  DBMS_STATS.CONVERT_RAW_VALUE(r_high, d_high);

  DBMS_OUTPUT.PUT_LINE('[방법1] 통계 MIN = ' || TO_CHAR(d_low,  'YYYY-MM-DD'));
  DBMS_OUTPUT.PUT_LINE('[방법1] 통계 MAX = ' || TO_CHAR(d_high, 'YYYY-MM-DD'));
EXCEPTION
  WHEN NO_DATA_FOUND THEN
    DBMS_OUTPUT.PUT_LINE('[방법1] 통계 없음 — 미수집 또는 컬럼명 확인 필요');
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('[방법1] 디코딩 실패(' || SQLERRM ||
                         ') — VARCHAR/TIMESTAMP 컬럼이면 방법3 사용');
END;
/

-- 1-c) VARCHAR 날짜 컬럼('YYYYMMDD' 등)일 때 — 문자열 최솟값 디코딩
DECLARE
  r_low user_tab_col_statistics.low_value%TYPE;
  v_low VARCHAR2(100);
BEGIN
  SELECT low_value INTO r_low
  FROM   user_tab_col_statistics
  WHERE  table_name = '&TBL' AND column_name = '&COL';

  DBMS_STATS.CONVERT_RAW_VALUE(r_low, v_low);
  DBMS_OUTPUT.PUT_LINE('[방법1-VARCHAR] 통계 MIN 문자열 = ' || v_low);
EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('[방법1-VARCHAR] 해당 없음/실패: ' || SQLERRM);
END;
/


-- -----------------------------------------------------------------------------
-- [방법 2] 파티션 경계값  ← 날짜로 RANGE 파티션된 테이블이면 스캔 없이 획득
--   * 파티션 키가 날짜 컬럼인지 확인 후, 가장 오래된 파티션의 경계값 사용
--   * 첫 파티션 high_value = 가장 오래된 데이터의 상한 (실제 최솟값은 그 이하)
-- -----------------------------------------------------------------------------

-- 2-a) 이 테이블의 파티션 키 컬럼
SELECT name AS table_name, column_name, column_position
FROM   user_part_key_columns
WHERE  name = '&TBL'
ORDER  BY column_position;

-- 2-b) 가장 오래된 파티션 2개의 경계값(high_value)
--      partition_position 1 의 high_value 가 최소 날짜대를 알려줌
SELECT partition_position,
       partition_name,
       high_value          -- LONG: TO_DATE(' 2019-01-01 00:00:00', ...) 형태
FROM   user_tab_partitions
WHERE  table_name = '&TBL'
ORDER  BY partition_position
FETCH FIRST 2 ROWS ONLY;

-- 2-c) (참고) 가장 오래된 파티션만 직접 MIN — 전체 스캔보다 훨씬 빠름
--      &PART 에 위 2-b 의 첫 partition_name 을 넣어 실행
-- SELECT MIN("&COL") FROM "&TBL" PARTITION ("&PART");


-- -----------------------------------------------------------------------------
-- [방법 3] 직접 MIN()  ← 정확하지만 인덱스 없으면 풀스캔 (3억 row 면 느림)
--   * 실행계획을 먼저 보고 INDEX (MIN/MAX) 인지 TABLE FULL SCAN 인지 확인
-- -----------------------------------------------------------------------------

-- 3-a) 실행계획 확인 (실제 실행 전)
EXPLAIN PLAN FOR SELECT MIN("&COL") FROM "&TBL";
SELECT * FROM TABLE(DBMS_XPLAN.DISPLAY(NULL, NULL, 'BASIC'));
--  → 'INDEX FULL SCAN (MIN/MAX)' 면 빠름 / 'TABLE ACCESS FULL' 이면 느림

-- 3-b) 실제 실행 + 소요시간 (SET TIMING ON 으로 시간 출력)
SELECT MIN("&COL") AS min_value FROM "&TBL";

-- 3-c) 날짜컬럼에 인덱스가 있는지 (있으면 3-a 가 MIN/MAX 스캔으로 빨라짐)
SELECT i.index_name, ic.column_position, ic.column_name
FROM   user_indexes      i
JOIN   user_ind_columns  ic ON ic.index_name = i.index_name
WHERE  ic.table_name  = '&TBL'
  AND  ic.column_name = '&COL'
ORDER  BY i.index_name, ic.column_position;


SET TIMING OFF
-- =============================================================================
-- 해석 가이드
--   · 방법1 값 ≈ 방법3 값 이고 stats_age_days 가 작다 → 방법1 채택 (압도적으로 빠름)
--   · 방법1 값 > 방법3 값 (통계수집 후 더 과거 데이터 유입) → 통계 재수집 또는 방법3
--   · 방법2 가 가능(날짜 RANGE 파티션) → 첫 파티션만 MIN (방법2-c) 도 충분히 빠르고 정확
--   · 방법3 실행계획이 INDEX (MIN/MAX) 면 사실 방법3 도 충분히 빠름 → 별도 최적화 불필요
-- =============================================================================
