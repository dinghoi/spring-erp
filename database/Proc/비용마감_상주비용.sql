/*
#비용 마감 체크
*/
DROP PROCEDURE IF EXISTS USP_COMPANY_END_COST_CNT;
CREATE PROCEDURE USP_COMPANY_END_COST_CNT(
	IN p_from_date varchar(10),
	IN p_to_date varchar(10),
	IN p_end_month varchar(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 상주 비용 마감 > 비용 마감 체크
'
proc_body :
BEGIN
	SET @from_date = p_from_date;
	SET @to_date = p_to_date;
	SET @end_month = p_end_month;
	SET @reside_sw = 'Y';

	#세금계산서 비용 미등록 처리 내역 확인
	SELECT COUNT(*) INTO @taxBillTotCnt
	FROM tax_bill
	WHERE bill_id = '1' AND cost_reg_yn = 'N'
		AND (bill_date >= @from_date AND bill_date <= @to_date);

	IF @taxBillTotCnt > 0 THEN
		SET @reside_sw = 'N';
	END IF;

	#비용 마감 총 개수 조회
	SELECT COUNT(*) INTO @nonSideTotCnt
	FROM cost_end
	WHERE end_month = @end_month
		AND saupbu <> '상주비용';

	IF @nonSideTotCnt > 0 THEN
		SELECT COUNT(*) INTO @costEndTotCnt
		FROM cost_end
		WHERE end_month = @end_month
			AND (end_yn = 'N' OR end_yn = 'C')
			AND saupbu <> '상주비용'
			AND saupbu <> '공통비/직접비배분';

		IF @costEndTotCnt > 0 THEN
			SET @reside_sw = 'N';
		END IF;
	END IF;

	SELECT @reside_sw;
END;

/*
#인사 정보 조회(손익제외)
*/
DROP PROCEDURE IF EXISTS USP_COMPANY_END_ORG_SEL;
CREATE PROCEDURE USP_COMPANY_END_ORG_SEL(
	IN p_end_month varchar(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
COMMENT 'AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 상주 비용 마감 > 인사 정보 조회(손익제외)'
proc_body :
BEGIN
	SELECT emmt.emp_no, eomt.org_bonbu, eomt.org_code
	FROM emp_master_month AS emmt
	INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code
		AND eomt.org_month = p_end_month
	WHERE emmt.emp_month = p_end_month
		AND emmt.cost_center <> '손익제외';
END;

/*
#직원 별 관리사업부 지정
*/
DROP PROCEDURE IF EXISTS USP_COMPANY_END_ORG_UP;
CREATE PROCEDURE USP_COMPANY_END_ORG_UP(
	IN p_org_bonbu varchar(30) CHARSET utf8,
	IN p_cost_year varchar(4),
	IN p_end_month varchar(6),
	IN p_emp_no varchar(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 상주 비용 마감 > 직원 별 관리사업부 지정
'
proc_body :
BEGIN
	SET @saupbu = p_org_bonbu;
	SET @sales_year = p_cost_year;
	SET @end_month = p_end_month;
	SET @emp_no = p_emp_no;

	SELECT sort_seq INTO @sort_seq
	FROM sales_org
	WHERE saupbu = @saupbu
		AND sales_year = @sales_year;

	IF @sort_seq = '' OR @sort_seq = NULL THEN
		SET @saupbu = '';
	END IF;

	#월 직원 현황 수정
	UPDATE emp_master_month SET
		mg_saupbu = @saupbu
	WHERE emp_month = @end_month
		AND emp_no = @emp_no;

	#급여 직원 정보 수정
	UPDATE pay_month_give SET
		mg_saupbu = @saupbu
	WHERE pmg_yymm = @end_month
		AND pmg_emp_no = @emp_no;
END;

/*
#관리사업부 미지정된 상주 정보 조회
*/
DROP PROCEDURE IF EXISTS USP_COMPANY_END_RESIDE_SEL;
CREATE PROCEDURE USP_COMPANY_END_RESIDE_SEL(
	IN p_end_month varchar(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
COMMENT 'AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 상주 비용 마감 > 관리사업부 미지정된 상주 정보 조회
'
proc_body :
BEGIN
	SELECT emp_reside_company
	FROM emp_master_month
	WHERE emp_month = p_end_month
		AND mg_saupbu = ''
		AND emp_reside_company <> ''
		AND cost_center <> '손익제외'
		AND emp_pay_id <> '2';
END;

/*
#관리사업부 미지정 직원 정보 업데이트
*/
DROP PROCEDURE IF EXISTS USP_COMPANY_END_RESIDE_UP;
CREATE PROCEDURE USP_COMPANY_END_RESIDE_UP(
	IN p_end_month varchar(6),
	IN p_org_code varchar(4),
	IN p_emp_no varchar(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
COMMENT 'AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 상주 비용 마감 > 관리사업부 미지정 직원 정보 업데이트
'
proc_body :
BEGIN
	SET @end_month = p_end_month;
	SET @org_code = p_org_code;
	SET @emp_no = p_emp_no;

	SELECT org_bonbu INTO @org_bonbu
	FROM emp_org_mst_month
	WHERE org_month = @end_month
		AND org_code = @org_code;

	#직원 월 정보 관리사업부 업데이트
	UPDATE emp_master_month SET
		mg_saupbu =  @org_bonbu
	WHERE emp_month = ''
		AND mg_saupbu = ''
		AND emp_no = @emp_no;

	#급여 월 정보 관리사업부 업데이트
	UPDATE pay_month_give SET
		mg_saupbu =  @org_bonbu
	WHERE pmg_yymm = @end_month
		AND mg_saupbu = ''
		AND pmg_emp_no = @emp_no;
END;

/*
#아르바이트 마감 초기화 및 조회
*/
DROP PROCEDURE IF EXISTS USP_COMPANY_END_ALBA_INIT;
CREATE PROCEDURE USP_COMPANY_END_ALBA_INIT(
	IN p_end_month varchar(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
COMMENT 'AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 상주 비용 마감 > 아르바이트 마감 초기화 및 조회
'
proc_body :
BEGIN
	SET @end_month = p_end_month;

	#초기화
	UPDATE pay_alba_cost SET
		mg_saupbu = '',
		cost_center = ''
	WHERE rever_yymm = @end_month;

	UPDATE pay_alba_cost SET
		cost_center = '상주직접비'
	WHERE rever_yymm = @end_month
		AND cost_company NOT IN ('공통', '전사', '부문', '기타', '본사', '케이원정보통신', '케이원', '');

	SELECT company, org_name
	FROM pay_alba_cost
	WHERE rever_yymm = @end_month
		AND (cost_company = '공통' OR cost_company <> '전사' OR cost_company <> '부문'
			OR cost_company = '기타' 	OR cost_company = '본사' OR cost_company = '케이원정보통신'
			OR cost_company = '케이원' OR cost_company = '')
	GROUP BY company, org_name;
END;


