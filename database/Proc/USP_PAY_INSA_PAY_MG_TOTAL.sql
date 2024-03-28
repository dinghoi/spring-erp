/*
SHOW PROCEDURE STATUS;

CALL USP_PAY_INSA_PAY_MG_TOTAL('202107', '케이원', '', @totalCnt);
SELECT @totalCnt;

CALL USP_PAY_INSA_PAY_MG_TOTAL('202107', '케이원', '');

DROP PROCEDURE USP_PAY_INSA_PAY_MG_TOTAL;

*/

CREATE PROCEDURE USP_PAY_INSA_PAY_MG_TOTAL(
	IN p_pmg_yymm VARCHAR(6),
	IN p_emp_company VARCHAR(6) CHARACTER SET utf8,
	IN p_pmg_emp_name VARCHAR(20)  CHARACTER SET utf8,
	
	OUT totalCnt VARCHAR(10)
)
LANGUAGE SQL
NOT DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT "
AUTHOR : 허정호
DATE : 2021-08-05
DESC : - 급여 지급 현황 리스트 개수 조회
"
BEGIN	
	SELECT SQL_CALC_FOUND_ROWS *
	FROM pay_month_give AS pmgt 
	INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no	
		AND emmt.emp_month = p_pmg_yymm
	INNER JOIN emp_master AS emtt ON emmt.emp_no = emtt.emp_no				
	INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
	INNER JOIN pay_month_deduct AS pmdt ON emmt.emp_no = pmdt.de_emp_no 
		AND pmdt.de_id = '1'
		AND pmdt.de_yymm = p_pmg_yymm
	WHERE pmgt.pmg_id = '1'
		AND pmgt.pmg_yymm = p_pmg_yymm
		AND eomt.org_company = p_emp_company
		AND pmgt.pmg_emp_name LIKE CONCAT('%', p_pmg_emp_name, '%');	

	SET totalCnt := FOUND_ROWS();
END;