/*
SHOW PROCEDURE STATUS;

CALL USP_PAY_INSA_PAY_MG_LIST('202107', '케이시스템', '', '2', '10');

DROP PROCEDURE USP_PAY_INSA_PAY_MG_LIST;
*/

CREATE PROCEDURE USP_PAY_INSA_PAY_MG_LIST(
	IN p_pmg_yymm VARCHAR(6),
	IN p_emp_company VARCHAR(6) CHARSET utf8,
	IN p_pmg_emp_name VARCHAR(20) CHARSET utf8,
	IN p_stpage INT(6),
	IN p_pgsize INT(6)
)
LANGUAGE SQL
NOT DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT "
AUTHOR : 허정호
DATE : 2021-08-05
DESC : - 급여 지급 현황 리스트 조회
"
BEGIN	
	SELECT pmgt.pmg_emp_no, pmgt.pmg_give_total, pmgt.pmg_emp_name, 
		pmgt.pmg_grade, pmgt.pmg_position, pmgt.pmg_base_pay,
		pmgt.pmg_give_total, pmgt.pmg_org_name, pmgt.pmg_company, 
		pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team, 
		pmgt.pmg_yymm, pmgt.pmg_date, emmt.emp_first_date, emmt.emp_in_date, 
		eomt.org_code, eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team,
		pmdt.de_deduct_total
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
		AND pmgt.pmg_emp_name LIKE CONCAT('%', p_pmg_emp_name, '%')
	ORDER BY pmgt.pmg_company, pmgt.pmg_org_code, pmgt.pmg_emp_no ASC	
	LIMIT p_stpage, p_pgsize;
END;


