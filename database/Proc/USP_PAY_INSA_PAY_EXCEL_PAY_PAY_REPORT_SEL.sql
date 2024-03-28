/*
--등록 프로시저(전체)
SHOW PROCEDURE STATUS;

--해당 프로시저 확인
SHOW CREATE PROCEDURE USP_PAY_INSA_PAY_EXCEL_PAY_PAY_REPORT_SEL;

--실행
CALL USP_PAY_INSA_PAY_EXCEL_PAY_PAY_REPORT_SEL('202107', '케이시스템', '');

--삭제
DROP PROCEDURE USP_PAY_INSA_PAY_EXCEL_PAY_PAY_REPORT_SEL;
*/

CREATE PROCEDURE USP_PAY_INSA_PAY_EXCEL_PAY_PAY_REPORT_SEL(
	IN p_pmg_yymm VARCHAR(6),
	IN p_emp_company VARCHAR(6) CHARSET utf8,
	IN p_pmg_emp_name VARCHAR(20) CHARSET utf8
)
LANGUAGE SQL
NOT DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT "
AUTHOR : 허정호
DATE : 2021-08-05
DESC : - 급여 엑셀 다운로드 리스트 조회
"
BEGIN	
	SELECT pmgt.pmg_emp_no, pmgt.pmg_company, pmgt.pmg_give_total, pmgt.pmg_base_pay, pmgt.pmg_meals_pay,
		pmgt.pmg_postage_pay, pmgt.pmg_re_pay, pmgt.pmg_overtime_pay, pmgt.pmg_car_pay, pmgt.pmg_position_pay,
		pmgt.pmg_custom_pay, pmgt.pmg_job_pay, pmgt.pmg_job_support, pmgt.pmg_jisa_pay, pmgt.pmg_long_pay,
		pmgt.pmg_disabled_pay, pmgt.pmg_give_total, pmgt.pmg_emp_name, pmgt.pmg_in_date, pmgt.pmg_grade,
		pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team, pmgt.pmg_org_name, pmgt.pmg_reside_place,
		pmgt.pmg_reside_company, pmgt.cost_group, pmgt.cost_center, eomt.org_name, eomt.org_company,
		eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_reside_place, eomt.org_reside_company,
		emmt.cost_group AS costGroup, emmt.cost_center AS costCenter, 
		pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt, pmdt.de_income_tax, 
		pmdt.de_wetax, pmdt.de_year_incom_tax, pmdt.de_year_wetax, pmdt.de_year_incom_tax2, 
		pmdt.de_year_wetax2, pmdt.de_other_amt1, pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, pmdt.de_school_amt, 
		pmdt.de_nhis_bla_amt, pmdt.de_long_bla_amt, pmdt.de_deduct_total 
	FROM pay_month_give AS pmgt 
	INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no 
		AND emmt.emp_month = p_pmg_yymm
	INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code 
	INNER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no 
		AND pmdt.de_id = '1'
		AND pmdt.de_yymm = p_pmg_yymm
	WHERE pmgt.pmg_id = '1'
		AND pmgt.pmg_yymm = p_pmg_yymm
		AND eomt.org_company = p_emp_company
		AND pmgt.pmg_emp_name LIKE CONCAT('%', p_pmg_emp_name, '%')
	ORDER BY pmgt.pmg_company, pmgt.pmg_org_code, pmgt.pmg_emp_no ASC;
END;


