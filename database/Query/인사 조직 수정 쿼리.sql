

SELECT emp_company
FROM emp_master
GROUP BY emp_company
;

/*
에스유에이치
케이더봄

케이시스템
코리아디엔씨

케이네트웍스

케이원
케이원정보통신
*/

SELECT COUNT(*)
FROM emp_master
WHERE emp_pay_id <> '2'
	/*AND (emp_company = '케이원' OR emp_company = '케이원정보통신')*/
	/*AND emp_company = '케이네트웍스'*/
;




SELECT emtt.emp_bonbu
FROM emp_master AS emtt
INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
LEFT OUTER JOIN emp_master_month AS emmt ON emtt.emp_no = emmt.emp_no
	AND emmt.emp_month = '202104'
LEFT OUTER JOIN pay_month_give AS pmgt ON emmt.emp_no = pmgt.pmg_emp_no
	AND pmgt.pmg_yymm = '202104'
WHERE emtt.emp_pay_id <> '2'
	AND emtt.emp_company = '케이원정보통신'
GROUP BY emtt.emp_bonbu	
;

/*

*/

SELECT emtt.emp_saupbu
FROM emp_master AS emtt
INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
LEFT OUTER JOIN emp_master_month AS emmt ON emtt.emp_no = emmt.emp_no
	AND emmt.emp_month = '202104'
LEFT OUTER JOIN pay_month_give AS pmgt ON emmt.emp_no = pmgt.pmg_emp_no
	AND pmgt.pmg_yymm = '202104'
WHERE emtt.emp_pay_id <> '2'
	AND emtt.emp_company = '케이원정보통신'
	
	/*AND (emtt.emp_bonbu = '' OR emtt.emp_bonbu IS NULL) */
	AND emtt.emp_bonbu = 'SM수행본부'
GROUP BY emtt.emp_saupbu	
;

/*

*/

SELECT pmgt.pmg_org_code, pmgt.cost_center, pmgt.cost_group, pmgt.mg_saupbu,
	emmt.emp_org_code, emmt.cost_center, emmt.cost_group, emmt.mg_saupbu, emmt.cost_except,
	eomt.org_code, eomt.org_level, eomt.org_company, eomt.org_bonbu, 
	eomt.org_saupbu, eomt.org_team, eomt.org_reside_company, eomt.org_reside_place,
	eomt.org_cost_center, eomt.org_cost_group,
	emtt.emp_org_code, emtt.emp_no, emtt.emp_name, emtt.emp_type, emtt.emp_grade,
	emtt.emp_in_date, emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team,
	emtt.emp_reside_company, emtt.emp_reside_place, emtt.emp_stay_name, 
	emtt.cost_center, emtt.cost_group
FROM emp_master AS emtt
INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
LEFT OUTER JOIN emp_master_month AS emmt ON emtt.emp_no = emmt.emp_no
	AND emmt.emp_month = '202104'
LEFT OUTER JOIN pay_month_give AS pmgt ON emmt.emp_no = pmgt.pmg_emp_no
	AND pmgt.pmg_yymm = '202104'
WHERE emtt.emp_pay_id <> '2'
	AND emtt.emp_company = '케이원정보통신'
	
	/*AND (emtt.emp_bonbu = '' OR emtt.emp_bonbu IS NULL)*/
	AND emtt.emp_bonbu = 'SM수행본부'
	
	/*AND (emtt.emp_saupbu = '' OR emtt.emp_saupbu IS NULL)*/
	AND emtt.emp_saupbu = 'SM1사업부'
;
