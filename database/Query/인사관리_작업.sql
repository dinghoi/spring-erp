#인사 마스터(월)
SELECT emp_month, emp_name, emmt.emp_pay_id,    
	emp_org_code, emp_org_name, 
	emp_company, emp_bonbu, emp_saupbu, emp_team,
	emp_reside_company, emp_reside_place,
	cost_center
FROM emp_master_month AS emmt
INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code
WHERE substring(emp_month, 1, 4) = '2021'
	AND emp_no = '100244'
;




/*

SELECT *
FROM emp_master_month
WHERE emp_month = '202012'
AND emp_no = '100294'
;


#급여
SELECT *
FROM pay_month_give
WHERE substring(pmg_yymm, 1, 4) = '2021'
#WHERE pmg_yymm = '202101'
	AND pmg_emp_no = '101278'
;

SELECT *
FROM emp_master
WHERE emp_org_code = '6675'
;

SELECT *
FROM emp_org_mst
WHERE org_code = '6554'
;


#인사발령
SELECT *
FROM emp_appoint
WHERE substring(app_date, 1, 4) = '2021'
AND app_empno = '101773'
;


SELECT *
FROM emp_master
WHERE emp_no = '102433'
;



SELECT *
FROM memb
WHERE user_id = '102433'
;


*/




