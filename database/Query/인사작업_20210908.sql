
#회원 정보
SELECT *
FROM memb
#WHERE user_name = '이재준'
WHERE user_id = '100668'
;

#조직 정보
SELECT *
FROM emp_org_mst
WHERE org_code = '6519'
;

#인사 마스터
SELECT emp_company, emp_bonbu, emp_saupbu, emp_team, 
	emp_org_name, emp_reside_place, emp_reside_company
FROM emp_master
WHERE emp_no = '102643'
;

#인사 마스터(월)
SELECT emp_company, emp_bonbu, emp_saupbu, emp_team, 
	emp_org_name, emp_reside_place, emp_reside_company,
	emp_org_code, emp_org_name
FROM emp_master_month
#WHERE emp_month = '202101'
WHERE substring(emp_month, 1, 4) = '2021'
	AND emp_no = '102643'
;

/*
UPDATE emp_master_month SET
	emp_company = '케이시스템',
	emp_bonbu = 'DI사업부문',
	emp_saupbu = '',
	emp_team = '',
	emp_org_name = 'DI사업부문'
WHERE emp_month = '202101'
	AND emp_no = '102645'
;
*/

#급여
SELECT *
FROM pay_month_give
WHERE substring(pmg_yymm, 1, 4) = '2021'
#WHERE pmg_yymm = '202101'
	AND pmg_emp_no = '102644'
;


#인사발령
SELECT *
FROM emp_appoint
WHERE substring(app_date, 1, 4) = '2021'
AND app_empno = '102489'
;


#일반경비
SELECT *
FROM general_cost
WHERE cost_reg = '0'
	AND (tax_bill_yn <> 'Y' or isnull(tax_bill_yn))
	AND (slip_gubun = '비용')  
	AND substring(slip_date, 1, 4) = '2021'
	AND emp_no = '102489'
ORDER BY slip_date DESC 	 
;

#야특근
SELECT *
FROM overtime
WHERE substring(work_date, 1, 4) = '2021'
	AND mg_ce_id = '102489'
;
	

#교통비
SELECT *
FROM transit_cost
WHERE substring(run_date, 1, 4) = '2021'
AND mg_ce_id = '102489'
;


#매입세금계산서
SELECT *
FROM general_cost
WHERE tax_bill_yn = 'Y'
	AND substring(slip_date, 1, 4) = '2021'
	AND emp_no = '102489'
;	

#카드전표
SELECT *
FROM card_slip
#WHERE substring(slip_date, 1, 4) = '2021'
WHERE substring(slip_date, 1, 7) = '2021-08'
	AND account_end = 'Y'
	AND emp_no = '102489'
ORDER BY slip_date DESC 	
LIMIT 100
;


/*
UPDATE card_slip SET
	#emp_company = '',
	bonbu = 'DI사업부문',
	#saupbu = '',
	#team = '',
	org_name = 'DI사업부문'
	#reside_place = '',
	#reside_company = ''
WHERE substring(slip_date, 1, 7) = '2021-08'
	AND account_end = 'Y'
	AND emp_no = '102489';
*/	

=====================================




SET @insa_date = '202101';

/*
SELECT pmgt.pmg_company
FROM pay_month_give AS pmgt
INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no
	AND emmt.emp_month = @insa_date
INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code
INNER JOIN emp_master AS emtt ON emmt.emp_no = emtt.emp_no
WHERE pmgt.pmg_id = '1'
	AND pmgt.pmg_yymm = @insa_date	
	#AND pmgt.pmg_company IN ('케이원', '케이원정보통신')
GROUP BY pmgt.pmg_company
;
*/

/*
# 케이원

DI사업부문
공공SI본부
금융SI본부

NI본부

ICT본부
SI2본부
OA수행본부
SI수행본부
SI1본부
SM1 사업본부

- SM2 사업본부
- SM수행본부
- ITO 사업본부

null
*/

SELECT pmgt.pmg_emp_no, pmgt.pmg_emp_name, pmgt.pmg_grade,
	pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team,
	pmgt.pmg_org_code, pmgt.pmg_org_name,
	
	emtt.emp_org_code, emtt.emp_org_name,
	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team,	
	emtt.emp_reside_company, emtt.emp_reside_place,
	emtt.cost_center, emtt.cost_group, emtt.emp_first_date, emtt.emp_end_date
FROM pay_month_give AS pmgt
INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no
	AND emmt.emp_month = @insa_date
INNER JOIN emp_master AS emtt ON emmt.emp_no = emtt.emp_no
WHERE pmgt.pmg_id = '1'
	AND pmgt.pmg_yymm = @insa_date
	#AND pmgt.pmg_company = '에스유에이치'
	#AND pmgt.pmg_company IN ('코리아디엔씨', '케이시스템')
	#AND pmgt.pmg_company = '케이네트웍스'
	AND pmgt.pmg_company IN ('케이원', '케이원정보통신')
	#AND pmgt.pmg_bonbu IN ('ITO 사업본부')
	AND pmgt.pmg_bonbu = '' OR pmgt.pmg_bonbu IS NULL 
ORDER BY emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, emtt.emp_no
;

======================================


#인사 마스터(월) - 월별 급여 일자, 월별 인사발령 여부
SELECT #emmt.emp_no, emmt.emp_name, emmt.emp_month,
	emmt.emp_grade, emmt.emp_pay_id, emmt.emp_first_date, emmt.emp_end_date, 
	
	(SELECT count(*) FROM emp_appoint WHERE app_empno = emmt.emp_no
	AND substring(replace(app_date, '-', ''), 1, 6) = emmt.emp_month
	AND app_id = '이동발령') AS 'appYN',
	
	pmgt.pmg_date, pmgt.pmg_company,
	
	emtt.emp_company, 
	
	emmt.emp_company, emmt.emp_bonbu, emmt.emp_saupbu, emmt.emp_team, 
	emmt.emp_reside_place, emmt.emp_reside_company,
	emmt.emp_org_code, emmt.emp_org_name,
	
	pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team,
	pmgt.pmg_org_name	
FROM emp_master_month AS emmt
LEFT OUTER JOIN pay_month_give AS pmgt ON emmt.emp_no = pmgt.pmg_emp_no
	AND emmt.emp_month = pmgt.pmg_yymm
INNER JOIN emp_master AS emtt ON emmt.emp_no = emtt.emp_no	
WHERE substring(emmt.emp_month, 1, 4) = '2021'
	AND emmt.emp_no = '102636'
;

/*

UPDATE emp_master_month SET
	#emp_company = '케이시스템',
	#emp_bonbu = '금융SI본부',
	emp_saupbu = '',
	#emp_team = '인사총무'
	#emp_org_code = '6520',
	#emp_org_name = '금융SI본부'	
WHERE substring(emp_month, 1, 4) = '2021'
#WHERE emp_month = '202108'
	AND emp_no = '100703'
;
*/

#비용 건수
SELECT 
	(SELECT count(*)	FROM general_cost WHERE cost_reg = '0'
	AND (tax_bill_yn <> 'Y' or isnull(tax_bill_yn))
	AND slip_gubun = '비용'  
	AND emp_no = emtt.emp_no
	AND substring(slip_date, 1, 4) = '2021') AS '일반',
		
	(SELECT count(*) FROM overtime WHERE mg_ce_id = emtt.emp_no
	AND substring(work_date, 1, 4) = '2021') AS '야특근',
		
	(SELECT count(*) FROM transit_cost WHERE mg_ce_id = emtt.emp_no
	AND substring(run_date, 1, 4) = '2021') AS '교통비',
	
	(SELECT count(*) FROM general_cost WHERE tax_bill_yn = 'Y'
	AND emp_no = emtt.emp_no
	AND substring(slip_date, 1, 4) = '2021') AS '매입',
	
	(SELECT count(*) FROM card_slip WHERE account_end = 'Y'
	AND emp_no = emtt.emp_no
	AND substring(slip_date, 1, 4) = '2021') AS '카드'
FROM emp_master AS emtt
WHERE emp_no = '102617'
;


#==============================================


SET @insa_date = '202101';

SELECT pmgt.pmg_emp_no, pmgt.pmg_emp_name, pmgt.pmg_grade,
	pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team,
	pmgt.pmg_org_code, pmgt.pmg_org_name,
	
	emtt.emp_org_code, emtt.emp_org_name,
	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team,	
	emtt.emp_reside_company, emtt.emp_reside_place,
	emtt.cost_center, emtt.cost_group, emtt.emp_first_date, emtt.emp_end_date
FROM pay_month_give AS pmgt
INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no
	AND emmt.emp_month = @insa_date
INNER JOIN emp_master AS emtt ON emmt.emp_no = emtt.emp_no
WHERE pmgt.pmg_id = '1'
	AND pmgt.pmg_yymm = @insa_date
	AND pmgt.pmg_emp_no = '101382'	

;


#====================================================


#일반경비


SELECT slip_date, slip_seq, slip_gubun, 
	emp_company, bonbu, saupbu, team, org_name, reside_place, 
	emp_name, emp_no, emp_grade, cost_center, 
	reg_id, reg_user
FROM general_cost
WHERE cost_reg = '0'
	AND (tax_bill_yn <> 'Y' or isnull(tax_bill_yn))
	AND (slip_gubun = '비용')  
	AND substring(slip_date, 1, 4) = '2021'
	AND emp_no = '101887'
ORDER BY slip_date DESC 	 
;




#야특근


SELECT concat("'", ovtt.work_date, "'") AS work_date, 
	ovtt.mg_ce_id, ovtt.user_name, ovtt.user_grade, ovtt.acpt_no, 
	ovtt.emp_company, ovtt.bonbu, ovtt.saupbu, ovtt.team, ovtt.org_name, ovtt.reside_place,
	ovtt.reg_id, ovtt.reg_user,
	
	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team,
	emtt.emp_org_name, emtt.emp_reside_place
FROM overtime AS ovtt
INNER JOIN emp_master AS emtt ON ovtt.mg_ce_id = emtt.emp_no
WHERE substring(work_date, 1, 4) = '2021'
	AND mg_ce_id = '102433'
;


#교통비


SELECT concat("'", trct.run_date, "'") AS run_date, 
	trct.mg_ce_id, 
	concat("'", trct.run_seq, "'") AS run_seq, 
	trct.user_name,
	trct.emp_company, trct.bonbu, trct.saupbu, trct.team, trct.org_name, trct.reside_place,
	trct.reg_id, trct.reg_user, trct.mod_id, trct.mod_user,
	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team,
	emtt.emp_org_name, emtt.emp_reside_place
FROM transit_cost AS trct
INNER JOIN emp_master AS emtt ON trct.mg_ce_id = emtt.emp_no
WHERE trct.mg_ce_id = '101912'
	AND substring(trct.run_date, 1, 4) = '2021'
#	AND substring(run_date, 1, 7) = '2021-09'
;

SELECT *
FROM transit_cost
WHERE mg_ce_id = '101912'
	AND substring(run_date, 1, 4) = '2021'
;




#매입세금계산서


SELECT *
FROM general_cost
WHERE tax_bill_yn = 'Y'
	AND substring(slip_date, 1, 4) = '2021'
	AND emp_no = '102489'
;	

#카드전표


SELECT cslt.approve_no, cslt.cancel_yn, cslt.slip_date, cslt.card_type, 
	cslt.emp_no, cslt.emp_name, cslt.emp_company, cslt.bonbu, cslt.saupbu, cslt.team, 
	cslt.org_name, cslt.reside_place, cslt.reside_company, 
	cslt.mod_id, cslt.mod_name,
	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team,
	emtt.emp_org_name, emtt.emp_reside_place, emtt.emp_reside_company
FROM card_slip AS cslt
INNER JOIN emp_master AS emtt ON cslt.emp_no = emtt.emp_no
WHERE substring(cslt.slip_date, 1, 4) = '2021'

	AND cslt.account_end = 'Y'
	AND cslt.emp_no = '101912'
ORDER BY cslt.slip_date DESC
;

#A/S

SELECT count(*)
FROM as_acpt
WHERE mg_ce_id = '101912'
	AND substring(acpt_date, 1, 4) = '2021'
;


SELECT asat.acpt_no, asat.acpt_date, asat.acpt_man, asat.acpt_grade,
	asat.mg_ce_id, asat.mg_ce, asat.as_process, asat.team, asat.saupbu,	
	asat.reg_id, 
	
	emtt.emp_name, 
	(SELECT emp_team FROM emp_master WHERE emp_no = asat.mg_ce_id) AS as_team,
	emtt.emp_saupbu
FROM as_acpt AS asat
INNER JOIN emp_master AS emtt ON asat.reg_id = emtt.emp_no
WHERE asat.mg_ce_id = '101912'
	AND substring(asat.acpt_date, 1, 4) = '2021'
ORDER BY asat.acpt_date ASC
;







