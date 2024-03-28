

SELECT eomt.org_code, eomt.org_level, eomt.org_name, 
	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team,
	eomt.org_reside_company, eomt.org_reside_place, 
	eomt.org_cost_center, eomt.org_cost_group,
	pmgt.pmg_org_code, pmgt.pmg_org_name, 
	pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team,
	pmgt.pmg_reside_company, pmgt.pmg_reside_place,
	pmgt.cost_center, pmgt.cost_group, pmgt.mg_saupbu,	
	emmt.emp_org_code, emmt.emp_org_name,
	emmt.emp_company, emmt.emp_bonbu, emmt.emp_saupbu, emmt.emp_team,
	emmt.emp_reside_company, emmt.emp_reside_place,
	emmt.cost_center, emmt.cost_group, emmt.mg_saupbu, emmt.cost_except,
	emtt.emp_org_code, emtt.emp_no, emtt.emp_name, emtt.emp_type, emtt.emp_grade, 
	emtt.emp_in_date,
	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, emtt.emp_org_name,
	emtt.emp_reside_company, emtt.emp_reside_place, emtt.emp_stay_name, emtt.cost_center
FROM emp_master AS emtt
INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
LEFT OUTER JOIN pay_month_give AS pmgt ON emtt.emp_no = pmgt.pmg_emp_no
	AND pmgt.pmg_yymm = '202105'
LEFT OUTER JOIN emp_master_month AS emmt ON emtt.emp_no = emmt.emp_no
	AND emmt.emp_month = '202105'
WHERE emtt.emp_no = '101109'
;


-- 조직
SELECT eomt.org_code, eomt.org_level, eomt.org_name, 
	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team,
	eomt.org_reside_company, eomt.org_reside_place, 
	eomt.org_cost_center, eomt.org_cost_group
FROM emp_org_mst AS eomt
WHERE org_code = '6519'
;

-- 급여(월)
SELECT pmgt.pmg_org_code, pmgt.pmg_org_name, 
	pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team,
	pmgt.pmg_reside_company, pmgt.pmg_reside_place,
	pmgt.cost_center, pmgt.cost_group, pmgt.mg_saupbu
FROM pay_month_give AS pmgt
WHERE pmgt.pmg_yymm = '202105'
	AND pmgt.pmg_emp_no = '102694'
;

-- 인사(월)
SELECT emp_org_code, emp_org_name,
	emp_company, emp_bonbu, emp_saupbu, emp_team,
	emp_reside_company, emp_reside_place,
	cost_center, cost_group, mg_saupbu, cost_except
FROM emp_master_month
WHERE emp_month = '202105'
	AND emp_no = '101476'
;

-- 인사
SELECT emp_org_code, emp_org_name,
	emp_name,
	emp_company, emp_bonbu, emp_saupbu, emp_team,
	emp_reside_company, emp_reside_place,
	cost_center, cost_group, mg_saupbu
FROM emp_master
WHERE emp_no = '101210'
;

--4056

SELECT *
FROM emp_org_mst
WHERE org_name = 'N/W 영업팀'
;

SELECT *
FROM emp_org_mst
WHERE org_code = '6319'
;




SELECT emp_org_name
FROM emp_master
WHERE emp_no IN
(
'101109',
'101476',
'102555',
'102580',
'100285',
'101089',
'101210',
'101368',
'101585',
'101601',
'101737',
'101743',
'101746',
'101934',
'102313',
'100486',
'100489',
'100490',
'100491',
'101114',
'101570',
'101573',
'101578',
'101983',
'102428',
'102675',
'102676',
'102718',
'102722',
'102444',
'100298',
'100326',
'100431',
'100521',
'100604',
'102571',
'102274',
'102591',
'102612',
'102621',
'102622',
'102641',
'100032',
'100107',
'100125',
'100173',
'100249',
'100294',
'100335',
'100354',
'101161',
'101523',
'101692',
'102048',
'102188',
'102269',
'102310',
'102386',
'102421',
'102423',
'102449',
'102625',
'101778'
)
ORDER BY emp_no
;


