
#일반경비	============================================


SELECT slip_date, slip_seq, slip_gubun, 
	emp_company, bonbu, saupbu, team, org_name, reside_place, 
	emp_name, cost_center, 
	reg_user, mod_user
FROM general_cost
WHERE cost_reg = '0'
	AND (tax_bill_yn <> 'Y' or isnull(tax_bill_yn))
	AND (slip_gubun = '비용')  
	AND substring(slip_date, 1, 4) = '2021'
	AND emp_no = '100283'	
ORDER BY slip_date DESC 	 
;


SELECT *
FROM general_cost
WHERE slip_date = '2021-06-24' AND slip_seq = '041'	


	
======

UPDATE general_cost SET
	#emp_company = '케이원',
	#bonbu = 'SI1본부',
	#saupbu = '충청사업부',
	#team = '대전지원팀',
	org_name = '국세청 세종시'
	#reside_place = 'KB손해보험'
WHERE cost_reg = '0'
	AND (tax_bill_yn <> 'Y' or isnull(tax_bill_yn))
	AND (slip_gubun = '비용')  
	AND substring(slip_date, 1, 4) = '2021'
	AND emp_no = '100283'
#WHERE slip_date = '2021-04-22' AND slip_seq = '018'	
;

#야특근	============================================


SELECT work_date,
	mg_ce_id, user_name, user_grade, acpt_no, 
	emp_company, bonbu, saupbu, team, org_name, reside_place,
	reg_id, reg_user
FROM overtime
WHERE substring(work_date, 1, 4) = '2021'
#WHERE work_date = '2021-08-15'
	AND mg_ce_id = '102269'
ORDER BY work_date DESC 	
;


UPDATE overtime SET
	#bonbu = 'SI1본부',
	saupbu = '영남사업부',
	team = '부산지사',
	org_name = '한화생명 부산'
	#reside_place = '한화손해보험
WHERE substring(work_date, 1, 4) = '2021'
#WHERE work_date = '2021-07-31'
	AND mg_ce_id = '102269'
;

#교통비	============================================


SELECT run_date, mg_ce_id, run_seq, user_name,	
	emp_company, bonbu, saupbu, team, org_name, reside_place,
	reg_user, mod_user	
FROM transit_cost
WHERE mg_ce_id = '102550'
	AND substring(run_date, 1, 4) = '2021'
	#AND substring(run_date, 1, 7) = '2021-08'
ORDER BY run_date DESC 
;



UPDATE transit_cost SET
	#emp_company = '케이원',
	#bonbu = 'SI2본부',
	saupbu = '영남사업부',
	team = '부산지사',
	org_name = '한화생명 부산'
	#reside_place = '한화손해보험 전북'
WHERE mg_ce_id = '102058'
	AND substring(run_date, 1, 4) = '2021'
	#AND substring(run_date, 1, 7) = '2021-08'
;



#카드전표	============================================


SELECT approve_no, cancel_yn, slip_date, card_type, 
	emp_no, emp_name, emp_company, bonbu, saupbu, team, 
	org_name, reside_place, reside_company, 
	mod_id, mod_name
FROM card_slip
WHERE substring(slip_date, 1, 4) = '2021'
	AND account_end = 'Y'
	AND emp_no = '102269'
ORDER BY slip_date DESC
;



UPDATE card_slip SET
	#bonbu = 'SI2본부', 
	#saupbu = '영남사업부',
	#team = '상주1팀',
	#org_name = '삼성생명 경인'
	reside_place = 'KT사내망'	
WHERE account_end = 'Y'
	AND emp_no = '102269'
	#AND substring(slip_date, 1, 4) = '2021'
	AND substring(slip_date, 1, 7) = '2021-09'
;

 
