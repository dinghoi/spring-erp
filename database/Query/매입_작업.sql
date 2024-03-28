

#매입세금계산서


SELECT *
FROM general_cost
WHERE tax_bill_yn = 'Y'
	AND substring(slip_date, 1, 4) = '2021'
	AND emp_no = '101954'
;	




SELECT *
FROM card_slip
WHERE substring(slip_date, 1, 7) = '2021-07'
	AND account_end = 'Y'
	AND emp_no = '100029'
	#AND bonbu = '경영본부'
ORDER BY slip_date DESC
;

SELECT *
FROM card_slip
WHERE substring(slip_date, 1, 7) = '2021-07'
	AND account_end = 'Y'
	#AND emp_no = '100029'
	AND emp_no = '101793'
	#AND bonbu = '경영본부'
	#AND emp_name = '이재원'
ORDER BY slip_date DESC
;


/*
UPDATE card_slip SET
	emp_no = '101793',
	emp_company = '케이네트웍스',
	bonbu = '경영본부',	
	org_name = '경영본부'
WHERE substring(slip_date, 1, 7) = '2021-08'
	AND account_end = 'Y'
	AND emp_name = '이재원'
;

*/