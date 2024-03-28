

CREATE PROCEDURE cost_select(
	IN p_emp_no varchar(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 
DESC :
'
proc_body :
BEGIN
	SELECT 
		(SELECT count(*) FROM general_cost WHERE cost_reg = '0'
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
	WHERE emp_no = p_emp_no;

END;
