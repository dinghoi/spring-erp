
/*

SHOW FULL COLUMNS FROM emp_master_month;

SHOW FULL COLUMNS FROM pay_month_give;

call nkp_sys_mem_month_update('202108');
*/

DROP PROCEDURE IF EXISTS nkp_sys_mem_month_update;
CREATE DEFINER=`root`@`localhost` PROCEDURE `nkp_sys_mem_month_update`(
	IN p_curr_month varchar(6)
)
DETERMINISTIC
COMMENT '\nAUTHOR : 허정호\nDATE : 2021-09-13\nDESC :\n- 직원월별업데이트\n\nRETURN VALUE :\n'
proc_body :
BEGIN 
	DECLARE done INT DEFAULT FALSE;	
	DECLARE v_cnt INT DEFAULT -1;

	DECLARE v_emp_no varchar(6);
	DECLARE v_cost_group varchar(30) CHARSET utf8;
	DECLARE v_cost_center varchar(20) CHARSET utf8;
	DECLARE v_emp_company varchar(30) CHARSET utf8;
	DECLARE v_emp_bonbu varchar(30) CHARSET utf8;
	DECLARE v_emp_saupbu varchar(30) CHARSET utf8;
	DECLARE v_emp_team varchar(30) CHARSET utf8;
	DECLARE v_emp_org_code varchar(4) CHARSET utf8;
	DECLARE v_emp_org_name varchar(30) CHARSET utf8;
	DECLARE v_pay_cost_group varchar(50) CHARSET utf8;
	DECLARE v_pay_cost_center varchar(20) CHARSET utf8;
	DECLARE v_pmg_company varchar(20) CHARSET utf8;
	DECLARE v_pmg_bonbu varchar(30) CHARSET utf8;
	DECLARE v_pmg_saupbu varchar(30) CHARSET utf8;
	DECLARE v_pmg_team varchar(30) CHARSET utf8;
	DECLARE v_pmg_org_name varchar(30) CHARSET utf8;
	DECLARE v_mg_saupbu varchar(30) CHARSET utf8;
	
	#트랜잭션 설정
	#START TRANSACTION;
	
	#select 조회 결과를 CURSOR로 정의
	DECLARE mem_cursor CURSOR FOR
		SELECT emmt.emp_no, emmt.cost_group, emmt.cost_center, 
			emmt.emp_company, emmt.emp_bonbu, emmt.emp_saupbu, emmt.emp_team,
			emmt.emp_org_code, emmt.emp_org_name,
			pmgt.cost_group, pmgt.cost_center, 
			pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team,
			pmgt.pmg_org_name, pmgt.mg_saupbu
		FROM pay_month_give AS pmgt 
		INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no 
		AND emmt.emp_month = (p_curr_month - 1)
		INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code 
		WHERE pmgt.pmg_id = '1' 
		AND pmgt.pmg_yymm = (p_curr_month - 1);			
	
	DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE;
	
	OPEN mem_cursor;
	
	month_loop: LOOP	
		#loop 하며 mem_cursor의 데이터를 불러와 변수에 넣는다.
		FETCH mem_cursor
		INTO v_emp_no, v_cost_group, v_cost_center,
			v_emp_company, v_emp_bonbu, v_emp_saupbu, v_emp_team,
			v_emp_org_code, v_emp_org_name, 
			v_pay_cost_group, v_pay_cost_center,
			v_pmg_company, v_pmg_bonbu, v_pmg_saupbu, v_pmg_team,
			v_pmg_org_name, v_mg_saupbu;
			
		SET v_cnt = v_cnt + 1;
		
		#mem_cursor 반복이 끝나면 loop에서 빠져나간다.
		IF done THEN
			LEAVE month_loop;
		END IF;	
		
		#월별 마스터 업데이트
		UPDATE emp_master_month SET
			cost_group = v_cost_group,
			cost_center = v_cost_center,
			emp_company = v_emp_company,
			emp_bonbu = v_emp_bonbu,
			emp_saupbu = v_emp_saupbu,
			emp_team = v_emp_team,
			emp_org_code = v_emp_org_code,
			emp_org_name = v_emp_org_name
		WHERE emp_month = p_curr_month
			AND emp_no = v_emp_no;
		
		#급여 정보 업데이트
		UPDATE pay_month_give SET
			cost_group = v_pay_cost_group,
			cost_center = v_pay_cost_center,
			pmg_company = v_emp_company,
			pmg_bonbu = v_emp_bonbu,
			pmg_saupbu = v_emp_saupbu,
			pmg_team = v_emp_team,
			pmg_org_name = v_emp_org_name,
			mg_saupbu = v_mg_saupbu
		WHERE pmg_yymm = p_curr_month
			AND pmg_emp_no = v_emp_no;
	END LOOP;
	
	SELECT v_cnt;
	
	#커서를 닫는다.
	CLOSE mem_cursor;
END

