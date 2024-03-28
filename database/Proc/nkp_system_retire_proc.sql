/*
SHOW CREATE PROCEDURE nkp_system_retire_proc;


CALL nkp_system_retire_proc('102211', '2021-07-31', '퇴직발령', '지현주');

*/
DROP PROCEDURE IF EXISTS nkp_system_retire_proc;
CREATE PROCEDURE nkp_system_retire_proc(
	IN p_emp_no VARCHAR(6),
	IN p_emp_end_date DATE,
	IN p_app_id VARCHAR(10) CHARSET utf8,
	IN p_emp_mod_user VARCHAR(20) CHARSET utf8
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-09-13
DESC :
- 퇴직 발령 일괄 처리

RETURN VALUE :
'
proc_body :
BEGIN 
	DECLARE v_master_cnt INT;
	DECLARE v_mem_cnt INT;
	DECLARE v_sawo_cnt INT;
	DECLARE v_max_seq INT;
	
	#처리 상태 값
	DECLARE state INT DEFAULT 0;
	
	DECLARE EXIT HANDLER FOR SQLEXCEPTION
	BEGIN
		ROLLBACK;
		SET state = -1;
	END;	
	
	#트랜잭션 설정
	START TRANSACTION;
	
	#인사마스터 조회
	SELECT count(*) INTO @master_cnt
	FROM emp_master
	WHERE emp_no = p_emp_no;
	
	SET v_master_cnt = @master_cnt;	
	
	#인사마스터 퇴직 처리
	IF v_master_cnt > 0 THEN
		UPDATE emp_master SET
			emp_end_date = p_emp_end_date, 
			emp_pay_id = '2',
			emp_mod_user = p_emp_mod_user,
			emp_mod_date = NOW()
		WHERE emp_no = p_emp_no;
	END IF;
	
	#회원정보 조회
	SELECT count(*) INTO @mem_cnt
	FROM memb
	WHERE user_id = p_emp_no;
	
	SET v_mem_cnt = @mem_cnt;
	
	#회원정보 퇴직 처리
	IF v_mem_cnt > 0 THEN 
		UPDATE memb SET
			grade = '6'
		WHERE user_id = p_emp_no;
	END IF;
	
	#사우회 조회
	SELECT count(*) INTO @sawo_cnt	
	FROM emp_sawo_mem
	WHERE sawo_empno = p_emp_no;
	
	SET v_sawo_cnt = @sawo_cnt;	
	
	#사우회 퇴직 처리
	IF v_sawo_cnt > 0 THEN 
		UPDATE emp_sawo_mem SET
			sawo_out = '퇴직',
			sawo_out_date = p_emp_end_date
		WHERE sawo_empno = p_emp_no;
	END IF;
	
	#인사발령 조회
	SELECT MAX(app_seq) INTO @max_seq
	FROM emp_appoint
	WHERE app_empno = p_emp_no;	
	
	SET v_max_seq = @max_seq;
	
	#app_seq 설정
	IF IFNULL(v_max_seq, '') = '' THEN
		SET @app_seq = '001';
	ELSE 
		SET @app_seq = CONCAT('00', v_max_seq + 1);
	END IF;	
	
	#인사발령 처리	
	INSERT INTO emp_appoint(app_empno, app_seq, app_id, app_date, app_emp_name, app_id_type,
		app_to_company, app_to_orgcode, app_to_org, app_to_grade, app_to_job,
		app_to_position, app_be_enddate, app_comment, app_reg_date)
	SELECT p_emp_no, @app_seq, p_app_id, p_emp_end_date, emp_name, '개인사정',
		emp_company, emp_org_code, emp_org_name, emp_grade, emp_job,
		emp_position, NULL, '계약종료', NOW()
	FROM emp_master
	WHERE emp_no = p_emp_no;
		
	COMMIT;
	
	SELECT state;
END;
	
