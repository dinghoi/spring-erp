
### 프로시저 참고 쿼리 ###
/*

#프로시저 네이밍 규칙
CREATE PROCEDURE USP_폴더명_페이지명_실행명

# 프로시저 확인
SHOW PROCEDURE STATUS;

SHOW CREATE PROCEDURE 프로시저명;

# 프로시저 수정
삭제 후 신규로 생성

# 실행 예제
CALL USP_PAY_INSA_PAY_MG_CNT('202107', '케이원', '', @totalCnt);
SELECT @totalCnt;

DROP PROCEDURE USP_PAY_INSA_PAY_MG_CNT;
*/

CREATE PROCEDURE USP_PAY_INSA_PAY_MG_CNT(
	IN p_pmg_yymm VARCHAR(6),
	IN p_emp_company VARCHAR(6) CHARACTER SET utf8,
	IN p_pmg_emp_name VARCHAR(20)  CHARACTER SET utf8,

	OUT totalCnt VARCHAR(10)
)
LANGUAGE SQL
NOT DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT "
AUTHOR : 허정호
DATE : 2021-08-05
DESC :
"
BEGIN
	#실행 명령문 값 설정
	SET @t_yymm1 = p_pmg_yymm;
	SET @t_yymm2 = p_pmg_yymm;
	SET @t_yymm3 = p_pmg_yymm;
	SET @t_company = p_emp_company;
	SET @t_name = p_pmg_emp_name;

	SET @v_sql = CONCAT("SELECT SQL_CALC_FOUND_ROWS *
		FROM pay_month_give AS pmgt
		INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no
			AND emmt.emp_month = ?
		INNER JOIN emp_master AS emtt ON emmt.emp_no = emtt.emp_no
		INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
		INNER JOIN pay_month_deduct AS pmdt ON emmt.emp_no = pmdt.de_emp_no
			AND pmdt.de_id = '1'
			AND pmdt.de_yymm = ?
		WHERE pmgt.pmg_id = '1'
			AND pmgt.pmg_yymm = ?
			AND eomt.org_company = ?
			AND pmgt.pmg_emp_name LIKE CONCAT('%', ?, '%')
	");

	PREPARE stmt FROM @v_sql;
	EXECUTE stmt USING @t_yymm1, @t_yymm2, @t_yymm3, @t_company, @t_name;
	DEALLOCATE PREPARE stmt;

	SET totalCnt := FOUND_ROWS();
END;


DROP PROCEDURE IF EXISTS USP_PERSON_INDIVIDUAL_MOD;
CREATE PROCEDURE USP_PERSON_INDIVIDUAL_MOD(
	OUT err_state INT
)
LANGUAGE SQL
NOT DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT "
AUTHOR : 허정호
DATE : 2021-08-24
DESC :
- 개인 정보 관리 > 인사기본사항 변경 처리
"
BEGIN
	DECLARE EXIT HANDLER FOR SQLEXCEPTION
	BEGIN
		ROLLBACK;
		SET err_state = -1;
	END;

	START TRANSACTION;

	INSERT INTO TranTest SET col01 = '1234';
    #INSERT INTO TranTest2 SET col012 = '1234';

    SET err_state = 1;
    COMMIT;
END;


CALL USP_PERSON_INDIVIDUAL_MOD(@err_state);
SELECT @err_state;


DROP PROCEDURE IF EXISTS USP_PERSON_INDIVIDUAL_MOD;
CREATE PROCEDURE USP_PERSON_INDIVIDUAL_MOD(
	#OUT err_state INT
)
LANGUAGE SQL
NOT DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT "
AUTHOR : 허정호
DATE : 2021-08-24
DESC :
- 개인 정보 관리 > 인사기본사항 변경 처리
"
BEGIN
	DECLARE state INT DEFAULT 0;

	DECLARE EXIT HANDLER FOR SQLEXCEPTION
	BEGIN
		ROLLBACK;
		SET state = -1;

		SELECT state;
	END;

	START TRANSACTION;

	INSERT INTO TranTest SET col01 = '1234';
    #INSERT INTO TranTest2 SET col012 = '1234';

    COMMIT;
	SET state = 1;

    SELECT state;
END;