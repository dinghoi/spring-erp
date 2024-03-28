# SHOW PROCEDURE STATUS;


# 인사 기타 정보 > 교육 사항
DROP PROCEDURE IF EXISTS USP_PERSON_LANGUAGE_INFO;
CREATE PROCEDURE USP_PERSON_LANGUAGE_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-20
DESC :
- 인사 기타 정보 > 교육 사항'
proc_body :
BEGIN
	SELECT lang_id, lang_id_type, lang_point, lang_grade, lang_get_date
	FROM emp_language
	WHERE lang_empno = p_emp_no
	ORDER BY lang_empno, lang_seq ASC;
END;

# 개인 인사 정보 조회
DROP PROCEDURE IF EXISTS USP_PERSON_INSA_INFO;
CREATE PROCEDURE USP_PERSON_INSA_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-12
DESC :
- 개인 인사 정보 조회'
proc_body :
BEGIN
	SELECT emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, emtt.emp_in_date,
		emtt.emp_first_date, emtt.emp_reside_place, emtt.emp_birthday, emtt.emp_type, emtt.emp_org_baldate,
		emtt.emp_grade_date, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team
	FROM emp_master AS emtt
	INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
	WHERE emtt.emp_no = p_emp_no;
END;

# 개인 인사 기록 카드 조회
DROP PROCEDURE IF EXISTS USP_PERSON_CARD_VIEW;
CREATE PROCEDURE USP_PERSON_CARD_VIEW(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-12
DESC :
- 개인 인사 기록 카드 조회'
proc_body :
BEGIN
	SELECT emtt.emp_name, emtt.emp_org_code, emtt.emp_jikgun, emtt.emp_jikmu, emtt.emp_person1,
 		emtt.emp_person2, emtt.emp_position, emtt.emp_grade, emtt.emp_job, emtt.emp_image,
 		emtt.emp_military_date1, emtt.emp_military_date2, emtt.emp_marry_date, emtt.emp_grade_date, emtt.emp_end_date,
 		emtt.emp_org_baldate, emtt.emp_sawo_date, emtt.emp_first_date, emtt.emp_in_date, emtt.emp_tel_ddd,
 		emtt.emp_tel_no1, emtt.emp_tel_no2, emtt.emp_hp_ddd, emtt.emp_hp_no1, emtt.emp_hp_no2,
 		emtt.emp_ename, emtt.emp_sido, emtt.emp_gugun, emtt.emp_dong, emtt.emp_addr,
 		emtt.emp_gunsok_date, emtt.emp_end_gisan, emtt.emp_email, emtt.emp_faith, emtt.emp_military_id,
 		emtt.emp_military_grade, emtt.emp_military_comm
	FROM emp_master AS emtt
	WHERE emtt.emp_no = p_emp_no;
END;

# 자격사항 정보 조회
DROP PROCEDURE USP_PERSON_CARD_CAREER_INFO;
CREATE PROCEDURE USP_PERSON_CARD_CAREER_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-12
DESC :
- 자격사항 정보 조회'
proc_body :
BEGIN
	SELECT career_join_date, career_end_date, career_office,
		career_dept, career_position, career_task
	FROM emp_career
	WHERE career_empno = p_emp_no
	ORDER BY career_empno, career_seq ASC
	LIMIT 2;
END;

# 자격사항 정보 조회
DROP PROCEDURE USP_PERSON_CARD_QUAL_INFO;
CREATE PROCEDURE USP_PERSON_CARD_QUAL_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-12
DESC :
- 자격사항 정보 조회'
proc_body :
BEGIN
	SELECT qual_type, qual_grade, qual_pass_date, qual_org, qual_no
	FROM emp_qual
	WHERE qual_empno = p_emp_no
	ORDER BY qual_empno, qual_seq ASC
	LIMIT 3;
END;

# 학력사항 정보 조회
DROP PROCEDURE USP_PERSON_CARD_SCHOOL_INFO;
CREATE PROCEDURE USP_PERSON_CARD_SCHOOL_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-12
DESC :
- 학력사항 정보 조회'
proc_body :
BEGIN
	 SELECT sch_start_date, sch_end_date, sch_school_name, sch_dept, sch_major,
	 	sch_sub_major, sch_degree, sch_finish
	 FROM emp_school
	 WHERE sch_empno = p_emp_no
	 ORDER BY sch_empno, sch_seq ASC
	 LIMIT 2;
END;

# 인사기본사항 변경 조회
DROP PROCEDURE IF EXISTS USP_PERSON_INDIVIDUAL_INFO;
CREATE PROCEDURE USP_PERSON_INDIVIDUAL_IN(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-20
DESC :
- 인사기본사항 변경 조회'
proc_body :
BEGIN

	SELECT emp_name, emp_ename, emp_type, emp_sex, emp_person1
		, emp_person2, emp_image, emp_first_date, emp_in_date, emp_gunsok_date
		, emp_yuncha_date, emp_end_gisan
		, IF(emp_end_date = '1900-01-01' OR emp_end_date IS NULL, '', emp_end_date) AS 'end_date'
		, emp_company, emp_bonbu
		, emp_saupbu, emp_team, emp_org_code, emp_org_name
		, IF(emp_org_baldate = '1900-01-01' OR emp_org_baldate IS NULL, '', emp_org_baldate) AS 'org_baldate'
		, emp_stay_code, emp_reside_place, emp_reside_company, emp_grade
		, IF(emp_grade_date = 1900-01-01 OR emp_grade_date IS NULL, '', emp_grade_date) AS 'grade_date'
		, emp_job, emp_position, emp_jikgun, emp_jikmu
		, IF(emp_birthday = '1900-01-01' OR emp_birthday IS NULL, '', emp_birthday) AS 'birthday'
		, emp_birthday_id, emp_family_zip, emp_family_sido, emp_family_gugun
		, emp_family_dong, emp_family_addr, emp_zipcode, emp_sido, emp_gugun
		, emp_dong, emp_addr, emp_tel_ddd, emp_tel_no1, emp_tel_no2
		, emp_hp_ddd, emp_hp_no1, emp_hp_no2, emp_email, emp_military_id
		, IF(emp_military_date1 = '1900-01-01' OR emp_military_date1 IS NULL, ''
			, emp_military_date1) AS 'military_date1'
		, IF(emp_military_date2 = '1900-01-01' OR emp_military_date2 IS NULL, ''
			, emp_military_date2) AS 'military_date2'
		, emp_military_grade, emp_military_comm, emp_hobby, emp_faith
		, emp_last_edu
		, IF(emp_marry_date = '1900-01-01' OR emp_marry_date IS NULL, '', emp_marry_date) AS 'marry_date'
		, emp_disabled, emp_disab_grade, emp_sawo_id
		, IF(emp_sawo_date = '1900-01-01' OR emp_sawo_date IS NULL, '', emp_sawo_date) AS 'sawo_date'
		, emp_emergency_tel, emp_nation_code, emp_extension_no, emp_reg_user
		, emp_mod_user
	FROM emp_master
	WHERE emp_no = p_emp_no;
END;

# 인사 기타 정보 > 발령 사항
DROP PROCEDURE USP_PERSON_CARD_APPOINT_INFO;
CREATE PROCEDURE USP_PERSON_CARD_APPOINT_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-19
DESC :
- 인사 기타 정보 > 발령 사항'
proc_body :
BEGIN
	SELECT app_date, app_id, app_id_type, app_to_company, app_to_orgcode
		, app_to_org, app_to_grade, app_to_job, app_to_position, app_to_enddate
		, app_be_company, app_be_orgcode, app_be_org, app_be_grade, app_be_job
		, app_be_position, app_be_enddate, app_start_date, app_finish_date, app_reward
		, app_comment
	FROM emp_appoint
	WHERE app_empno = p_emp_no
	ORDER BY app_empno, app_seq ASC
	LIMIT 2;
END;

# 인사 기타 정보 > 교육 사항
DROP PROCEDURE USP_PERSON_CARD_EDU_INFO;
CREATE PROCEDURE USP_PERSON_CARD_EDU_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-20
DESC :
- 인사 기타 정보 > 교육 사항'
proc_body :
BEGIN
	SELECT edu_name, edu_office, edu_finish_no, edu_start_date, edu_end_date
		, edu_comment
	FROM emp_edu
	WHERE edu_empno = p_emp_no
	ORDER BY edu_empno, edu_seq ASC
	LIMIT 2;
END;

# 가족 사항
DROP PROCEDURE USP_PERSON_CARD_FAMILY_INFO;
CREATE PROCEDURE USP_PERSON_CARD_FAMILY_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-19
DESC :
- 가족 사항'
proc_body :
BEGIN
	SELECT family_rel, family_name, family_birthday, family_birthday_id, family_job
		, family_person1, family_person2, family_live
	FROM emp_family
	WHERE family_empno = p_emp_no
	ORDER BY family_empno, family_seq ASC
	LIMIT 2;
END;

# 인사기록카드-기타사항
DROP PROCEDURE USP_PERSON_CARD_ETC_INFO;
CREATE PROCEDURE USP_PERSON_CARD_ETC_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-19
DESC :
- 인사기록카드-기타사항'
proc_body :
BEGIN
	SELECT emtt.emp_stay_name, emtt.emp_stay_code, emtt.emp_type, emtt.emp_company, emtt.emp_bonbu
		, emtt.emp_saupbu, emtt.emp_team, emtt.emp_reside_place, emtt.emp_yuncha_date, emtt.emp_sawo_id
		, emtt.emp_disabled, emtt.emp_disab_grade, emtt.emp_hobby, emtt.emp_birthday_id, emtt.emp_family_gugun
		, emtt.emp_family_dong, emtt.emp_family_addr, emtt.emp_emergency_tel, eomt.org_company, eomt.org_bonbu
		, eomt.org_saupbu, eomt.org_team, eomt.org_name, eomt.org_reside_place, emtt.emp_family_sido
		, IF(emtt.emp_end_date, '1900-01-01', '') AS 'end_date'
		, IF(emtt.emp_birthday, '1900-01-01', '') AS 'birthday'
		, IF(emtt.emp_grade_date, '1900-01-01', '') AS 'grade_date'
		, IF(emtt.emp_org_baldate, '1900-01-01', '') AS 'org_baldate'
		, IF(emtt.emp_sawo_date, '1900-01-01', '') AS 'sawo_date'
		, emyt.stay_name, emyt.stay_sido, emyt.stay_gugun, emyt.stay_dong, emyt.stay_addr
		, emtt.emp_disabled_yn, emtt.emp_org_code
	FROM emp_master AS emtt
	INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
	LEFT OUTER JOIN emp_stay AS emyt ON emtt.emp_stay_code = emyt.stay_code
	WHERE emp_no = p_emp_no;
END;

# 개인 인사기본사항 업데이트
DROP PROCEDURE IF EXISTS USP_PERSON_INDIVIDUAL_UPDATE;
CREATE PROCEDURE USP_PERSON_INDIVIDUAL_UPDATE(
	IN p_emp_no VARCHAR(6),
	IN p_emp_ename VARCHAR(30),
	IN p_emp_birthday DATE,
	IN p_emp_birthday_id VARCHAR(2) CHARSET utf8,
	IN p_emp_family_zip VARCHAR(12),
	IN p_emp_family_sido VARCHAR(10) CHARSET utf8,
	IN p_emp_family_gugun VARCHAR(20) CHARSET utf8,
	IN p_emp_family_dong VARCHAR(50) CHARSET utf8,
	IN p_emp_family_addr VARCHAR(50) CHARSET utf8,
	IN p_emp_zipcode VARCHAR(12),
	IN p_emp_sido VARCHAR(10) CHARSET utf8,
	IN p_emp_gugun VARCHAR(20) CHARSET utf8,
	IN p_emp_dong VARCHAR(50) CHARSET utf8,
	IN p_emp_addr VARCHAR(50) CHARSET utf8,
	IN p_emp_tel_ddd VARCHAR(3),
	IN p_emp_tel_no1 VARCHAR(4),
	IN p_emp_tel_no2 VARCHAR(4),
	IN p_emp_hp_ddd VARCHAR(3),
	IN p_emp_hp_no1 VARCHAR(4),
	IN p_emp_hp_no2 VARCHAR(4),
	IN p_emp_email VARCHAR(30),
	IN p_emp_military_id VARCHAR(10) CHARSET utf8,
	IN p_emp_military_date1 DATE,
	IN p_emp_military_date2 DATE,
	IN p_emp_military_grade VARCHAR(10) CHARSET utf8,
	IN p_emp_military_comm VARCHAR(20) CHARSET utf8,
	IN p_emp_hobby VARCHAR(20) CHARSET utf8,
	IN p_emp_faith VARCHAR(20) CHARSET utf8,
	IN p_emp_last_edu VARCHAR(30) CHARSET utf8,
	IN p_emp_marry_date DATE,
	IN p_emp_emergency_tel VARCHAR(13),
	IN p_emp_extension_no VARCHAR(14),
	IN p_fine_name VARCHAR(100) CHARSET utf8,
	IN p_emp_name VARCHAR(20) CHARSET utf8,
	IN p_hp VARCHAR(20),
	IN p_email VARCHAR(50),
	IN p_dt5_now DATETIME(0)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-24
DESC :
- 개인 정보 관리 > 인사기본사항 업데이트

RETURN VALUE : state
	0 = 에러가 없습니다.
	-1 = 예상치 못한 런타임 오류가 발생하였습니다.
'
proc_body :
BEGIN
	# ERROR LOG PARAM
	DECLARE v_vch_proc_name VARCHAR(100) DEFAULT 'USP_PERSON_INDIVIDUAL_UPDATE';
    DECLARE v_iny_proc_step TINYINT UNSIGNED DEFAULT 0;
    DECLARE v_vch_sql_state VARCHAR(5);
    DECLARE v_int_error_no INT;
    DECLARE v_txt_error_msg TEXT;

	# 처리 상태 값
	DECLARE state INT DEFAULT 0;

	# 예외 처리 선언
	DECLARE EXIT HANDLER FOR SQLEXCEPTION
	BEGIN
		GET DIAGNOSTICS CONDITION 1 v_vch_sql_state = RETURNED_SQLSTATE
            , v_int_error_no = MYSQL_ERRNO, v_txt_error_msg = MESSAGE_TEXT;

        # ROLLBACK 처리
		ROLLBACK;

		# error_log input
		CALL USP_ERROR_LOG_INPUT(v_vch_proc_name, v_iny_proc_step, v_vch_sql_state, v_int_error_no
			, v_txt_error_msg, p_dt5_now);

		# 상태값 설정
		SET state = -1;

		SELECT state;
	END;

	# 파일 이미지 변수 선언
	SET @emp_image = p_fine_name;

	# 파일  이미지 체크(기존 파일 이미지 유지 체크)
	IF @emp_image = "" THEN
		SELECT emp_image INTO @emp_image
		FROM emp_master
		WHERE emp_no = p_emp_no;
	END IF;

	# 트랜잭션 시작
	START TRANSACTION;

		# emp_master 정보 수정
		UPDATE emp_master SET
			emp_ename = p_emp_ename
			, emp_image = @emp_image
			, emp_birthday = p_emp_birthday
			, emp_birthday_id = p_emp_birthday_id
			, emp_family_zip = p_emp_family_zip
			, emp_family_sido = p_emp_family_sido
			, emp_family_gugun = p_emp_family_gugun
			, emp_family_dong = p_emp_family_dong
			, emp_family_addr = p_emp_family_addr
			, emp_zipcode = p_emp_zipcode
			, emp_sido = p_emp_sido
			, emp_gugun = p_emp_gugun
			, emp_dong = p_emp_dong
			, emp_addr = p_emp_addr
			, emp_tel_ddd = p_emp_tel_ddd
			, emp_tel_no1 = p_emp_tel_no1
			, emp_tel_no2 = p_emp_tel_no2
			, emp_hp_ddd = p_emp_hp_ddd
			, emp_hp_no1 = p_emp_hp_no1
			, emp_hp_no2 = p_emp_hp_no2
			, emp_email = p_emp_email
			, emp_military_id = p_emp_military_id
			, emp_military_date1 = p_emp_military_date1
			, emp_military_date2 = p_emp_military_date2
			, emp_military_grade = p_emp_military_grade
			, emp_military_comm = p_emp_military_comm
			, emp_hobby = p_emp_hobby
			, emp_faith = p_emp_faith
			, emp_last_edu = p_emp_last_edu
			, emp_marry_date = p_emp_marry_date
			, emp_emergency_tel = p_emp_emergency_tel
			, emp_extension_no = p_emp_extension_no
			, emp_mod_user = p_emp_name
			, emp_mod_date = NOW(0)
		WHERE emp_no = p_emp_no;

		# memb 정보 수정
		UPDATE memb SET
			hp = p_hp,
			email = p_email,
			mod_id = p_emp_no,
			mod_date = NOW(0)
		WHERE user_id = p_emp_no;

    # COMMIT 처리
    COMMIT;

    # 상태값 설정
    SET state = 0;

    SELECT state;
END;

#주소(동코드) 검색
DROP PROCEDURE IF EXISTS USP_COMM_ZIPCODE_INFO;
CREATE PROCEDURE USP_COMM_ZIPCODE_INFO(
	IN p_dong VARCHAR(20) CHARSET utf8
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-27
DESC :
- 주소(동코드) 검색
'
proc_body :
BEGIN
	SET @t_dong = p_dong;
	SET @t_mg_group = '1';

	SET @v_sql = CONCAT("
		SELECT zipcode, sido, gugun, dong
		FROM area_mg
		WHERE mg_group = ?
	");

	IF p_dong = '' THEN
		SET @v_sql = concat(@v_sql, "AND dong = ?");
	ELSE
		SET @v_sql = concat(@v_sql, "AND dong LIKE CONCAT('%', ?, '%')");
	END IF;

	PREPARE stmt FROM @v_sql;
	EXECUTE stmt USING @t_mg_group, @t_dong;
	DEALLOCATE PREPARE stmt;
END;


#가족 사항 조회
DROP PROCEDURE IF EXISTS USP_PERSON_FAMILY_LIST;
CREATE PROCEDURE USP_PERSON_FAMILY_LIST(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-27
DESC :
- 가족 사항 조회
'
proc_body :
BEGIN
	SELECT family_empno, family_seq, family_rel, family_name, family_birthday,
		family_birthday_id, family_job, family_person1, family_person2, family_tel_ddd,
		family_tel_no1, family_tel_no2,	family_live
	FROM emp_family
	WHERE family_empno = p_emp_no
	ORDER BY family_empno, family_seq ASC;
END;
