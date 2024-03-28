
# SHOW PROCEDURE STATUS;

# 발령 사항
DROP PROCEDURE USP_INSA_APPOINT_INFO;
CREATE PROCEDURE USP_INSA_APPOINT_INFO(
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
- 발령 사항 '
proc_body :
BEGIN
	SELECT app_date, app_id, app_id_type, app_to_company, app_to_orgcode
		, app_to_org, app_to_grade, app_to_position, app_be_company, app_be_orgcode
		, app_be_org, app_be_grade, app_be_position, app_start_date, app_finish_date
		, app_be_enddate, app_reward, app_comment
	FROM emp_appoint
	WHERE app_empno = p_emp_no
	ORDER BY app_empno, app_date, app_seq ASC;
END;

# 경력사항 조회
DROP PROCEDURE USP_INSA_CAREER_INFO;
CREATE PROCEDURE USP_INSA_CAREER_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-17
DESC : 경력사항 조회'
proc_body :
BEGIN
	SELECT career_task, career_join_date, career_end_date, career_office, career_dept
	, career_position
	FROM emp_career
	WHERE career_empno = p_emp_no
	ORDER BY career_empno, career_seq ASC;
END;

# 교육 사항
DROP PROCEDURE USP_INSA_EDU_INFO;
CREATE PROCEDURE USP_INSA_EDU_INFO(
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
- 교육 사항'
proc_body :
BEGIN
	SELECT edu_name, edu_office, edu_finish_no, edu_start_date, edu_end_date, edu_comment
	FROM emp_edu
	WHERE edu_empno = p_emp_no
	ORDER BY edu_empno,edu_seq ASC;
END;

# 가족 사항
DROP PROCEDURE USP_INSA_FAMILY_INFO;
CREATE PROCEDURE USP_INSA_FAMILY_INFO(
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
- 가족 사항'
proc_body :
BEGIN
	SELECT family_rel, family_name, family_birthday, family_birthday_id, family_job
		, family_tel_ddd, family_tel_no1, family_tel_no2, family_person1, family_person2
		, family_live
	FROM emp_family
	WHERE family_empno = p_emp_no
	ORDER BY family_empno, family_seq ASC;

END;

# 학력사항 조회
DROP PROCEDURE USP_INSA_SCHOOL_INFO;
CREATE PROCEDURE USP_INSA_SCHOOL_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-17
DESC : 학력사항 조회'
proc_body :
BEGIN
	SELECT sch_start_date, sch_end_date, sch_school_name, sch_dept, sch_major
		, sch_sub_major, sch_degree, sch_finish
	FROM emp_school
	WHERE sch_empno = p_emp_no;
END;

# 자격증 사항 조회
DROP PROCEDURE USP_INSA_QUAL_INFO;
CREATE PROCEDURE USP_INSA_QUAL_INFO(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-08-18
DESC : 자격증 사항 조회'
proc_body :
BEGIN
	SELECT qual_pay_id, qual_type, qual_grade, qual_pass_date, qual_org
	, qual_no, qual_passport
	FROM emp_qual
	WHERE qual_empno = p_emp_no
	ORDER BY qual_empno, qual_seq ASC;
END;

# 조직 현황 리스트
DROP PROCEDURE IF EXISTS USP_INSA_ORG_MST_LIST;
CREATE PROCEDURE USP_INSA_ORG_MST_LIST(
	IN p_company VARCHAR(30) CHARSET utf8,
	IN p_orgType VARCHAR(25),
	IN p_search VARCHAR(50) CHARSET utf8,
	IN p_stpage INT,
	IN p_pgsize INT
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-09-03
DESC :
- 조직 현황 리스트
'
proc_body :
BEGIN
	SET @v_company = p_company;
	SET @v_orgType = p_orgType;
	SET @v_stpage = p_stpage;
	SET @v_pgsize = p_pgsize;

	#회사 검색 조건
	IF @v_company <> '전체' THEN
		SET @v_condi = CONCAT("AND org_company = '", @v_company, "' ");
	ELSE
		SET @v_condi = "";
	END IF;

	#조직 구분 검색 조건
	IF @v_orgType = 'bonbu' THEN
		SET @v_condi = CONCAT(@v_condi, "AND org_bonbu LIKE '%", p_search, "%' ");
	ELSEIF @v_orgType = 'saupbu' THEN
		SET @v_condi = CONCAT(@v_condi, "AND org_saupbu LIKE '%", p_search, "%' ");
	ELSEIF @v_orgType = 'team' THEN
		SET @v_condi = CONCAT(@v_condi, "AND org_team LIKE '%", p_search, "%' ");
	ELSEIF @v_orgType = 'org_name' THEN
		SET @v_condi = CONCAT(@v_condi, "AND org_name LIKE '%", p_search, "%' ");
	ELSEIF @v_orgType = 'reside_company' THEN
		SET @v_condi = CONCAT(@v_condi, "AND org_reside_company LIKE '%", p_search, "%' ");
	ELSEIF @v_orgType = 'org_code' THEN
		SET @v_condi = CONCAT(@v_condi, "AND org_code ='", p_search, "' ");
	ELSE
		SET @v_condi = CONCAT(@v_condi, "");
	END IF;

	#Total Count
	SET @v_cnt_query = "SELECT COUNT(*) INTO @v_total
	FROM emp_org_mst
	WHERE (ISNULL(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '0000-00-00')
	";

	SET @v_cnt_sql = CONCAT(@v_cnt_query, @v_condi);

	PREPARE c_stmt FROM @v_cnt_sql;
	EXECUTE c_stmt;
	DEALLOCATE PREPARE c_stmt;

	#Page List
	SET @v_sql = CONCAT("SELECT ", @v_total, ",
		org_code, org_name, org_level, org_table_org, org_empno,
		org_emp_name, org_company, org_bonbu, org_saupbu, org_team,
		org_reside_company, org_date, org_owner_empno, org_owner_empname,
		org_reside_place, trade_code
	FROM emp_org_mst
	WHERE (ISNULL(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '0000-00-00') ",
	@v_condi,
	"ORDER BY FIELD(org_level, '회사', '본부', '사업부', '팀', '상주처', '파트') ASC,
		org_code, org_company, org_bonbu, org_team, org_reside_place
	LIMIT ", @v_stpage, ", ", @v_pgsize);

	PREPARE stmt FROM @v_sql;
	EXECUTE stmt;
	DEALLOCATE PREPARE stmt;
END;

#차량관리 > 엑셀 리스트
DROP PROCEDURE IF EXISTS USP_INSA_CAR_INFO_SELECT;
CREATE PROCEDURE USP_INSA_CAR_INFO_SELECT(
	IN p_owner_view varchar(10),
	IN p_field_check varchar(20) CHARSET utf8,
	IN p_field_view varchar(250) CHARSET utf8
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-07
DESC :인사 > 차량관리 > 엑셀 리스트
'
proc_body :
BEGIN
	SET @v_gubun = p_owner_view;
	SET @v_field = p_field_check;
	SET @v_search = p_field_view;

	IF @v_gubun = 'C' THEN
		SET @v_condi = CONCAT("AND car_owner = '회사' ");
	ELSEIF @v_gubun = 'P' THEN
		SET @v_condi = CONCAT("AND car_owner = '개인' ");
	ELSEIF @v_gubun = 'T' THEN
		SET @v_condi = CONCAT("AND (car_owner = '개인' OR car_owner = '회사') ");
	END IF;

	IF @v_field = 'buy_gubun' THEN
		SET @v_condi = CONCAT(@v_condi, "AND buy_gubun LIKE '%", @v_search, "%' ");
	ELSEIF @v_field = 'owner_emp_name' THEN
		SET @v_condi = CONCAT(@v_condi, "AND owner_emp_name LIKE '%", @v_search, "%' ");
	ELSEIF @v_field = 'oil_kind' THEN
		SET @v_condi = CONCAT(@v_condi, "AND oil_kind LIKE '%", @v_search, "%' ");
	ELSEIF @v_field = 'car_no' THEN
		SET @v_condi = CONCAT(@v_condi, "AND car_no LIKE '%", @v_search, "%' ");
	END IF;

	SET @v_sql = CONCAT("SELECT car_no, car_name, car_year, oil_kind, car_owner, ");
	SET @v_sql = CONCAT(@v_sql, "car_company, car_use_dept, car_use, owner_emp_name, owner_emp_no, ");
	SET @v_sql = CONCAT(@v_sql, "car_reg_date, last_km, insurance_date, insurance_company, insurance_amt, ");
	SET @v_sql = CONCAT(@v_sql, "last_check_date, car_status, car_comment ");
	SET @v_sql = CONCAT(@v_sql, "FROM car_info ");
	SET @v_sql = CONCAT(@v_sql, "WHERE (end_date = '' OR end_date IS NULL OR end_date = '1900-01-01') ");
	SET @v_sql = CONCAT(@v_sql, @v_condi, "ORDER BY car_no DESC ");

	PREPARE stmt FROM @v_sql;
	EXECUTE stmt;
	DEALLOCATE PREPARE stmt;
END

