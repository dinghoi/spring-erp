
#비용마감관리 리스트
DROP PROCEDURE IF EXISTS USP_COST_END_MG_SEL;
CREATE PROCEDURE USP_COST_END_MG_SEL(
	IN p_emp_no varchar(6),
	IN p_cost_grade varchar(1)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 비용마감 > 비용마감관리 리스트
'
proc_body :
BEGIN
	SET @emp_no = p_emp_no;
	SET @cost_grade = p_cost_grade;
				
	SELECT MAX(org_month) INTO @org_month
	FROM emp_org_mst_month;
	
	IF @org_month = NULL OR @org_month = '' THEN 
		SET @v_max_month = '000000';
	END IF; 
	
	IF @cost_grade <> '0' THEN 
		SELECT eomt.org_bonbu INTO @org_bonbu
		FROM emp_master_month AS emmt 
		INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code 
		WHERE emmt.emp_month = @org_month
			AND emmt.emp_no = @emp_no;
	END IF;

	SET @v_sql = CONCAT("SELECT @v_max_month AS 'max_org_month', org_name, org_date FROM emp_org_mst ");
	SET @v_sql = CONCAT(@v_sql, "WHERE org_level = '본부' ");
	SET @v_sql = CONCAT(@v_sql, "AND (ISNULL(org_end_date) OR org_end_date = '0000-00-00') ");
	SET @v_sql = CONCAT(@v_sql, "AND org_name NOT IN ('전략부문', 'ICT연구소', '빅데이타연구소', ");
	SET @v_sql = CONCAT(@v_sql, "'기술연구소', '한진그룹사업본부')");
	
	IF @cost_grade = '0' THEN 
		SET @v_sql = CONCAT(@v_sql, "GROUP BY org_bonbu, org_name ");
		SET @v_sql = CONCAT(@v_sql, "ORDER BY FIELD(org_company, '케이원', '케이네트웍스', '케이시스템'), ");
		SET @v_sql = CONCAT(@v_sql, "FIELD(org_bonbu, '스마트본부', '공공SI본부', '금융SI본부', 'ICT본부', ");
		SET @v_sql = CONCAT(@v_sql, "'공공본부', 'NI본부', 'SI2본부', 'SI1본부') DESC ");
	ELSE
		SET @v_sql = CONCAT(@v_sql, "AND (org_name = '", @org_bonbu,"' Or org_empno = '", @emp_no,"') ");
		SET @v_sql = CONCAT(@v_sql, "GROUP BY org_name ");
	END IF;
	
	PREPARE stmt FROM @v_sql;
	EXECUTE stmt;
	DEALLOCATE PREPARE stmt;
END;


#비용마감관리 > 영업부서 별 리스트
DROP PROCEDURE IF EXISTS USP_COST_END_ORG_SEL;
CREATE PROCEDURE USP_COST_END_ORG_SEL(
	IN p_org_name varchar(30) CHARSET utf8
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 비용마감 > 비용마감관리 > 영업부서 별 리스트
'
proc_body :
BEGIN
	SET @org_name = p_org_name;
				
	SELECT MAX(end_month) INTO @end_month
	FROM cost_end
	WHERE saupbu = @org_name;
	
	SELECT end_month, end_yn, reg_name, reg_id, reg_date, 
		batch_yn, ceo_yn, bonbu_yn
	FROM cost_end
	WHERE saupbu = @org_name
		AND end_month = @end_month;
END;

