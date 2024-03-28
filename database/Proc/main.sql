#메인 게시판 리스트

DROP PROCEDURE IF EXISTS USP_MAIN_BOARD_LIST;
CREATE PROCEDURE USP_MAIN_BOARD_LIST(
	IN p_gubun VARCHAR(2),
	IN p_type VARCHAR(50),
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
- 메인 게시판 리스트
'
proc_body :
BEGIN
	SET @v_gubun = p_gubun;
	SET @v_type = p_type;
	SET @v_stpage = p_stpage;
	SET @v_pgsize = p_pgsize;

	#회사 검색 조건
	IF @v_gubun <> '0' THEN
		SET @v_condi = CONCAT("WHERE board_gubun = '", @v_gubun, "' ");
		SET @v_condi = CONCAT(@v_condi, " AND ");
	ELSE
		SET @v_condi = "WHERE ";
	END IF;

	#조직 구분 검색 조건
	IF @v_type = 'board_title' THEN
		SET @v_condi = CONCAT(@v_condi, "board_title LIKE '%", p_search, "%' ");
	ELSEIF @v_type = 'board_body' THEN
		SET @v_condi = CONCAT(@v_condi, "board_body LIKE '%", p_search, "%' ");
	ELSEIF @v_type = 'reg_name' THEN
		SET @v_condi = CONCAT(@v_condi, "reg_name LIKE '%", p_search, "%' ");
	ELSE
		SET @v_condi = CONCAT(@v_condi, "board_title LIKE '%", p_search, "%' ");
		SET @v_condi = CONCAT(@v_condi, "OR board_body LIKE '%", p_search, "%' ");
	END IF;

	#Total Count
	SET @v_cnt_query = "SELECT COUNT(*) INTO @v_total FROM board ";

	SET @v_cnt_sql = CONCAT(@v_cnt_query, @v_condi);

	PREPARE c_stmt FROM @v_cnt_sql;
	EXECUTE c_stmt;
	DEALLOCATE PREPARE c_stmt;

	#Page List
	SET @v_sql = CONCAT("SELECT ", @v_total, ",
		board_seq, board_gubun, reg_name, board_title, reg_date, read_cnt, att_file
	FROM board ",
	@v_condi,
	"ORDER BY reg_date DESC
	LIMIT ", @v_stpage, ", ", @v_pgsize);

	PREPARE stmt FROM @v_sql;
	EXECUTE stmt;
	DEALLOCATE PREPARE stmt;
END;
