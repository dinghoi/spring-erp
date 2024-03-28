
/*
call USP_SALES_COMPANY_PROFIT_SEL('기타사업부', '2021', '2021-01', '01');

call USP_SALES_COMPANY_PROFIT_SEL('SI1본부', '2021', '2021-01', '01');
*/

#거래처 별 손익 리스트
DROP PROCEDURE IF EXISTS USP_SALES_COMPANY_PROFIT_SEL;
CREATE PROCEDURE USP_SALES_COMPANY_PROFIT_SEL(
	IN p_saupbu varchar(30) CHARSET utf8,
	IN p_year varchar(4),
	IN p_date varchar(7),
	IN p_month varchar(2)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-11-02
DESC :
- 영업관리 > 거래처 별 손익
'
proc_body :
BEGIN
	SET @saupbu = p_saupbu;
	SET @sales_saupbu = p_saupbu;
	SET @yyyy = p_year;
	SET @cost_date = p_date;
	SET @mm = p_month;
	SET @cost_month = CONCAT(p_year, p_month);
	SET @etc_cost = 10000000; #기타 비용 기준(기준 이하 일 경우 기타 항목으로 포함)

	IF @saupbu = '기타사업부' THEN
		SET @saupbu = '';
	END IF;

	#거래처 별 비용 집계
	SET @v_sql = CONCAT("SELECT r3.company, ");
	SET @v_sql = CONCAT(@v_sql, "	SUM(r3.sales_cost) AS 'sales_cost', ");
	SET @v_sql = CONCAT(@v_sql, "	SUM(r3.company_cost) AS 'company_cost', ");
	SET @v_sql = CONCAT(@v_sql, "	SUM(r3.pay_cost) AS 'pay_cost', ");
	SET @v_sql = CONCAT(@v_sql, "	SUM(r3.general_cost) AS 'general_cost', ");
	SET @v_sql = CONCAT(@v_sql, "   SUM(r3.as_cnt) AS 'as_cnt' ");
	SET @v_sql = CONCAT(@v_sql, "FROM ( ");

	#거래처 별 상주직접비 집계 및 기타 구분
	SET @v_sql = CONCAT(@v_sql, "	SELECT CASE WHEN r2.company_cost < ", @etc_cost, " THEN '기타' ");
	SET @v_sql = CONCAT(@v_sql, "			ELSE r2.company ");
	SET @v_sql = CONCAT(@v_sql, "			END AS 'company', ");
	SET @v_sql = CONCAT(@v_sql, "			sales_cost, company_cost, ");

	SET @v_sql = CONCAT(@v_sql, "			IFNULL((SELECT SUM(cost_amt_", @mm, ") FROM company_cost ");
	SET @v_sql = CONCAT(@v_sql, "			WHERE cost_year = '", @yyyy, "' AND cost_center = '상주직접비' AND cost_id = '인건비' ");

	IF p_saupbu <> '기타사업부' THEN
		SET @v_sql = CONCAT(@v_sql, "				AND (company <> '' AND company IS NOT NULL AND company <> '공통') ");
		SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '", @saupbu, "' AND company = r2.company  ");
	ELSE
		SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '' AND company = r2.company  ");
	END IF;

	#SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '", @saupbu, "' AND company = r2.company  ");
	#SET @v_sql = CONCAT(@v_sql, "			GROUP BY company), 0) AS 'pay_cost', ");
	SET @v_sql = CONCAT(@v_sql, "			), 0) AS 'pay_cost', ");

	SET @v_sql = CONCAT(@v_sql, "			IFNULL((SELECT SUM(cost_amt_", @mm, ") FROM company_cost ");
	SET @v_sql = CONCAT(@v_sql, "			WHERE cost_year = '", @yyyy, "' AND cost_center = '상주직접비' AND cost_id <> '인건비' ");

	IF p_saupbu <> '기타사업부' THEN
		SET @v_sql = CONCAT(@v_sql, "				AND (company <> '' AND company IS NOT NULL AND company <> '공통') ");
		SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '", @saupbu, "' AND company = r2.company  ");
	ELSE
		SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '' AND company = r2.company  ");
	END IF;

	#SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '", @saupbu, "' AND company = r2.company  ");
	#SET @v_sql = CONCAT(@v_sql, "			GROUP BY company), 0) AS 'general_cost', ");
	SET @v_sql = CONCAT(@v_sql, "			), 0) AS 'general_cost', ");

	SET @v_sql = CONCAT(@v_sql, "			ifnull((select as_total - as_set from as_acpt_status AS aast ");
	SET @v_sql = CONCAT(@v_sql, "                  inner join trade as trdt on aast.as_company = trdt.trade_name ");
	SET @v_sql = CONCAT(@v_sql, "                     and trdt.saupbu = '", @saupbu, "' ");
	SET @v_sql = CONCAT(@v_sql, "                  where as_month = '", @cost_month, "' and as_company = r2.company), 0) as as_cnt ");
	SET @v_sql = CONCAT(@v_sql, "	FROM ( ");

	#거래처 별 매출, 비용 집계
	SET @v_sql = CONCAT(@v_sql, "		SELECT r1.company,  ");
	SET @v_sql = CONCAT(@v_sql, "			IFNULL(SUM(sast.cost_amt), 0) AS 'sales_cost', ");
	SET @v_sql = CONCAT(@v_sql, "           IFNULL((SELECT SUM(cost_amt_", @mm, ") FROM company_cost ");
	SET @v_sql = CONCAT(@v_sql, "			       WHERE cost_year = '", @yyyy, "' AND saupbu = @saupbu ");
	SET @v_sql = CONCAT(@v_sql, "				     AND company = r1.company AND cost_center = '상주직접비'), 0) AS company_cost ");
	SET @v_sql = CONCAT(@v_sql, "		FROM ( ");

	#비용 & 매출 거래처
	SET @v_sql = CONCAT(@v_sql, "			SELECT company ");
	SET @v_sql = CONCAT(@v_sql, "			FROM company_cost ");
	SET @v_sql = CONCAT(@v_sql, "			WHERE cost_year = '", @yyyy, "' ");
	SET @v_sql = CONCAT(@v_sql, "				AND cost_center = '상주직접비' ");

	IF p_saupbu <> '기타사업부' THEN
		SET @v_sql = CONCAT(@v_sql, "				AND (company <> '' AND company IS NOT NULL AND company <> '공통') ");
		SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '", @saupbu, "' ");
	ELSE
		SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '' ");
	END IF;

	#SET @v_sql = CONCAT(@v_sql, "				AND (company <> '' AND company IS NOT NULL AND company <> '공통') ");
	#SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '", @saupbu, "' ");

	SET @v_sql = CONCAT(@v_sql, "			GROUP BY company ");
	SET @v_sql = CONCAT(@v_sql, "           UNION ");
	SET @v_sql = CONCAT(@v_sql, "			SELECT company  ");
	SET @v_sql = CONCAT(@v_sql, "			FROM saupbu_sales ");
	SET @v_sql = CONCAT(@v_sql, "			WHERE SUBSTRING(sales_date, 1, 7) = '", @cost_date, "' ");
	SET @v_sql = CONCAT(@v_sql, "				AND saupbu = '", @sales_saupbu, "' ");
	SET @v_sql = CONCAT(@v_sql, "			GROUP BY company ");
	SET @v_sql = CONCAT(@v_sql, "		) r1 ");
	SET @v_sql = CONCAT(@v_sql, "		LEFT OUTER JOIN saupbu_sales AS sast ON r1.company = sast.company ");
	SET @v_sql = CONCAT(@v_sql, "			AND sast.saupbu = '", @sales_saupbu, "' ");
	SET @v_sql = CONCAT(@v_sql, "			AND SUBSTRING(sast.sales_date, 1, 7) = '", @cost_date, "' ");
	SET @v_sql = CONCAT(@v_sql, "		GROUP BY r1.company ");
	SET @v_sql = CONCAT(@v_sql, "	) r2 ");
	SET @v_sql = CONCAT(@v_sql, "	WHERE r2.sales_cost > 0 OR r2.company_cost > 0 ");
	SET @v_sql = CONCAT(@v_sql, ")r3 ");
	SET @v_sql = CONCAT(@v_sql, "GROUP BY r3.company ");
	SET @v_sql = CONCAT(@v_sql, "ORDER BY FIELD(r3.company, '기타') ASC, ");
	SET @v_sql = CONCAT(@v_sql, "r3.company ASC; ");

	PREPARE stmt FROM @v_sql;
	EXECUTE stmt;
	DEALLOCATE PREPARE stmt;
END;

/*
CALL USP_SALES_SAUPBU_SALES_TOTAL_SEL('SI1본부', '2021-01');
*/

#사업부 별 매출
DROP PROCEDURE IF EXISTS USP_SALES_SAUPBU_SALES_TOTAL_SEL;
CREATE PROCEDURE USP_SALES_SAUPBU_SALES_TOTAL_SEL(
	IN p_saupbu varchar(30) CHARSET utf8,
	IN p_sales_date varchar(7)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-11-02
DESC :
- 영업관리 > 사업부 별 매출
'
proc_body :
BEGIN
	SET @saupbu = p_saupbu;
	SET @sales_date = p_sales_date;

	SELECT SUM(cost_amt) AS 'sales_total'
	FROM saupbu_sales
	WHERE SUBSTRING(sales_date, 1, 7) = @sales_date
		AND saupbu = @saupbu;
END;

