
/*
# 유류비 단가 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_OIL_UNIT_SEL;
CREATE PROCEDURE USP_ORG_END_OIL_UNIT_SEL(
	IN p_emp_month varchar(6)
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
- 사업부별 비용마감 > 유류비 단가 조회
'
proc_body :
BEGIN
	SELECT oil_unit_id
	FROM oil_unit
	WHERE oil_unit_month = p_emp_month;
END;


/*==================================================*/

/*
# 유류비 단가 및 계산(개인)
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_OIL_SEL;
CREATE PROCEDURE USP_ORG_END_OIL_SEL(
	IN p_end_month varchar(6),
	IN p_from_date varchar(10),
	IN p_to_date varchar(10),
	IN p_dept_name varchar(30) CHARSET utf8
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
- 사업부별 비용마감 > 유류비 단가 및 계산(개인)
'
proc_body :
BEGIN
	SELECT o1.mg_ce_id, o1.run_date, o1.run_seq, o1.far, o1.liter,
		(select oil_unit_average
		from oil_unit
		where oil_unit_month = p_end_month
			and oil_unit_id = o1.oil_unit_id
			and oil_kind = o1.oil_kind) as 'oil_unit_average'
	FROM (
	SELECT trct.mg_ce_id,
		trct.oil_kind,
		trct.far,
		trct.run_date,
		trct.run_seq,
		case when eomt.org_team = '본사팀' or eomt.org_team = 'Repair팀' then '1'
		else '2'
		end AS 'oil_unit_id',
		case when trct.oil_kind = '가스' then '7'
		else
			case when emmt.emp_reside_company = '한화화약' then '8'
			else '10'
			END
		end AS 'liter'
	FROM transit_cost AS trct
	INNER JOIN emp_master_month AS emmt ON trct.mg_ce_id = emmt.emp_no
		AND emp_month = p_end_month
	INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code
		AND eomt.org_month = p_end_month
		AND (ISNULL(eomt.org_end_date) OR eomt.org_end_date = '0000-00-00')
	WHERE (trct.run_date >= p_from_date AND trct.run_date <= p_to_date)
		AND trct.car_owner ='개인'
		AND trct.far > 0
		AND eomt.org_bonbu = p_dept_name
	) o1;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_OIL_UPDATE;
# 유류 단가 업데이트
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_OIL_UP;
CREATE PROCEDURE USP_ORG_END_OIL_UP(
	IN p_mg_ce_id varchar(20),
	IN p_run_date varchar(10),
	IN p_run_seq varchar(2),
	IN p_oil_unit int(10),
	IN p_oil_price int(10)
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
- 사업부별 비용마감 > 유류 단가 업데이트
'
proc_body :
BEGIN
	UPDATE transit_cost SET
		oil_unit = p_oil_unit,
		oil_price = p_oil_price
	WHERE mg_ce_id = p_mg_ce_id
		AND run_date = p_run_date
		AND run_seq = p_run_seq;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_EMP_SELECT;
#전사 직원 정보
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_PERSON_ORG_SEL;
CREATE PROCEDURE USP_ORG_END_PERSON_ORG_SEL(
	IN p_end_month varchar(6),
	IN p_start_date varchar(10),
	IN p_from_date varchar(10),
	IN p_dept_name varchar(30) CHARSET utf8
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
- 사업부별 비용마감 > 전사 직원 정보
'
proc_body :
BEGIN
	SELECT eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team,
		eomt.org_name, emmt.emp_reside_place, emmt.emp_reside_company,
		emmt.emp_no, emmt.emp_end_date, emmt.emp_name, emmt.emp_job,
		case when eomt.org_team = '본사팀' or eomt.org_team = 'Repair팀' then '1'
		else '2'
		end AS 'oil_unit_id',
		case when emmt.emp_reside_company = '한화화약' then '8'
		else '10'
		end	AS 'liter'
	FROM emp_master_month AS emmt
	INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code
		AND eomt.org_month = p_end_month
	WHERE emmt.emp_month = p_end_month
		AND (emmt.emp_end_date = '1900-01-01' OR ISNULL(emmt.emp_end_date) OR emmt.emp_end_date >= p_start_date)
		AND eomt.org_bonbu = p_dept_name;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_ORG_TRAN_UPDATE;
DROP PROCEDURE IF EXISTS USP_COST_END_ORG_OVER_UPDATE;
DROP PROCEDURE IF EXISTS USP_COST_END_ORG_CARD_UPDATE;
#비용 별 조직 업데이트
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_COST_ORG_UP;
CREATE PROCEDURE USP_ORG_END_COST_ORG_UP(
	IN p_from_date varchar(10),
	IN p_to_date varchar(10),
	IN p_mg_ce_id varchar(20),
	IN p_org_company varchar(30) CHARSET utf8,
	IN p_org_bonbu varchar(30) CHARSET utf8,
	IN p_org_saupbu varchar(30) CHARSET utf8,
	IN p_org_team varchar(30) CHARSET utf8,
	IN p_org_name varchar(30) CHARSET utf8,
	IN p_emp_reside_place varchar(30) CHARSET utf8,
	IN p_emp_reside_company varchar(50)  CHARSET utf8
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
- 사업부별 비용마감 > 조직 별 비용 업데이트
'
proc_body :
BEGIN
	#교통비
	UPDATE transit_cost SET
		emp_company = p_org_company,
		bonbu = p_org_bonbu,
		saupbu = p_org_saupbu,
		team = p_org_team,
		org_name = p_org_name,
		reside_place = p_emp_reside_place
	WHERE mg_ce_id = p_mg_ce_id
		AND (run_date >= p_from_date AND run_date <= p_to_date);

	#야특근
	UPDATE overtime SET
		emp_company = p_org_company,
		bonbu = p_org_bonbu,
		saupbu = p_org_saupbu,
		team = p_org_team,
		org_name = p_org_name,
		reside_place = p_emp_reside_place
	WHERE mg_ce_id = p_mg_ce_id
		AND (work_date >= p_from_date AND work_date <= p_to_date);

	#카드
	UPDATE card_slip SET
		emp_company = p_org_company,
		bonbu = p_org_bonbu,
		saupbu = p_org_saupbu,
		team = p_org_team,
		org_name = p_org_name,
		reside_place = p_emp_reside_place,
		reside_company = p_emp_reside_company
	WHERE emp_no = p_mg_ce_id
		AND (slip_date >= p_from_date AND slip_date <= p_to_date);
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_GENERAL_SELECT;
#일반 경비 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_GENERAL_COST_SEL;
CREATE PROCEDURE USP_ORG_END_GENERAL_COST_SEL(
	IN p_emp_no varchar(20),
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 전사 조직 조회 > 일반 경비 조회
'
proc_body :
BEGIN
	SELECT pay_yn, COUNT(slip_seq) AS c_cnt, SUM(cost) AS cost
	FROM general_cost
	WHERE slip_gubun = '비용'
		AND (tax_bill_yn = 'N' OR ISNULL(tax_bill_yn))
		AND cancel_yn = 'N'
		AND (slip_date >= p_from_date AND slip_date <= p_to_date)
		AND emp_no = p_emp_no
	GROUP BY pay_yn;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_OVERTIME_SELECT;
#야특근 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_OVERTIME_COST_SEL;
CREATE PROCEDURE USP_ORG_END_OVERTIME_COST_SEL(
	IN p_emp_no varchar(20),
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 전사 조직 조회 > 야특근 조회
'
proc_body :
BEGIN
	SELECT cancel_yn, COUNT(work_date) AS c_cnt, SUM(overtime_amt) AS cost
	FROM overtime
	WHERE mg_ce_id = p_emp_no
	AND (work_date >= p_from_date AND work_date <= p_to_date)
	AND cancel_yn = 'N'
	GROUP BY cancel_yn;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_TRANSIT_SELECT;
#교통비 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_TRANSIT_COST_SEL;
CREATE PROCEDURE USP_ORG_END_TRANSIT_COST_SEL(
	IN p_emp_no varchar(20),
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 전사 조직 조회 > 교통비 조회
'
proc_body :
BEGIN
	SELECT car_owner, fare, far, oil_kind, oil_price, repair_cost, parking, toll
	FROM transit_cost
	WHERE mg_ce_id = p_emp_no
		AND (run_date >= p_from_date AND run_date <= p_to_date)
		AND cancel_yn = 'N';
END;


#교통비 조회
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_TRANSIT_SELECT;
CREATE PROCEDURE USP_COST_END_PERSON_TRANSIT_SELECT(
	IN p_emp_no varchar(20),
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 전사 조직 조회 > 교통비 조회
'
proc_body :
BEGIN
	SELECT car_owner, fare, far, oil_kind, oil_price, repair_cost, parking, toll
	FROM transit_cost
	WHERE mg_ce_id = p_emp_no
		AND (run_date >= p_from_date AND run_date <= p_to_date)
		AND cancel_yn = 'N';
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_OIL_UNIT_SELECT;
#차량 연료 구분 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_OIL_COST_UNIT_SEL;
CREATE PROCEDURE USP_ORG_END_OIL_COST_UNIT_SEL(
	IN p_end_month varchar(6),
	IN p_oil_unit_id varchar(1)
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
- 사업부별 비용마감 > 전사 조직 조회 > 차량 연료 구분 조회
'
proc_body :
BEGIN
	SELECT oil_unit_average, oil_kind
	FROM oil_unit
	WHERE oil_unit_month = p_end_month
	AND oil_unit_id = p_oil_unit_id;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_CARD_OIL_SELECT;
#주유 카드 사용 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_CARD_OIL_SEL;
CREATE PROCEDURE USP_ORG_END_CARD_OIL_SEL(
	IN p_emp_no varchar(20),
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 전사 조직 조회 > 주유 카드 사용
'
proc_body :
BEGIN
	SELECT COUNT(*) AS c_cnt, SUM(cost) AS cost, SUM(cost_vat) AS cost_vat
	FROM card_slip
	WHERE emp_no = p_emp_no
		AND (slip_date >= p_from_date AND slip_date <= p_to_date)
		AND card_type LIKE '%주유%';
END;


/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_CARD_SELECT;
#카드 사용 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_CARD_SEL;
CREATE PROCEDURE USP_ORG_END_CARD_SEL(
	IN p_emp_no varchar(20),
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 전사 조직 조회 > 카드 사용 조회
'
proc_body :
BEGIN
	SELECT COUNT(*) AS c_cnt , SUM(cost) AS cost , SUM(cost_vat) AS cost_vat
	FROM card_slip
	WHERE emp_no = p_emp_no
		AND (slip_date >= p_from_date AND slip_date <= p_to_date)
		AND card_type NOT LIKE '%주유%';
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_CAR_SELECT;
#차량 소유 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_CAR_SEL;
CREATE PROCEDURE USP_ORG_END_CAR_SEL(
	IN p_emp_no varchar(20))
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-05
DESC :
- 사업부별 비용마감 > 전사 조직 조회 > 차량 소유 조회
'
proc_body :
BEGIN
	SELECT car_owner
	FROM car_info
	WHERE owner_emp_no = p_emp_no;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_MEMO_SELECT;
#비용 비고 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_PERSON_MEMO_SEL;
CREATE PROCEDURE USP_ORG_END_PERSON_MEMO_SEL(
	IN p_emp_no varchar(20),
	IN p_emp_month varchar(6)
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
- 사업부별 비용마감 > 전사 조직 조회 > 비용 비고
'
proc_body :
BEGIN
	SELECT variation_memo
	FROM person_cost
	WHERE cost_month = p_emp_month
		AND emp_no = p_emp_no;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_COST_DEL;
#개인 비용 정보 초기화
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_PERSON_COST_DEL;
CREATE PROCEDURE USP_ORG_END_PERSON_COST_DEL(
	IN p_emp_no varchar(6),
	IN p_end_month varchar(6)
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
- 사업부별 비용마감 > 개인 비용 정보 초기화
'
proc_body :
BEGIN
	DELETE FROM person_cost
	WHERE cost_month = p_end_month
		AND emp_no = p_emp_no;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PERSON_COST_IN;
#개인 비용 입력
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_PERSON_COST_IN;
CREATE PROCEDURE USP_ORG_END_PERSON_COST_IN(
	IN p_end_month varchar(6),
	IN p_emp_no varchar(6),
	IN p_emp_name varchar(20) CHARSET utf8,
	IN p_emp_job varchar(20) CHARSET utf8,
	IN p_emp_end varchar(10) CHARSET utf8,
	IN p_car_owner varchar(10) CHARSET utf8,
	IN p_org_company varchar(30) CHARSET utf8,
	IN p_org_bonbu varchar(30) CHARSET utf8,
	IN p_org_saupbu varchar(30) CHARSET utf8,
	IN p_org_team varchar(30) CHARSET utf8,
	IN p_org_name varchar(30) CHARSET utf8,
	IN p_emp_reside_place varchar(30) CHARSET utf8,
	IN p_emp_reside_company varchar(50) CHARSET utf8,
	IN p_general_cnt int(3),
	IN p_general_cost int(11),
	IN p_general_pre_cnt int(3),
	IN p_general_pre_cost int(11),
	IN p_overtime_cnt int(3),
	IN p_overtime_cost int(11),
	IN p_gas_km int(11),
	IN p_gas_unit int(11),
	IN p_gas_cost int(11),
	IN p_diesel_km int(11),
	IN p_diesel_unit int(11),
	IN p_diesel_cost int(11),
	IN p_gasol_km int(11),
	IN p_gasol_unit int(11),
	IN p_gasol_cost int(11),
	IN p_tot_km int(11),
	IN p_tot_cost int(11),
	IN p_somopum_cost int(11),
	IN p_fare_cnt int(11),
	IN p_fare_cost int(11),
	IN p_oil_cash_cost int(11),
	IN p_repair_cost int(11),
	IN p_repair_pre_cost int(11),
	IN p_parking_cost int(11),
	IN p_toll_cost int(11),
	IN p_juyoo_card_cnt int(11),
	IN p_juyoo_card_cost int(11),
	IN p_juyoo_card_cost_vat int(11),
	IN p_card_cnt int(3),
	IN p_card_cost int(11),
	IN p_card_cost_vat int(11),
	IN p_return_cash int(11),
	IN p_variation_memo LONGTEXT CHARSET utf8
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
- 사업부별 비용마감 > 개인 비용 입력
'
proc_body :
BEGIN
	INSERT INTO person_cost
	VALUES(
		p_end_month, p_emp_no, p_emp_name,
		p_emp_job, p_emp_end, p_car_owner,
		p_org_company, p_org_bonbu, p_org_saupbu,
		p_org_team, p_org_name, p_emp_reside_place,
		p_emp_reside_company, p_general_cnt, p_general_cost,
		p_general_pre_cnt, p_general_pre_cost, p_overtime_cnt,
		p_overtime_cost, p_gas_km, p_gas_unit,
		p_gas_cost, p_diesel_km, p_diesel_unit,
		p_diesel_cost, p_gasol_km, p_gasol_unit,
		p_gasol_cost, p_tot_km, p_tot_cost,
		p_somopum_cost, p_fare_cnt, p_fare_cost,
		p_oil_cash_cost, p_repair_cost, p_repair_pre_cost,
		p_parking_cost, p_toll_cost, p_juyoo_card_cnt,
		p_juyoo_card_cost, p_juyoo_card_cost_vat, p_card_cnt,
		p_card_cost, p_card_cost_vat, p_return_cash,
		p_variation_memo, NOW(), 0
	);
END;


/*
DROP PROCEDURE IF EXISTS USP_COST_END_INSURE_SELECT;
#4대보험율
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_INSURE_SEL;
CREATE PROCEDURE USP_ORG_END_INSURE_SEL(
	IN p_insure_year varchar(4)
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
- 사업부별 비용마감 > 4대보험율
'
proc_body :
BEGIN
	SELECT insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per
	FROM insure_per
	WHERE insure_year = p_insure_year;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_ORG_COST_RESET;
#조직 비용 마감 초기화
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_COST_RESET_UP;
CREATE PROCEDURE USP_ORG_END_COST_RESET_UP(
	IN p_cost_year varchar(4),
	IN p_cost_month varchar(2),
	IN p_bonbu varchar(30) CHARSET utf8
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-07
DESC :
- 비용마감 > 조직 비용 마감 초기화
'
proc_body :
BEGIN
	SET @cost_year = p_cost_year;
	SET @cost_month = p_cost_month;
	SET @bonbu = p_bonbu;

	IF @cost_month = '01' THEN
		SET @v_col = CONCAT("cost_amt_01");
	ELSEIF @cost_month = '02' THEN
		SET @v_col = CONCAT("cost_amt_02");
	ELSEIF @cost_month = '03' THEN
		SET @v_col = CONCAT("cost_amt_03");
	ELSEIF @cost_month = '04' THEN
		SET @v_col = CONCAT("cost_amt_04");
	ELSEIF @cost_month = '05' THEN
		SET @v_col = CONCAT("cost_amt_05");
	ELSEIF @cost_month = '06' THEN
		SET @v_col = CONCAT("cost_amt_06");
	ELSEIF @cost_month = '07' THEN
		SET @v_col = CONCAT("cost_amt_07");
	ELSEIF @cost_month = '08' THEN
		SET @v_col = CONCAT("cost_amt_08");
	ELSEIF @cost_month = '09' THEN
		SET @v_col = CONCAT("cost_amt_09");
	ELSEIF @cost_month = '10' THEN
		SET @v_col = CONCAT("cost_amt_10");
	ELSEIF @cost_month = '11' THEN
		SET @v_col = CONCAT("cost_amt_11");
	ELSEIF @cost_month = '12' THEN
		SET @v_col = CONCAT("cost_amt_12");
	END IF;

	SET @v_sql = CONCAT("UPDATE org_cost SET ", @v_col, " = 0 ");
	SET @v_sql = CONCAT(@v_sql, "WHERE cost_year = ? AND bonbu = ?;");

	PREPARE stmt FROM @v_sql;
	EXECUTE stmt USING @cost_year, @bonbu;
	DEALLOCATE PREPARE stmt;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_PAY_SELECT;
#급여 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_PAY_SEL;
CREATE PROCEDURE USP_ORG_END_PAY_SEL(
	IN p_end_month varchar(6),
	IN p_dept_name varchar(30) CHARSET utf8
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
- 사업부별 비용마감 > 급여 조회
'
proc_body :
BEGIN
	SELECT eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team,
		eomt.org_name, pmgt.pmg_id,
		SUM(pmgt.pmg_give_total) AS tot_cost, SUM(pmgt.pmg_base_pay) AS base_pay,
		SUM(pmgt.pmg_meals_pay) AS meals_pay, SUM(pmgt.pmg_overtime_pay) AS overtime_pay,
		SUM(pmgt.pmg_research_pay) AS research_pay, SUM(pmgt.pmg_tax_no) AS tax_no
	FROM pay_month_give AS pmgt
	INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no
		AND emmt.emp_month = p_end_month
	INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code
		AND eomt.org_month = p_end_month
	WHERE eomt.org_bonbu = p_dept_name
		AND pmgt.pmg_yymm = p_end_month
		AND pmgt.pmg_id = '1'
	GROUP BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name;
END;


/*
DROP PROCEDURE IF EXISTS USP_COST_END_ORG_COST_INIT;
#조직별 비용 등록/수정
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_COST_ID_IN_UP;
CREATE PROCEDURE USP_ORG_END_COST_ID_IN_UP(
	IN p_cost_year varchar(4),
	IN p_emp_company varchar(30) CHARSET utf8,
	IN p_bonbu varchar(30) CHARSET utf8,
	IN p_saupbu varchar(29) CHARSET utf8,
	IN p_team varchar(30) CHARSET utf8,
	IN p_org_name varchar(30) CHARSET utf8,
	IN p_cost_id varchar(30) CHARSET utf8,
	IN p_cost_detail varchar(30) CHARSET utf8,
	IN p_total_cost bigint(20),
	IN p_sort_seq int(2),
	IN p_cost_month varchar(4)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-07
DESC :
- 비용마감 > 조직별 비용 등록/수정
'
proc_body :
BEGIN
	SET @cost_year = p_cost_year;
	SET @emp_company = p_emp_company;
	SET @bonbu = p_bonbu;
	SET @saupbu = p_saupbu;
	SET @team = p_team;
	SET @org_name = p_org_name;
	SET @cost_id = p_cost_id;
	SET @cost_detail = p_cost_detail;
	SET @total_cost = p_total_cost;
	SET @sort_seq = p_sort_seq;
	SET @cost_month = p_cost_month;

	SELECT COUNT(*) INTO @cost_cnt
	FROM org_cost
	WHERE cost_year = @cost_year
		AND emp_company = @emp_company
		AND bonbu = @bonbu
		AND saupbu = @saupbu
		AND team = @team
		AND org_name = @org_name
		AND cost_id = @cost_id
		AND cost_detail = @cost_detail;

	IF @cost_month = '01' THEN
		SET @v_col = CONCAT("cost_amt_01");
	ELSEIF @cost_month = '02' THEN
		SET @v_col = CONCAT("cost_amt_02");
	ELSEIF @cost_month = '03' THEN
		SET @v_col = CONCAT("cost_amt_03");
	ELSEIF @cost_month = '04' THEN
		SET @v_col = CONCAT("cost_amt_04");
	ELSEIF @cost_month = '05' THEN
		SET @v_col = CONCAT("cost_amt_05");
	ELSEIF @cost_month = '06' THEN
		SET @v_col = CONCAT("cost_amt_06");
	ELSEIF @cost_month = '07' THEN
		SET @v_col = CONCAT("cost_amt_07");
	ELSEIF @cost_month = '08' THEN
		SET @v_col = CONCAT("cost_amt_08");
	ELSEIF @cost_month = '09' THEN
		SET @v_col = CONCAT("cost_amt_09");
	ELSEIF @cost_month = '10' THEN
		SET @v_col = CONCAT("cost_amt_10");
	ELSEIF @cost_month = '11' THEN
		SET @v_col = CONCAT("cost_amt_11");
	ELSEIF @cost_month = '12' THEN
		SET @v_col = CONCAT("cost_amt_12");
	END IF;

	IF @cost_cnt > 0 THEN
		SET @v_sql = concat("UPDATE org_cost SET ", @v_col, " = ?, ");
		SET @v_sql = concat(@v_sql, "sort_seq = ? ");
		SET @v_sql = concat(@v_sql, "WHERE cost_year = ? ");
		SET @v_sql = concat(@v_sql, "AND emp_company = ? ");
		SET @v_sql = concat(@v_sql, "AND bonbu = ? ");
		SET @v_sql = concat(@v_sql, "AND saupbu = ? ");
		SET @v_sql = concat(@v_sql, "AND team = ? ");
		SET @v_sql = concat(@v_sql, "AND org_name = ? ");
		SET @v_sql = concat(@v_sql, "AND cost_id = ? ");
		SET @v_sql = concat(@v_sql, "AND cost_detail = ? ");

		PREPARE stmt FROM @v_sql;
		EXECUTE stmt USING @total_cost, @sort_seq, @cost_year, @emp_company, @bonbu, @saupbu,
			@team, @org_name, @cost_id, @cost_detail;
		DEALLOCATE PREPARE stmt;
	ELSE
		SET @v_sql = CONCAT("INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team,");
		SET @v_sql = CONCAT(@v_sql, "org_name, cost_id, cost_detail, ", @v_col, ", sort_seq)");
		SET @v_sql = CONCAT(@v_sql, "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?);");

		PREPARE stmt FROM @v_sql;
		EXECUTE stmt USING @cost_year, @emp_company, @bonbu, @saupbu, @team, @org_name, @cost_id,
			@cost_detail, @total_cost, @sort_seq;
		DEALLOCATE PREPARE stmt;
	END IF;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_BONUS_SELECT;
#상여 비용 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_BONUS_SEL;
CREATE PROCEDURE USP_ORG_END_BONUS_SEL(
	IN p_end_month varchar(6),
	IN p_dept_name varchar(30) CHARSET utf8
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
- 사업부별 비용마감 > 상여 비용 조회
'
proc_body :
BEGIN
	SELECT eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_name,
		pmgt.pmg_id,
		SUM(pmgt.pmg_give_total) AS cost
	FROM pay_month_give AS pmgt
	INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no
		AND emmt.emp_month = p_end_month
	INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code
		AND eomt.org_month = p_end_month
	WHERE eomt.org_bonbu = p_dept_name
		AND pmgt.pmg_yymm = p_end_month
		AND pmgt.pmg_id = '2'
	GROUP BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_ALBA_SELECT;
#알바 비용 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_ALBA_SEL;
CREATE PROCEDURE USP_ORG_END_ALBA_SEL(
	IN p_end_month varchar(6),
	IN p_dept_name varchar(30) CHARSET utf8
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
- 사업부별 비용마감 > 알바 비용 조회
'
proc_body :
BEGIN
	SELECT company, bonbu, saupbu, team, org_name, SUM(alba_give_total) AS cost
	FROM pay_alba_cost
	WHERE bonbu = p_dept_name
		AND rever_yymm = p_end_month
	GROUP BY company, bonbu, saupbu, team, org_name;
END;

/*
#야특근 마감 처리
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_OVERTIME_END_UP;
CREATE PROCEDURE USP_ORG_END_OVERTIME_END_UP(
	IN p_from_date varchar(10),
	IN p_to_date varchar(10),
	IN p_dept_name varchar(30) CHARSET utf8
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
- 사업부별 비용마감 > 야특근 마감 처리
'
proc_body :
BEGIN
	UPDATE overtime SET
		end_yn = 'Y'
	WHERE (work_date >= p_from_date AND work_date <= p_to_date)
		AND bonbu = p_dept_name	;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_GENERAL_SELECT;
#일반 경비 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_GENERAL_SEL;
CREATE PROCEDURE USP_ORG_END_GENERAL_SEL(
	IN p_end_month varchar(6),
	IN p_dept_name varchar(30) CHARSET utf8,
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 일반 경비 조회
'
proc_body :
BEGIN
	SELECT glct.slip_date, glct.slip_seq
	FROM general_cost AS glct
	INNER JOIN emp_master_month AS emmt ON glct.emp_no = emmt.emp_no
		AND emmt.emp_month = p_end_month
	INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code
	WHERE (glct.slip_date >= p_from_date AND glct.slip_date <= p_to_date)
		AND eomt.org_bonbu = p_dept_name;
END;

/*
#일반경비 마감 처리
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_GENERAL_END_UP;
CREATE PROCEDURE USP_ORG_END_GENERAL_END_UP(
	IN p_slip_date varchar(10),
	IN p_slip_seq varchar(3)
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
- 사업부별 비용마감 > 일반경비 마감 처리
'
proc_body :
BEGIN
	UPDATE general_cost SET
		end_yn = 'Y'
	WHERE slip_date = p_slip_date
		AND slip_seq = p_slip_seq;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_GENERAL_COST_SELECT;
#일반 경비 비용 처리 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_GENERAL_ORG_COST_SEL;
CREATE PROCEDURE USP_ORG_END_GENERAL_ORG_COST_SEL(
	IN p_dept_name varchar(30) CHARSET utf8,
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 일반 경비 비용 처리
'
proc_body :
BEGIN
	SELECT emp_company, bonbu, saupbu, team, org_name, account, SUM(cost) AS cost
	FROM general_cost
	WHERE slip_gubun = '비용' AND cancel_yn = 'N'
		AND (slip_date >= p_from_date AND slip_date <= p_to_date)
		AND bonbu = p_dept_name
	GROUP BY emp_company, bonbu, saupbu, team, org_name, account;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_GENERAL_ETC_SELECT;
#일반 경비 비용 외 처리 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_GENERAL_ETC_SEL;
CREATE PROCEDURE USP_ORG_END_GENERAL_ETC_SEL(
	IN p_dept_name varchar(30) CHARSET utf8,
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 일반 경비  비용 외 처리
'
proc_body :
BEGIN
	SELECT emp_company, bonbu, saupbu, team, org_name, slip_gubun, account, SUM(cost) AS cost
	FROM general_cost
	WHERE slip_gubun <> '비용' AND cancel_yn = 'N'
		AND (slip_date >= p_from_date AND slip_date <= p_to_date)
		AND bonbu = p_dept_name
	GROUP BY emp_company, bonbu, saupbu, team, org_name, account;
END;

/*
#교통비 마감 처리
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_TRANSIT_END_UP;
CREATE PROCEDURE USP_ORG_END_TRANSIT_END_UP(
	IN p_from_date varchar(10),
	IN p_to_date varchar(10),
	IN p_dept_name varchar(30) CHARSET utf8
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
- 사업부별 비용마감 > #교통비 마감 처리
'
proc_body :
BEGIN
	UPDATE transit_cost SET
		end_yn = 'Y'
	WHERE (run_date >= p_from_date AND run_date <= p_to_date)
		AND bonbu = p_dept_name	;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_TRANSIT_SELECT;
#교통비 마감 처리 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_TRANSIT_SEL;
CREATE PROCEDURE USP_ORG_END_TRANSIT_SEL(
	IN p_dept_name varchar(30) CHARSET utf8,
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 교통비 마감 처리
'
proc_body :
BEGIN
	SELECT emp_company, bonbu, saupbu, team, org_name, car_owner,
		SUM(somopum + oil_price + fare + parking + toll) AS cost
	FROM transit_cost
	WHERE cancel_yn = 'N'
		AND (run_date >= p_from_date AND run_date <= p_to_date)
		AND bonbu = p_dept_name
	GROUP BY emp_company, bonbu, saupbu, team, org_name, car_owner;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_TRAN_REPAIR_SELECT;
#교통비 차량수리비 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_TRAN_REPAIR_SEL;
CREATE PROCEDURE USP_ORG_END_TRAN_REPAIR_SEL(
	IN p_dept_name varchar(30) CHARSET utf8,
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 교통비 차량수리비
'
proc_body :
BEGIN
	SELECT emp_company, bonbu, saupbu, team, org_name, SUM(repair_cost) AS cost
	FROM transit_cost
	WHERE cancel_yn = 'N'
		AND repair_cost > 0
		AND (run_date >= p_from_date AND run_date <= p_to_date)
		AND bonbu = p_dept_name
	GROUP BY emp_company, bonbu, team, org_name;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_TRAN_COMP_SELECT;
#회사 차량 운행 사번 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_TRAN_COMP_SEL;
CREATE PROCEDURE USP_ORG_END_TRAN_COMP_SEL(
	IN p_dept_name varchar(30) CHARSET utf8,
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 회사 차량 운행 사번 조회
'
proc_body :
BEGIN
	SELECT mg_ce_id
	FROM transit_cost
	WHERE (run_date >= p_from_date AND run_date <= p_to_date)
		AND car_owner = '회사'
		AND bonbu = p_dept_name
	GROUP BY mg_ce_id;
END;

/*
#회사 차량 운행  카드 비용 마감
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_TRAN_CARD_UP;
CREATE PROCEDURE USP_ORG_END_TRAN_CARD_UP(
	IN p_emp_no varchar(6),
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 회사 차량 운행  카드 비용 마감
'
proc_body :
BEGIN
	UPDATE card_slip SET
		com_drv_yn = 'Y'
	WHERE (slip_date >= p_from_date AND slip_date <= p_to_date)
		AND emp_no = p_emp_no;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_CARD_OIL_SELECT;
#회사 차량 운행 정보 조회
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_CARD_COST_SEL;
CREATE PROCEDURE USP_ORG_END_CARD_COST_SEL(
	IN p_dept_name varchar(30) CHARSET utf8,
	IN p_from_date varchar(10),
	IN p_to_date varchar(10)
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
- 사업부별 비용마감 > 회사 차량 운행 정보 조회
'
proc_body :
BEGIN
	SELECT owner_company as emp_company, bonbu, saupbu, team, org_name, account, SUM(cost) AS cost
	FROM card_slip
	WHERE (slip_date >= p_from_date AND slip_date <= p_to_date)
		AND (card_type NOT LIKE '%주유%' OR com_drv_yn = 'Y')
		AND bonbu = p_dept_name
	GROUP BY owner_company, bonbu, team, org_name, account;
END;


/*
#사업부 별 비용 마감 구분 처리
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_PROC;
CREATE PROCEDURE USP_ORG_END_PROC(
	IN p_end_month varchar(6),
	IN p_saupbu varchar(30) CHARSET utf8,
	IN p_end_yn varchar(1),
	IN p_reg_id varchar(20),
	IN p_reg_name varchar(20) CHARSET utf8
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
- 사업부별 비용마감 > 사업부 별 비용 마감 구분 처리
'
proc_body :
BEGIN
	SET @end_month = p_end_month;
	SET @saupbu = p_saupbu;
	SET @end_yn = p_end_yn;
	SET @reg_id = p_reg_id;
	SET @reg_name = p_reg_name;

	IF @end_yn = 'C' THEN
		UPDATE cost_end SET
			end_yn = 'Y',
			mod_id = @reg_id,
			mod_name = @reg_name,
			mod_date = NOW()
		WHERE end_month = @end_month
			AND saupbu = @saupbu;
	ELSE
		DELETE FROM cost_end
		WHERE end_month = @end_month
			AND saupbu = @saupbu;

		INSERT INTO cost_end (end_month, saupbu, end_yn, batch_yn, bonbu_yn, ceo_yn,
			reg_id, reg_name, reg_date)
		VALUES(@end_month, @saupbu, '', '', '', '', @reg_id, @reg_name, NOW());
	END IF;
END;

/*
DROP PROCEDURE IF EXISTS USP_COST_END_SALES_ORG_SEL;
# 영업 본부 조회(일괄 정산용)
*/
DROP PROCEDURE IF EXISTS USP_ORG_END_SALES_SEL;
CREATE PROCEDURE USP_ORG_END_SALES_SEL(
	#IN p_emp_month varchar(6)
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
- 비용마감 > 영업 본부 조회(일괄 정산용)
'
proc_body :
BEGIN
	SELECT org_name
	FROM emp_org_mst
	WHERE org_level = '본부'
		AND (ISNULL(org_end_date) OR org_end_date = '0000-00-00')
		AND org_name NOT IN ('전략부문', 'ICT연구소', '빅데이타연구소', '기술연구소', '한진그룹사업본부')
	GROUP BY org_bonbu, org_name
	ORDER BY FIELD(org_company, '케이원', '케이네트웍스', '케이시스템'),
		FIELD(org_bonbu, '스마트본부', '공공SI본부', '금융SI본부', 'ICT본부', '공공본부',
			'NI본부', 'SI2본부', 'SI1본부') DESC;
END;