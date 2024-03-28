<%
'거래처 별 손익 자료 생성
Dim manage_type

Dim rsCowork, as_give_cowork, as_get_cowork, cowork_give_cost, cowork_get_cost

Dim exceptDate

'202204월부터 전사공통비 SI1본부 고객사 삼성생명보험(주) 매출 제외 처리(재무 요청)[허정호_20220511]
exceptDate = "202204"

'거래처 손익 초기화
objBuilder.Append "DELETE FROM company_cost_profit "
objBuilder.Append "WHERE cost_month = '"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'영업 사업부 조회
objBuilder.Append "SELECT saupbu FROM sales_org "
objBuilder.Append "WHERE sales_year = '"&cost_year&"' "
objBuilder.Append "ORDER BY sort_seq ASC "

Set rsSalesOrg = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsSalesOrg.EOF Then
	arrSalesOrg = rsSalesOrg.getRows()
End If
rsSalesOrg.Close() : Set rsSalesOrg = Nothing

If IsArray(arrSalesOrg) Then
	For i = LBound(arrSalesOrg) To UBound(arrSalesOrg, 2)
		saupbu = arrSalesOrg(0, i)

		'사업부별 매출 조회
		objBuilder.Append "SELECT SUM(cost_amt) AS 'sales_total' "
		objBuilder.Append "FROM saupbu_sales "
		objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&cost_date&"' "
		objBuilder.Append "	AND saupbu = '"&saupbu&"'; "

		Set rsSalesTot = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		sales_total = CDbl(f_toString(rsSalesTot(0), 0))	'사업부 별 총 매출

		rsSalesTot.Close() : Set rsSalesTot = Nothing

		'상주직접비 Total(공통 제외)
		objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS 'company_total' "
		objBuilder.Append "FROM company_cost "
		objBuilder.Append "WHERE cost_year = '"&cost_year&"' AND cost_center = '상주직접비' "

		If saupbu <> "기타사업부" Then
			objBuilder.Append "	AND (company <> '' AND company IS NOT NULL AND company <> '공통') "
			objBuilder.Append " AND saupbu = '"&saupbu&"'; "
		Else
			objBuilder.Append "	AND saupbu = '' "
		End If

		Set rsCompanyTot = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		company_tot = CDbl(rsCompanyTot(0))	'사업부 별 상주직접비(공통 제외)

		rsCompanyTot.Close() : Set rsCompanyTot = Nothing

		'공통경비(직접비 + 상주직접비(공통))
		objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS 'comm_cost', "

		'If mm = "06" Or mm = "12" Then
		''	objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") - (SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비' AND cost_id = '일반경비' AND cost_detail = '급여' AND saupbu = '"&saupbu&"') FROM company_cost  "
		'Else
		''	objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") FROM company_cost  "
		'End If
		'objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비'  "

		If saupbu = "기타사업부" Then
			objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") FROM company_cost  "
			objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비'  "
			objBuilder.Append "		AND (saupbu = '' OR saupbu = '"&saupbu&"')) AS 'direct_cost' "
		Else
			'objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") - (SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비' AND cost_id = '일반경비' AND cost_detail = '급여' AND saupbu = '"&saupbu&"') FROM company_cost  "
			'objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비'  "
			If mm = "06" Or mm = "12" Then
				objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") - (SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비' AND cost_id = '일반경비' AND cost_detail = '급여' AND saupbu = '"&saupbu&"') FROM company_cost  "
			Else
				objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") FROM company_cost  "
			End If
			objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비'  "

			objBuilder.Append "		AND saupbu = '"&saupbu&"') AS 'direct_cost' "
		End If

		objBuilder.Append "FROM company_cost "
		objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
		objBuilder.Append "	AND (company = '' OR company is null OR company = '공통') "
		objBuilder.Append "	AND cost_center = '상주직접비' "
		objBuilder.Append "	AND saupbu = '"&saupbu&"' "

		'If saupbu = "스마트본부" then
		'dbconn.rollbacktrans
		'Response.write objBuilder.toString()
		'Response.end
		'end if
		Set rsComm = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		comm_cost = CDbl(f_toString(rsComm("comm_cost"), 0))	'상주직접비(공통)
		direct_cost = CDbl(f_toString(rsComm("direct_cost"), 0))	'직접비

		'공통경비 = 상주직접비(공통) + 직접비(인건비+경비)
		common_total = comm_cost + direct_cost

		rsComm.Close() : Set rsComm = Nothing

		'전사공통비
		objBuilder.Append "SELECT ROUND((tot_cost_amt * 0.5 / tot_person * saupbu_person) "
		objBuilder.Append "	+ (tot_cost_amt * 0.5 / tot_sale * saupbu_sale), 1) AS tot_amt "
		objBuilder.Append "FROM ( "
		objBuilder.Append "	SELECT mgct.saupbu, mgct.tot_cost_amt, mgct.saupbu_person, mgct.tot_person, "

		objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
		objBuilder.Append "		FROM saupbu_sales "
		objBuilder.Append "		WHERE SUBSTRING(sales_date, 1, 7) = '"&cost_date&"' "
		objBuilder.Append "			AND mgct.saupbu = saupbu "

		If Replace(cost_date, "-", "") >= exceptDate Then
			objBuilder.Append "		AND company <> '삼성생명보험(주)' "
		End If

		objBuilder.Append "		) AS saupbu_sale, "

		objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
		objBuilder.Append "		FROM saupbu_sales "
		objBuilder.Append "		WHERE SUBSTRING(sales_date, 1, 7) = '"&cost_date&"' "
		objBuilder.Append "			AND saupbu <> '기타사업부' "

		If Replace(cost_date, "-", "") >= exceptDate Then
			objBuilder.Append "		AND company <> '삼성생명보험(주)' "
		End If

		objBuilder.Append "		) AS tot_sale "

		objBuilder.Append "	FROM management_cost AS mgct "
		objBuilder.Append "	WHERE cost_month = '"&end_month&"' "
		objBuilder.Append "		AND saupbu = '"&saupbu&"' "
		objBuilder.Append "	GROUP BY saupbu "
		objBuilder.Append ") r1 "

		Set rsManage = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsManage.EOF Or rsManage.BOF Then
			manage_tot = 0
		Else
			manage_tot = CDbl(f_toString(rsManage(0), 0))	'부서별 전사공통비
		End If
		rsManage.Close() : Set rsManage = Nothing

		'if saupbu = "스마트본부" then
		''	dbconn.rollbacktrans
		''	response.write manage_tot
		''	response.end
		'end if

		'부문공통비(배분)
		objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
		objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND cost_detail = '설치공사')) AS 'part_tot_cost', "
		objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&cost_year&mm&"') AS 'as_tot_cnt' "
		objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '부문공통비' "

		Set rsPart = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		part_tot_cost = CDbl(f_toString(rsPart("part_tot_cost"), 0))	'부문공통비(배분)
		as_tot_cnt = CInt(f_toString(rsPart("as_tot_cnt"), 0))	'AS 총 건수

		rsPart.Close() : Set rsPart = Nothing

		'사업부 별 AS 총 건수 조회
		objBuilder.Append "SELECT SUM(as_total - as_set) AS as_cnt "
		objBuilder.Append "FROM as_acpt_status AS aast "
		objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
		objBuilder.Append "	AND trdt.trade_id = '매출' "
		objBuilder.Append "WHERE as_month = '"&cost_year&mm&"' "
		If saupbu = "기타사업부" Then
			objBuilder.Append "AND trdt.saupbu = '' "
		Else
			objBuilder.Append "	AND trdt.saupbu = '"&saupbu&"' "
		End If

		Set rsSaupbuPart = DBConn.Execute(objBuilder.ToString())

		part_cnt = CInt(f_toString(rsSaupbuPart(0), 0))	'사업부 AS 총 건수

		objBuilder.Clear()
		rsSaupbuPart.Close() : Set rsSaupbuPart = Nothing

		'사업부별 배분 부분공통비
		If part_cnt > 0 Then
			part_tot = part_tot_cost / as_tot_cnt * part_cnt
		Else
			part_tot = 0
		End If

		'거래처별 비용 현황
		objBuilder.Append "CALL USP_SALES_COMPANY_PROFIT_SEL('"&saupbu&"', '"&cost_year&"', '"&MID(from_date, 1, 7)&"', '"&mm&"');"

		Set rsCompCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If Not rsCompCost.EOF Then
			arrCompCost = rsCompCost.getRows()
		End If
		rsCompCost.Close() : Set rsCompCost = Nothing

		'사이트별 전사공통비 안분 기준
		' 1. 상주인원 인건비
		' 2. 매출액
		' 3. 거래처별 비용계
		If manage_tot > 0 Then
			'상주직접비 기준 안분
			If company_tot > 0 Then
				'manage_cost = manage_tot * company_cost / company_tot
				manage_type = "company"
			Else
				'매출 기준 안분
				If sales_total > 0 Then
					'manage_cost =  manage_tot * sales_cost / sales_total	'사이트별 전사공통비(매출 기준)
					manage_type = "sales"
				Else
					'거래처 별 비용 합계 안분
					'manage_cost = manage_tot * (company_cost + common_cost) / (common_total + company_tot)
					manage_type = "common"
				End If
			End If
		Else
			manage_type = "none"
		End If

		If IsArray(arrCompCost) Then
			'사이트 별 분기 처리
			For j = LBound(arrCompCost) To UBound(arrCompCost, 2)
				company = arrCompCost(0, j)	'거래처명
				sales_cost = CDbl(arrCompCost(1, j))	'거래처별 매출
				company_cost = CDbl(arrCompCost(2, j))	'상주직접비(인건비+일반경비)
				pay_cost = CDbl(arrCompCost(3, j))	'상주직접비(인건비)
				general_cost = CDbl(arrCompCost(4, j))	'상주직접비(일반경비)
				as_cnt = CInt(arrCompCost(5, j))	'사이트별 AS 건수

				'사업부공통경비 = 거래처별 상주직접비 / 사업부 별 상주직접비(공통 제외) * 공통경비
				If company_tot > 0 Or company_tot < 0 Then
					common_cost = company_cost / company_tot * common_total
				Else
					common_cost = 0
				End If

				If as_cnt > 0 Then
					part_cost = part_tot / part_cnt * as_cnt	'사이트별 부분공통비(AS건수 기준)
				Else
					part_cost = 0
				End If

				'사이트별 전사공통비 배부
				Select Case manage_type
					Case "company"
						'상주직접비 기준 안분
						manage_cost = manage_tot * company_cost / company_tot
					Case "sales"
						'매출 기준 안분
						manage_cost =  manage_tot * sales_cost / sales_total	'사이트별 전사공통비(매출 기준)
					Case "common"
						'거래처 별 비용 합계 안분
						manage_cost = manage_tot * (company_cost + common_cost) / (common_total + company_tot)
					Case Else
						'manage_cost = 0
						manage_cost = manage_tot * company_cost / company_tot
				End Select

				'협업 건수
				objBuilder.Append "SELECT aast.as_give_cowork, aast.as_get_cowork FROM as_acpt_status AS aast "
				objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
				objBuilder.Append "WHERE aast.as_month = '"&end_month&"' AND aast.as_company = '"&company& "' "
				If saupbu = "기타사업부" Then
					objBuilder.Append "	AND trdt.saupbu = '' "
				Else
					objBuilder.Append "	AND trdt.saupbu = '"&saupbu&"' "
				End If

				Set rsCowork = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				If rsCowork.EOF Or rsCowork.BOF Then
					as_give_cowork = 0
					as_get_cowork = 0
				Else
					as_give_cowork = CDbl(rsCowork("as_give_cowork"))
					as_get_cowork = CDbl(rsCowork("as_get_cowork"))
				End If

				rsCowork.Close() : Set rsCowork = Nothing

				'협업 지원 비용(상주직접비 * 협업지원건수 / 사이트별 총건수)
				'cowork_give_cost = company_cost * as_give_cowork / as_cnt * -1
				cowork_give_cost = 30000 * as_give_cowork * -1

				'협업 받은 비용(상주직접비 * 협업받은건수 / 사이트별 총건수)
				'cowork_get_cost = company_cost * as_get_cowork / as_cnt
				cowork_get_cost = 30000 * as_get_cowork

				'pay_cost = pay_cost + cowork_give_cost + cowork_get_cost

				'손익 비용
				profit_cost = sales_cost - (pay_cost + general_cost + common_cost + part_cost + manage_cost)
				'profit_cost = sales_cost - (pay_cost + general_cost + common_cost + part_cost + manage_cost + cowork_give_cost + cowork_get_cost)

				objBuilder.Append "INSERT INTO company_cost_profit(cost_month, company_name, saupbu, sales_cost, pay_cost, "
				objBuilder.Append "general_cost, common_cost, part_cost, manage_cost, profit_cost, "
				objBuilder.Append "reg_date, reg_id, cowork_give_cost, cowork_get_cost)VALUES("
				objBuilder.Append "'"&end_month&"', '"&company&"', '"&saupbu&"', '"&sales_cost&"', '"&pay_cost&"', "
				objBuilder.Append "'"&general_cost&"', '"&common_cost&"', '"&part_cost&"', '"&manage_cost&"', '"&profit_cost&"', "
				objBuilder.Append "NOW(), '"&emp_no&"', '"&cowork_give_cost&"', '"&cowork_get_cost&"');"

				DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()
			Next

		End If
	Next
End If


%>