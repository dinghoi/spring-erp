<%
' 사업부별 인원수 집계
'sql = " select saupbu from sales_org where sales_year='" & cost_year & "' order by saupbu asc"
objBuilder.Append "SELECT saupbu "
objBuilder.Append "FROM sales_org "
objBuilder.Append "WHERE sales_year='" & cost_year & "' "
objBuilder.Append "ORDER BY saupbu ASC "

'이전 레코드셋 : rs
Set rsSalesDept = Server.CreateObject("ADODB.RecordSet")
rsSalesDept.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

i = 0
tot_person = 0

Do Until rsSalesDept.EOF
	'공통비 배부기준 변경 처리(2016-01-15)
	'sql = "select count(*) from pay_month_give  A ,emp_master_month B "
	'sql = sql & "where A.pmg_id = '1'  "
	'sql = sql & "and A.pmg_yymm = '"&end_month&"' "
	'sql = sql & "and A.mg_saupbu ='"&rs("saupbu")&"' "
	'sql = sql & "and A.pmg_emp_no=  B.emp_no "
	'sql = sql & "and B.cost_except in('0','1') "
	'sql = sql & "and B.emp_month ='"&end_month&"' "

	objBuilder.Append "SELECT COUNT(*) "
	objBuilder.Append "FROM pay_month_give AS pmgt "
	objBuilder.Append "LEFT OUTER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
	objBuilder.Append "	AND emmt.emp_month = '"&end_month&"' "
	objBuilder.Append "LEFT OUTER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
	objBuilder.Append "	AND eomt.org_month = '"&end_month&"' "
	objBuilder.Append "WHERE pmgt.pmg_id = '1' "
	objBuilder.Append "	AND pmgt.pmg_yymm = '"&end_month&"' "
	objBuilder.Append "	AND emmt.cost_except IN ('0', '1') "
	objBuilder.Append "	AND pmgt.mg_saupbu = '"&rsSalesDept("saupbu")&"' "
	objBuilder.Append "	AND eomt.org_name <> '육군본부2팀' "	'육군본부2팀 제외
	objBuilder.Append "	AND emmt.emp_type = '정직' "	'정직원만 포함

	'이전 레코드셋명 : rs_emp
	Set rsPayCnt = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'급여 기준 총 인원 수
	If rsPayCnt(0) = "" Or IsNull(rsPayCnt(0)) Then
		saupbu_person = 0
	Else
		saupbu_person = CLng(rsPayCnt(0))
	End If
	rsPayCnt.Close()

	i = i + 1
	saupbu_tab(i, 1) = rsSalesDept("saupbu")
	saupbu_tab(i, 2) = saupbu_person
	tot_person = tot_person + saupbu_person

	rsSalesDept.MoveNext()
Loop
rsSalesDept.Close() : Set rsSalesDept = Nothing

'전사공통비 합계
'sql = "select sum(cost_amt_"&mm&") as tot_cost from company_cost where cost_year ='"&cost_year&"' and cost_center = '전사공통비'"
objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS tot_cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_center = '전사공통비' "

'이전 레코드셋명 : rs
Set rsCompanyCostTot = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'전사 공통비 총액
tot_cost_amt = CLng(rsCompanyCostTot("tot_cost"))

rsCompanyCostTot.Close() : Set rsCompanyCostTot = Nothing

Dim rsSalesTot, salesTotal

'전체 매출 총액(기타사업부 제외)
objBuilder.Append "SELECT SUM(cost_amt) AS tot_sales "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE REPLACE(substring(sales_date, 1, 7), '-', '') = '"&end_month&"' "
objBuilder.Append "	AND saupbu <> '기타사업부' "

Set rsSalesTot = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'전체 매출 총액
salesTotal = CDbl(rsSalesTot("tot_sales"))

rsSalesTot.Close() : Set rsSalesTot = Nothing 

' 처리전 zero
'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='전사공통비') "
objBuilder.Append "UPDATE saupbu_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&" = '0' "
objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
objBuilder.Append "	AND cost_center ='전사공통비' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "delete from management_cost where cost_month ='"&end_month&"'"
objBuilder.Append "DELETE FROM management_cost "
objBuilder.Append "WHERE cost_month ='"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Set rsSalesCompCost = Server.CreateObject("ADODB.RecordSet")
Set rsCompanyCost = Server.CreateObject("ADODB.RecordSet")

' 전사공통비 배부
For i = 1 To 10
	If saupbu_tab(i, 1) = "" Or IsNull(saupbu_tab(i, 1)) Then
		Exit For
	End If

	If saupbu_tab(i, 1) <> "기타사업부" Then

		' 사업부별 매출 총액
		'sql = "select sum(cost_amt) from saupbu_sales where substring(sales_date,1,7) = '"&sales_month&"' and saupbu ='"&saupbu_tab(i,1)&"'"
		objBuilder.Append "SELECT SUM(sast.cost_amt) "
		objBuilder.Append "FROM saupbu_sales AS sast "
		objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&sales_month&"' "
		objBuilder.Append "	AND sast.saupbu = '"&saupbu_tab(i, 1)&"' "

		'rs_cost
		Set rsSalesCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsSalesCost(0) = "" Or IsNull(rsSalesCost(0)) Then
			saupbu_sales = 0
		Else
			saupbu_sales = CCur(rsSalesCost(0))	'사업부별 매출 총액
		End If
		rsSalesCost.Close()

		saupbu_per = saupbu_tab(i,2) / tot_person	'사업부별 인원 비율
		
		'saupbu_cost_amt = Int(tot_cost_amt * saupbu_per)
		'전사공통비(매출 50%) = (전사공통비 * 0.5) / 매출총액 * 사업부 매출 총액
		saupbu_cost_amt = Int((tot_cost_amt * 0.5) / salesTotal * saupbu_sales)

		'사업부별 고객사별 매출 총액
		'sql = "select company,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,7) = '"&sales_month&"' and saupbu ='"&saupbu_tab(i,1)&"' group by saupbu,company"
		objBuilder.Append "SELECT sast.company, SUM(sast.cost_amt) AS cost "
		objBuilder.Append "FROM saupbu_sales AS sast "
		objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&sales_month&"' "
		objBuilder.Append "	AND sast.saupbu = '"&saupbu_tab(i,1)&"' "
		objBuilder.Append "GROUP BY sast.saupbu, sast.company "

		'rs_etc
		rsSalesCompCost.Open objBuilder.ToString(), DBConn, 1
		objBuilder.Clear()

		k = 0

		Do Until rsSalesCompCost.EOF
			k = k + 1
			If saupbu_sales = 0 Then
				charge_per = 0
			Else
				charge_per = rsSalesCompCost("cost") / saupbu_sales
			End If

			cost_amt = Int(charge_per * saupbu_cost_amt)

			'sql = "insert into management_cost (cost_month,saupbu,company,tot_person,saupbu_person,saupbu_per,tot_cost_amt,saupbu_cost_amt,charge_per,cost_amt,reg_id,reg_name,reg_date) values ('"&end_month&"','"&saupbu_tab(i,1)&"','"&rs_etc("company")&"',"&tot_person&","&saupbu_tab(i,2)&","&saupbu_per&","&tot_cost_amt&","&saupbu_cost_amt&","&charge_per&","&cost_amt&",'"&user_Id&"','"&user_name&"',now())"
			objBuilder.Append "INSERT INTO management_cost(cost_month, saupbu, company, "
			objBuilder.Append "tot_person, saupbu_person, saupbu_per, "
			objBuilder.Append "tot_cost_amt, saupbu_cost_amt, charge_per, "
			objBuilder.Append "cost_amt, reg_id, reg_name, reg_date)VALUES("
			objBuilder.Append "'"&end_month&"', '"&saupbu_tab(i,1)&"', '"&rsSalesCompCost("company")&"', "
			objBuilder.Append ""&tot_person&", "&saupbu_tab(i,2)&", "&saupbu_per&", "
			objBuilder.Append ""&tot_cost_amt&", "&saupbu_cost_amt&", "&charge_per&", "
			objBuilder.Append ""&cost_amt&", '"&user_Id&"', '"&user_name&"', NOW()) "

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			rsSalesCompCost.MoveNext()
		Loop
		rsSalesCompCost.Close()

		' 매출이 제로인 경우
		If k = 0 Then
			'sql = "insert into management_cost (cost_month,saupbu,company,tot_person,saupbu_person,saupbu_per,tot_cost_amt,saupbu_cost_amt,charge_per,cost_amt,reg_id,reg_name,reg_date) values ('"&end_month&"','"&saupbu_tab(i,1)&"','',"&tot_person&","&saupbu_tab(i,2)&","&saupbu_per&","&tot_cost_amt&","&saupbu_cost_amt&",1,"&saupbu_cost_amt&",'"&user_Id&"','"&user_name&"',now())"
			objBuilder.Append "INSERT INTO management_cost(cost_month,saupbu,company,"
			objBuilder.Append "tot_person,saupbu_person,saupbu_per, "
			objBuilder.Append "tot_cost_amt,saupbu_cost_amt,charge_per, "
			objBuilder.Append "cost_amt,reg_id,reg_name,reg_date)VALUES("
			objBuilder.Append "'"&end_month&"','"&saupbu_tab(i,1)&"','', "
			objBuilder.Append ""&tot_person&","&saupbu_tab(i,2)&","&saupbu_per&", "
			objBuilder.Append ""&tot_cost_amt&","&saupbu_cost_amt&",1, "
			objBuilder.Append ""&saupbu_cost_amt&",'"&user_Id&"','"&user_name&"',NOW()) "

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		End If

		'sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '전사공통비' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
		objBuilder.Append "SELECT cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
		objBuilder.Append "FROM company_cost "
		objBuilder.Append "WHERE cost_center = '전사공통비' "
		objBuilder.Append "	AND cost_year ='"&cost_year&"' "
		objBuilder.Append "GROUP BY cost_id, cost_detail "

		'rs_etc
		rsCompanyCost.Open objBuilder.ToString(), DBConn, 1
		objBuilder.Clear()

		Do Until rsCompanyCost.EOF
			cost = Int(saupbu_per * CLng(rsCompanyCost("cost")))

			'sql = "select * from saupbu_profit_loss where cost_year ='"&cost_year&"' and saupbu ='"&saupbu_tab(i,1)&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			objBuilder.Append "SELECT cost_year "
			objBuilder.Append "FROM saupbu_profit_loss "
			objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND saupbu ='"&saupbu_tab(i,1)&"' "
			objBuilder.Append "	AND cost_center ='전사공통비' "
			objBuilder.Append "	AND cost_id ='"&rsCompanyCost("cost_id")&"' "
			objBuilder.Append "	AND cost_detail ='"&rsCompanyCost("cost_detail")&"' "

			'rs_cost
			Set rsCompanyCommCost = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsCompanyCommCost.EOF Or rsCompanyCommCost.BOF Then
				'sql = "insert into saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&saupbu_tab(i,1)&"','전사공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
				objBuilder.Append "INSERT INTO saupbu_profit_loss(cost_year, saupbu, cost_center, "
				objBuilder.Append "cost_id, cost_detail, cost_amt_"&cost_month&")VALUES("
				objBuilder.Append "'"&cost_year&"', '"&saupbu_tab(i,1)&"', '전사공통비', "
				objBuilder.Append "'"&rsCompanyCost("cost_id")&"', '"&rsCompanyCost("cost_detail")&"', "&cost&") "
			Else
				'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and saupbu ='"&saupbu_tab(i,1)&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
				objBuilder.Append "UPDATE saupbu_profit_loss SET "
				objBuilder.Append "	cost_amt_"&cost_month&"="&cost&" "
				objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
				objBuilder.Append "	AND saupbu ='"&saupbu_tab(i,1)&"' "
				objBuilder.Append "	AND cost_center ='전사공통비' AND cost_id ='"&rsCompanyCost("cost_id")&"'  "
				objBuilder.Append "	AND cost_detail ='"&rsCompanyCost("cost_detail")&"' "
			End If
			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			rsCompanyCost.MoveNext()
		Loop
		rsCompanyCost.Close()
	End If
Next
' 전사공통비 배부 끝

' 고객사별 손익 자료 생성
' 전사공통비 배부
' 처리전 zero
'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='전사공통비') "
objBuilder.Append "UPDATE company_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&" = '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_center ='전사공통비' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = " select company,saupbu_per, sum(charge_per) as charge_per from management_cost Where (cost_month = '"&end_month&"') GROUP BY company"
objBuilder.Append "SELECT company, saupbu_per, SUM(charge_per) AS charge_per "
objBuilder.Append "FROM management_cost "
objBuilder.Append "WHERE cost_month = '"&end_month&"' "
objBuilder.Append "GROUP BY company "

'rs
Set rsMgCost = Server.CreateObject("ADODB.RecordSet")
rsMgCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Set rsMgCompCost = Server.CreateObject("ADODB.RecordSet")

Do Until rsMgCost.EOF
	charge_per = rsMgCost("charge_per")

	'sql = "select * from trade where trade_name = '"&rs("company")&"'"
	objBuilder.Append "SELECT group_name "
	objBuilder.Append "FROM trade "
	objBuilder.Append "WHERE trade_name = '"&rsMgCost("company")&"' "

	'rs_trade
	Set rsMgCostTrade = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsMgCostTrade.EOF Or rsMgCostTrade.BOF Then
		group_name = "Error"
	Else
		group_name = rsMgCostTrade("group_name")
	End If
	rsMgCostTrade.Close()

	'sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '전사공통비' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	objBuilder.Append "SELECT cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_center = '전사공통비' "
	objBuilder.Append "	AND cost_year ='"&cost_year&"' "
	objBuilder.Append "GROUP BY cost_id, cost_detail "

	'rs_etc
	rsMgCompCost.Open objBuilder.ToString(), DBConn, 1
	objBuilder.Clear()

	Do Until rsMgCompCost.EOF
		cost = Int(charge_per * CLng(rsMgCompCost("cost")) * rsMgCost("saupbu_per"))

		'sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "	FROM company_profit_loss "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND company ='"&rsMgCost("company")&"' "
		objBuilder.Append "	AND group_name ='"&group_name&"' "
		objBuilder.Append "	AND cost_center ='전사공통비' "
		objBuilder.Append "	AND cost_id ='"&rsMgCompCost("cost_id")&"' "
		objBuilder.Append "	AND cost_detail ='"&rsMgCompCost("cost_detail")&"' "

		'rs_cost
		Set rsMgProfit = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsMgProfit.EOF Or rsMgProfit.BOF Then
			'sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','전사공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			objBuilder.Append "INSERT INTO company_profit_loss(cost_year, company, group_name,"
			objBuilder.Append "cost_center, cost_id, cost_detail, cost_amt_"&cost_month&")VALUES("
			objBuilder.Append "'"&cost_year&"', '"&rsMgCost("company")&"', '"&group_name&"', "
			objBuilder.Append "'전사공통비', '"&rsMgCompCost("cost_id")&"', '"&rsMgCompCost("cost_detail")&"', "&cost&")"
		Else
			'sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			objBuilder.Append "UPDATE company_profit_loss SET "
			objBuilder.Append "	cost_amt_"&cost_month&"="&cost&" "
			objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
			objBuilder.Append "	AND company = '"&rsMgCost("company")&"' "
			objBuilder.Append "	AND group_name = '"&group_name&"' "
			objBuilder.Append "	AND cost_center = '전사공통비' "
			objBuilder.Append "	AND cost_id = '"&rsMgCompCost("cost_id")&"' "
			objBuilder.Append "	AND cost_detail = '"&rsMgCompCost("cost_detail")&"' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rsMgProfit.Close()

		rsMgCompCost.MoveNext()
	Loop
	rsMgCompCost.Close()

	rsMgCost.MoveNext()
Loop
rsMgCost.Close() : Set rsMgCost = Nothing

' 고객사별 직접비 배부
' 처리전 zero
'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='직접비') "
objBuilder.Append "UPDATE company_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&" = '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_center ='직접비' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = " select saupbu,company, sum(charge_per) as charge_per from management_cost Where (cost_month = '"&end_month&"') GROUP BY saupbu,company"
objBuilder.Append "SELECT saupbu, company, SUM(charge_per) AS charge_per "
objBuilder.Append "FROM management_cost "
objBuilder.Append "WHERE cost_month = '"&end_month&"' "
objBuilder.Append "GROUP BY saupbu, company "

'rs
Set rsMgDeptCost = Server.CreateObject("ADODB.RecordSet")
rsMgDeptCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Set rsMgDeptCompany = Server.CreateObject("ADODB.RecordSet")

Do Until rsMgDeptCost.EOF
	charge_per = rsMgDeptCost("charge_per")

	'sql = "select * from trade where trade_name = '"&rs("company")&"'"
	objBuilder.Append "SELECT group_name "
	objBuilder.Append "FROM trade "
	objBuilder.Append "WHERE trade_name = '"&rsMgDeptCost("company")&"' "

	'rs_trade
	Set rsMgDeptTrade = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsMgDeptTrade.eof Or rsMgDeptTrade.bof Then
		group_name = "Error"
	Else
		group_name = rsMgDeptTrade("group_name")
	End If
	rsMgDeptTrade.Close()

	'sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '직접비' ) and (saupbu = '"&rs("saupbu")&"' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	objBuilder.Append "SELECT cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_center = '직접비' "
	objBuilder.Append "	AND saupbu = '"&rsMgDeptCost("saupbu")&"' "
	objBuilder.Append "	AND cost_year ='"&cost_year&"' "
	objBuilder.Append "GROUP BY cost_id, cost_detail "

	'rs_etc
	rsMgDeptCompany.Open objBuilder.ToString(), DBConn, 1
	objBuilder.Clear()

	Do Until rsMgDeptCompany.EOF
		cost = Int(charge_per * CDbl(rsMgDeptCompany("cost")))

		'sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='직접비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "FROM company_profit_loss "
		objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
		objBuilder.Append "	AND company = '"&rsMgDeptCost("company")&"' "
		objBuilder.Append "	AND group_name = '"&group_name&"' "
		objBuilder.Append "	AND cost_center = '직접비' "
		objBuilder.Append "	AND cost_id = '"&rsMgDeptCompany("cost_id")&"' "
		objBuilder.Append "	AND cost_detail = '"&rsMgDeptCompany("cost_detail")&"' "

		'rs_cost
		Set rsMgDeptProfitList = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsMgDeptProfitList.EOF Or rsMgDeptProfitList.BOF Then
			'sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','직접비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			objBuilder.Append "INSERT INTO company_profit_loss(cost_year, company, group_name, "
			objBuilder.Append "cost_center, cost_id, cost_detail, cost_amt_"&cost_month&")VALUES("
			objBuilder.Append "'"&cost_year&"', '"&rsMgDeptCost("company")&"', '"&group_name&"', "
			objBuilder.Append "'직접비', '"&rsMgDeptCompany("cost_id")&"', '"&rsMgDeptCompany("cost_detail")&"',"&cost&") "
		Else
			'sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='직접비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			objBuilder.Append "UPDATE company_profit_loss SET "
			objBuilder.Append "	cost_amt_"&cost_month&"="&cost&" "
			objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
			objBuilder.Append "	AND company = '"&rsMgDeptCost("company")&"' "
			objBuilder.Append "	AND group_name = '"&group_name&"' "
			objBuilder.Append "	AND cost_center = '직접비' "
			objBuilder.Append "	AND cost_id = '"&rsMgDeptCompany("cost_id")&"' "
			objBuilder.Append "	AND cost_detail = '"&rsMgDeptCompany("cost_detail")&"' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rsMgDeptProfitList.Close()

		rsMgDeptCompany.MoveNext()
	Loop
	rsMgDeptCompany.Close()

	rsMgDeptCost.MoveNext()
Loop
rsMgDeptCost.Close() : Set rsMgDeptCost = Nothing
' 고객사별 직접비 배부 끝
%>