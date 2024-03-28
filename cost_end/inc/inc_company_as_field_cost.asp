<%
' 처리전 zero
'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='부문공통비') "
objBuilder.Append "UPDATE saupbu_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&" = '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_center ='부문공통비' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = " select saupbu, sum(charge_per) as charge_per from company_as Where (as_month = '"&end_month&"') GROUP BY saupbu"
objBuilder.Append "SELECT saupbu, SUM(charge_per) AS charge_per "
objBuilder.Append "FROM company_as "
objBuilder.Append "WHERE as_month = '"&end_month&"' "
objBuilder.Append "GROUP BY saupbu "

'이전 레코드셋 명 : rs
Set rsProfitDept = Server.CreateObject("ADODB.RecordSet")
rsProfitDept.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Set rsProfitDeptCost = Server.CreateObject("ADODB.RecordSet")

Do Until rsProfitDept.EOF
	charge_per = rsProfitDept("charge_per")

	'sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '부문공통비' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	objBuilder.Append "SELECT cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_center = '부문공통비' "
	objBuilder.Append "	AND cost_year ='"&cost_year&"' "
	objBuilder.Append "GROUP BY cost_id, cost_detail "

	'rs_etc
	rsProfitDeptCost.Open objBuilder.ToString(), DBConn, 1
	objBuilder.Clear()

	Do Until rsProfitDeptCost.EOF
		'cost
		profit_cost = Int(charge_per * CLng(rsProfitDeptCost("cost")))

		'sql = "select * from saupbu_profit_loss where cost_year ='"&cost_year&"' and saupbu ='"&rs("saupbu")&"' and cost_center ='부문공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "FROM saupbu_profit_loss "
		objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
		objBuilder.Append "	AND saupbu = '"&rsProfitDept("saupbu")&"' "
		objBuilder.Append "	AND cost_center = '부문공통비' "
		objBuilder.Append "	AND cost_id = '"&rsProfitDeptCost("cost_id")&"' "
		objBuilder.Append "	AND cost_detail = '"&rsProfitDeptCost("cost_detail")&"' "

		'rs_cost
		Set rsProfitDeptList = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsProfitDeptList.EOF Or rsProfitDeptList.BOF Then
			'sql = "insert into saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("saupbu")&"','부문공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			objBuilder.Append "INSERT INTO saupbu_profit_loss(cost_year, saupbu, cost_center,"
			objBuilder.Append "cost_id, cost_detail, cost_amt_"&cost_month&")VALUES("
			objBuilder.Append "'"&cost_year&"', '"&rsProfitDept("saupbu")&"', '부문공통비',"
			objBuilder.Append "'"&rsProfitDeptCost("cost_id")&"', '"&rsProfitDeptCost("cost_detail")&"', "&profit_cost&")"
		Else
			'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and saupbu ='"&rs("saupbu")&"' and cost_center ='부문공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			objBuilder.Append "UPDATE saupbu_profit_loss SET "
			objBuilder.Append "	cost_amt_"&cost_month&" = "&profit_cost&" "
			objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
			objBuilder.Append "	AND saupbu = '"&rsProfitDept("saupbu")&"' "
			objBuilder.Append "	AND cost_center = '부문공통비' "
			objBuilder.Append "	AND cost_id = '"&rsProfitDeptCost("cost_id")&"' "
			objBuilder.Append "	AND cost_detail = '"&rsProfitDeptCost("cost_detail")&"' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rsProfitDeptList.Close()

		rsProfitDeptCost.MoveNext()
	Loop
	rsProfitDeptCost.Close()

	rsProfitDept.MoveNext()
Loop
Set rsProfitDeptCost = Nothing
rsProfitDept.Close() : Set rsProfitDept = Nothing
' 부분공통비 배부 끝

' 고객사별 손익 자료 생성
' 부문공통비 배부
' 처리전 zero
'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='부문공통비') "
objBuilder.Append "UPDATE company_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&"= '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_center ='부문공통비' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = " select as_company as company, sum(charge_per) as charge_per from company_as Where (as_month = '"&end_month&"') GROUP BY as_company"
objBuilder.Append "SELECT as_company AS company, SUM(charge_per) AS charge_per "
objBuilder.Append "FROM company_as "
objBuilder.Append "WHERE as_month = '"&end_month&"' "
objBuilder.Append "GROUP BY as_company "

'이전 레코드셋 : rs
Set rsAsCompany = Server.CreateObject("ADODB.RecordSet")
rsAsCompany.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Set rsAsCompanyCost = Server.CreateObject("ADODB.RecordSet")

Do Until rsAsCompany.EOF
	charge_per = rsAsCompany("charge_per")

	'sql = "select * from trade where trade_name = '"&rs("company")&"'"
	objBuilder.Append "SELECT group_name "
	objBuilder.Append "FROM trade "
	objBuilder.Append "WHERE trade_name = '"&rsAsCompany("company")&"' "

	'rs_trade
	Set rsAsCompTrade = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsAsCompTrade.EOF Or rsAsCompTrade.BOF Then
		group_name = "Error"
	Else
		group_name = rsAsCompTrade("group_name")
	End If
	rsAsCompTrade.Close()

	'sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '부문공통비' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	objBuilder.Append "SELECT cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_center = '부문공통비' "
	objBuilder.Append "	AND cost_year ='"&cost_year&"' "
	objBuilder.Append "GROUP BY cost_id, cost_detail "

	'rs_etc
	rsAsCompanyCost.Open objBuilder.ToString(), DBConn, 1
	objBuilder.Clear()

	Do Until rsAsCompanyCost.EOF
		company_cost = Int(charge_per * CLng(rsAsCompanyCost("cost")))

		'sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='부문공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "FROM company_profit_loss "
		objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
		objBuilder.Append "	AND company = '"&rsAsCompany("company")&"' "
		objBuilder.Append "	AND group_name = '"&group_name&"' "
		objBuilder.Append "	AND cost_center = '부문공통비' "
		objBuilder.Append "	AND cost_id = '"&rsAsCompanyCost("cost_id")&"' "
		objBuilder.Append "	AND cost_detail = '"&rsAsCompanyCost("cost_detail")&"' "

		Set rsAsCompanyList = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsAsCompanyList.EOF or rsAsCompanyList.BOF Then
			'sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','부문공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			objBuilder.Append "INSERT INTO company_profit_loss(cost_year, company, group_name, "
			objBuilder.Append "cost_center, cost_id, cost_detail, "
			objBuilder.Append "cost_amt_"&cost_month&")VALUES("
			objBuilder.Append "'"&cost_year&"', '"&rsAsCompany("company")&"', '"&group_name&"', "
			objBuilder.Append "'부문공통비', '"&rsAsCompanyCost("cost_id")&"', '"&rsAsCompanyCost("cost_detail")&"',"
			objBuilder.Append company_cost&")"
		Else
			'sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='부문공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			objBuilder.Append "UPDATE company_profit_loss SET "
			objBuilder.Append "	cost_amt_"&cost_month&" = "&company_cost&" "
			objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
			objBuilder.Append "	AND company = '"&rsAsCompany("company")&"' "
			objBuilder.Append "	AND group_name = '"&group_name&"' "
			objBuilder.Append "	AND cost_center ='부문공통비' "
			objBuilder.Append "	AND cost_id = '"&rsAsCompanyCost("cost_id")&"' "
			objBuilder.Append "	AND cost_detail = '"&rsAsCompanyCost("cost_detail")&"' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rsAsCompanyList.Close()

		rsAsCompanyCost.MoveNext()
	loop
	rsAsCompanyCost.Close()

	rsAsCompany.MoveNext()
Loop
Set rsAsCompanyCost = Nothing
rsAsCompany.Close() : Set rsAsCompany = Nothing
' 부분공통비 배부 끝
%>