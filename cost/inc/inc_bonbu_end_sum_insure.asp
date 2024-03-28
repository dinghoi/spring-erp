<%
Dim rsPay, rs_insure, rs_payCost, rs_insureCost, rs_incomeCost, rs_annualCost, rs_retireCost
Dim sort_seq, cost_detail
Dim insure_tot, income_tax, annual_pay, retire_pay
Dim cost_id
Dim insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per

'4대보험율과 기타 인건비율 검색
objBuilder.Append "SELECT insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per "
objBuilder.Append "FROM insure_per "
objBuilder.Append "WHERE insure_year = '"&cost_year&"' "

Set rs_insure = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

insure_tot_per = rs_insure("insure_tot_per")
income_tax_per = rs_insure("income_tax_per")
annual_pay_per = rs_insure("annual_pay_per")
retire_pay_per = rs_insure("retire_pay_per")

rs_insure.Close() : Set rs_insure = Nothing

'조직 비용 마감 초기화
objBuilder.Append "UPDATE org_cost SET "
objBuilder.Append "	cost_amt_"&cost_month&"= '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "AND bonbu = '' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'급여 조회 및 정산
objBuilder.Append "SELECT eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_name, "
objBuilder.Append "	pmgt.pmg_id, pmgt.pmg_emp_no, pmgt.pmg_emp_name, "
objBuilder.Append "	SUM(pmgt.pmg_give_total) AS tot_cost, SUM(pmgt.pmg_base_pay) AS base_pay, "
objBuilder.Append "	SUM(pmgt.pmg_meals_pay) AS meals_pay, SUM(pmgt.pmg_overtime_pay) AS overtime_pay, "
objBuilder.Append "	SUM(pmgt.pmg_research_pay) AS research_pay, SUM(pmgt.pmg_tax_no) AS tax_no "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&end_month&"' "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month = '"&end_month&"' "
objBuilder.Append "WHERE pmgt.pmg_yymm ='"&end_month&"' "
objBuilder.Append "	AND eomt.org_bonbu = '' "
objBuilder.Append "	AND pmgt.pmg_id ='1' "
objBuilder.Append "GROUP BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name "

Set rsPay = Server.CreateObject("ADODB.RecordSet")
rsPay.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsPay.EOF
	sort_seq = 0
	cost_detail = "급여"

	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsPay("org_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
	objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
	objBuilder.Append "	AND cost_id ='인건비' "
	objBuilder.Append "	AND cost_detail ='"&cost_detail&"' "

	Set rs_payCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_payCost.EOF Or rs_payCost.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsPay("org_company")&"', "
		objBuilder.Append "'"&rsPay("org_bonbu")&"', "
		objBuilder.Append "'"&rsPay("org_saupbu")&"', "
		objBuilder.Append "'"&rsPay("org_team")&"', "
		objBuilder.Append "'"&rsPay("org_name")&"', "
		objBuilder.Append "'인건비', "
		objBuilder.Append "'"&cost_detail&"', "
		objBuilder.Append rsPay("tot_cost")&", "
		objBuilder.Append sort_seq&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsPay("tot_cost")&", "
		objBuilder.Append "	sort_seq="&sort_seq&" "
		objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsPay("org_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
		objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='"&cost_detail&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_payCost.Close()

	'4대보험료
	insure_tot = CLng((CLng(rsPay("tot_cost"))) * insure_tot_per / 100)
	sort_seq = 2

	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsPay("org_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
	objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
	objBuilder.Append "	AND cost_id ='인건비' "
	objBuilder.Append "	AND cost_detail ='4대보험' "

	Set rs_insureCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_insureCost.EOF Or rs_insureCost.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsPay("org_company")&"', "
		objBuilder.Append "'"&rsPay("org_bonbu")&"', "
		objBuilder.Append "'"&rsPay("org_saupbu")&"', "
		objBuilder.Append "'"&rsPay("org_team")&"', "
		objBuilder.Append "'"&rsPay("org_name")&"', "
		objBuilder.Append "'인건비', "
		objBuilder.Append "'4대보험', "
		objBuilder.Append insure_tot&", "
		objBuilder.Append sort_seq&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "cost_amt_"&cost_month&"="&rsPay("tot_cost")&", "
		objBuilder.Append "sort_seq="&sort_seq&" "
		objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsPay("org_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
		objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='4대보험' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_insureCost.Close()

	' 소득세 종업원분
	income_tax = clng((clng(rsPay("tot_cost"))) * income_tax_per / 100)
	sort_seq = 3

	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsPay("org_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
	objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
	objBuilder.Append "	AND cost_id ='인건비' "
	objBuilder.Append "	AND cost_detail ='소득세종업원분' "

	Set rs_incomeCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_incomeCost.EOF Or rs_incomeCost.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsPay("org_company")&"', "
		objBuilder.Append "'"&rsPay("org_bonbu")&"', "
		objBuilder.Append "'"&rsPay("org_saupbu")&"', "
		objBuilder.Append "'"&rsPay("org_team")&"', "
		objBuilder.Append "'"&rsPay("org_name")&"', "
		objBuilder.Append "'인건비', "
		objBuilder.Append "'소득세종업원분', "
		objBuilder.Append income_tax&", "
		objBuilder.Append sort_seq&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "cost_amt_"&cost_month&"="&rsPay("tot_cost")&", "
		objBuilder.Append "sort_seq="&sort_seq&" "
		objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsPay("org_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
		objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='소득세종업원분' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_incomeCost.Close()

	'연차수당
	annual_pay = CLng((CLng(rsPay("base_pay")) + CLng(rsPay("meals_pay")) + CLng(rsPay("overtime_pay"))) * annual_pay_per / 100)
	sort_seq = 4

	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsPay("org_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
	objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
	objBuilder.Append "	AND cost_id ='인건비' "
	objBuilder.Append "	AND cost_detail ='연차수당' "

	Set rs_annualCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_annualCost.EOF Or rs_annualCost.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsPay("org_company")&"', "
		objBuilder.Append "'"&rsPay("org_bonbu")&"', "
		objBuilder.Append "'"&rsPay("org_saupbu")&"', "
		objBuilder.Append "'"&rsPay("org_team")&"', "
		objBuilder.Append "'"&rsPay("org_name")&"', "
		objBuilder.Append "'인건비', "
		objBuilder.Append "'연차수당', "
		objBuilder.Append annual_pay&", "
		objBuilder.Append sort_seq&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "cost_amt_"&cost_month&"="&rsPay("tot_cost")&", "
		objBuilder.Append "sort_seq="&sort_seq&" "
		objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsPay("org_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
		objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='연차수당' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_annualCost.Close()

	' 퇴직충당금
	retire_pay = CLng((CLng(rsPay("base_pay")) + CLng(rsPay("meals_pay")) + CLng(rsPay("overtime_pay"))) * retire_pay_per / 100)
	sort_seq = 5

	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsPay("org_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
	objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
	objBuilder.Append "	AND cost_id ='인건비' "
	objBuilder.Append "	AND cost_detail ='퇴직충당금' "

	Set rs_retireCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_retireCost.EOF Or rs_retireCost.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsPay("org_company")&"', "
		objBuilder.Append "'"&rsPay("org_bonbu")&"', "
		objBuilder.Append "'"&rsPay("org_saupbu")&"', "
		objBuilder.Append "'"&rsPay("org_team")&"', "
		objBuilder.Append "'"&rsPay("org_name")&"', "
		objBuilder.Append "'인건비', "
		objBuilder.Append "'퇴직충당금', "
		objBuilder.Append retire_pay&", "
		objBuilder.Append sort_seq&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "cost_amt_"&cost_month&"="&rsPay("tot_cost")&", "
		objBuilder.Append "sort_seq="&sort_seq&" "
		objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsPay("org_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsPay("org_bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsPay("org_saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsPay("org_team")&"' "
		objBuilder.Append "	AND org_name ='"&rsPay("org_name")&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='퇴직충당금' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_retireCost.Close()

	rsPay.MoveNext()
Loop

Set rs_payCost = Nothing
Set rs_insureCost = Nothing
Set rs_incomeCost = Nothing
Set rs_annualCost = Nothing
Set rs_retireCost = Nothing
rsPay.Close() : Set rsPay = Nothing
%>