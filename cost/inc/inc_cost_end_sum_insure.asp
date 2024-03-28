<%
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
objBuilder.Append "	AND bonbu = '"&deptName&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'급여 조회 및 정산
objBuilder.Append "SELECT eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "	eomt.org_name, pmgt.pmg_id, "
objBuilder.Append "	SUM(pmgt.pmg_give_total) AS tot_cost, SUM(pmgt.pmg_base_pay) AS base_pay, "
objBuilder.Append "	SUM(pmgt.pmg_meals_pay) AS meals_pay, SUM(pmgt.pmg_overtime_pay) AS overtime_pay, "
objBuilder.Append "	SUM(pmgt.pmg_research_pay) AS research_pay, SUM(pmgt.pmg_tax_no) AS tax_no "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&end_month&"' "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month ='"&end_month&"' "
objBuilder.Append "WHERE eomt.org_bonbu = '"&deptName&"' "
objBuilder.Append "	AND pmgt.pmg_yymm ='"&end_month&"' "
objBuilder.Append "	AND pmgt.pmg_id ='1' "
objBuilder.Append "GROUP BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name "

Set rsPay = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsPay.EOF Then
	arrPay = rsPay.getRows()
End If
rsPay.Close() : Set rsPay = Nothing

If IsArray(arrPay) Then
	For i = LBound(arrPay) To UBound(arrPay, 2)
		org_company = arrPay(0, i)
		org_bonbu = arrPay(1, i)
		org_saupbu = arrPay(2, i)
		org_team = arrPay(3, i)
		org_name = arrPay(4, i)
		pmg_id = arrPay(5, i)
		tot_cost = arrPay(6, i)
		base_pay = arrPay(7, i)
		meals_pay = arrPay(8, i)
		overtime_pay = arrPay(9, i)
		research_pay = arrPay(10, i)
		tax_no = arrPay(11, i)

		sort_seq = 0
		cost_detail = "급여"

		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "FROM org_cost "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company ='"&org_company&"' "
		objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
		objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
		objBuilder.Append "	AND team ='"&org_team&"' "
		objBuilder.Append "	AND org_name ='"&org_name&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='"&cost_detail&"' "

		Set rs_payCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_payCost.EOF Or rs_payCost.BOF Then
			objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
			objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
			objBuilder.Append "VALUES("
			objBuilder.Append "'"&cost_year&"',"
			objBuilder.Append "'"&org_company&"', "
			objBuilder.Append "'"&org_bonbu&"', "
			objBuilder.Append "'"&org_saupbu&"', "
			objBuilder.Append "'"&org_team&"', "
			objBuilder.Append "'"&org_name&"', "
			objBuilder.Append "'인건비', "
			objBuilder.Append "'"&cost_detail&"', "
			objBuilder.Append tot_cost&", "
			objBuilder.Append sort_seq&") "
		Else
			objBuilder.Append "UPDATE org_cost SET "
			objBuilder.Append "cost_amt_"&cost_month&"="&tot_cost&", "
			objBuilder.Append "sort_seq="&sort_seq&" "
			objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND emp_company = '"&org_company&"' "
			objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
			objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
			objBuilder.Append "	AND team ='"&org_team&"' "
			objBuilder.Append "	AND org_name ='"&org_name&"' "
			objBuilder.Append "	AND cost_id ='인건비' "
			objBuilder.Append "	AND cost_detail ='"&cost_detail&"' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rs_payCost.Close() : Set rs_payCost = Nothing

		'4대보험료
		insure_tot = CLng((CLng(tot_cost)) * insure_tot_per / 100)
		sort_seq = 2

		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "FROM org_cost "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company ='"&org_company&"' "
		objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
		objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
		objBuilder.Append "	AND team ='"&org_team&"' "
		objBuilder.Append "	AND org_name ='"&org_name&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='4대보험' "

		Set rs_insureCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_insureCost.EOF Or rs_insureCost.BOF Then
			objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
			objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
			objBuilder.Append "VALUES("
			objBuilder.Append "'"&cost_year&"',"
			objBuilder.Append "'"&org_company&"', "
			objBuilder.Append "'"&org_bonbu&"', "
			objBuilder.Append "'"&org_saupbu&"', "
			objBuilder.Append "'"&org_team&"', "
			objBuilder.Append "'"&org_name&"', "
			objBuilder.Append "'인건비', "
			objBuilder.Append "'4대보험', "
			objBuilder.Append insure_tot&", "
			objBuilder.Append sort_seq&") "
		Else
			objBuilder.Append "UPDATE org_cost SET "
			objBuilder.Append "cost_amt_"&cost_month&"="&insure_tot&", "
			objBuilder.Append "sort_seq="&sort_seq&" "
			objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND emp_company = '"&org_company&"' "
			objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
			objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
			objBuilder.Append "	AND team ='"&org_team&"' "
			objBuilder.Append "	AND org_name ='"&org_name&"' "
			objBuilder.Append "	AND cost_id ='인건비' "
			objBuilder.Append "	AND cost_detail ='4대보험' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rs_insureCost.Close() : Set rs_insureCost = Nothing

		' 소득세 종업원분
		income_tax = clng((clng(tot_cost)) * income_tax_per / 100)
		sort_seq = 3

		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "FROM org_cost "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company ='"&org_company&"' "
		objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
		objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
		objBuilder.Append "	AND team ='"&org_team&"' "
		objBuilder.Append "	AND org_name ='"&org_name&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='소득세종업원분' "

		Set rs_incomeCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_incomeCost.EOF Or rs_incomeCost.BOF Then
			objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
			objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
			objBuilder.Append "VALUES("
			objBuilder.Append "'"&cost_year&"',"
			objBuilder.Append "'"&org_company&"', "
			objBuilder.Append "'"&org_bonbu&"', "
			objBuilder.Append "'"&org_saupbu&"', "
			objBuilder.Append "'"&org_team&"', "
			objBuilder.Append "'"&org_name&"', "
			objBuilder.Append "'인건비', "
			objBuilder.Append "'소득세종업원분', "
			objBuilder.Append income_tax&", "
			objBuilder.Append sort_seq&") "
		Else
			objBuilder.Append "UPDATE org_cost SET "
			objBuilder.Append "cost_amt_"&cost_month&"="&income_tax&", "
			objBuilder.Append "sort_seq="&sort_seq&" "
			objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND emp_company = '"&org_company&"' "
			objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
			objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
			objBuilder.Append "	AND team ='"&org_team&"' "
			objBuilder.Append "	AND org_name ='"&org_name&"' "
			objBuilder.Append "	AND cost_id ='인건비' "
			objBuilder.Append "	AND cost_detail ='소득세종업원분' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rs_incomeCost.Close() : Set rs_incomeCost = Nothing

		'연차수당
		annual_pay = CLng((CLng(base_pay) + CLng(meals_pay) + CLng(overtime_pay)) * annual_pay_per / 100)
		sort_seq = 4

		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "FROM org_cost "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company ='"&org_company&"' "
		objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
		objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
		objBuilder.Append "	AND team ='"&org_team&"' "
		objBuilder.Append "	AND org_name ='"&org_name&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='연차수당' "

		Set rs_annualCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_annualCost.EOF Or rs_annualCost.BOF Then
			objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
			objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
			objBuilder.Append "VALUES("
			objBuilder.Append "'"&cost_year&"',"
			objBuilder.Append "'"&org_company&"', "
			objBuilder.Append "'"&org_bonbu&"', "
			objBuilder.Append "'"&org_saupbu&"', "
			objBuilder.Append "'"&org_team&"', "
			objBuilder.Append "'"&org_name&"', "
			objBuilder.Append "'인건비', "
			objBuilder.Append "'연차수당', "
			objBuilder.Append annual_pay&", "
			objBuilder.Append sort_seq&") "
		Else
			objBuilder.Append "UPDATE org_cost SET "
			objBuilder.Append "cost_amt_"&cost_month&"="&annual_pay&", "
			objBuilder.Append "sort_seq="&sort_seq&" "
			objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND emp_company = '"&org_company&"' "
			objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
			objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
			objBuilder.Append "	AND team ='"&org_team&"' "
			objBuilder.Append "	AND org_name ='"&org_name&"' "
			objBuilder.Append "	AND cost_id ='인건비' "
			objBuilder.Append "	AND cost_detail ='연차수당' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rs_annualCost.Close() : Set rs_annualCost = Nothing

		' 퇴직충당금
		retire_pay = CLng((CLng(base_pay) + CLng(meals_pay) + CLng(overtime_pay)) * retire_pay_per / 100)
		sort_seq = 5

		objBuilder.Append "SELECT cost_year "
		objBuilder.Append "FROM org_cost "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company ='"&org_company&"' "
		objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
		objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
		objBuilder.Append "	AND team ='"&org_team&"' "
		objBuilder.Append "	AND org_name ='"&org_name&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='퇴직충당금' "

		Set rs_retireCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_retireCost.EOF Or rs_retireCost.BOF Then
			objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
			objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
			objBuilder.Append "VALUES("
			objBuilder.Append "'"&cost_year&"',"
			objBuilder.Append "'"&org_company&"', "
			objBuilder.Append "'"&org_bonbu&"', "
			objBuilder.Append "'"&org_saupbu&"', "
			objBuilder.Append "'"&org_team&"', "
			objBuilder.Append "'"&org_name&"', "
			objBuilder.Append "'인건비', "
			objBuilder.Append "'퇴직충당금', "
			objBuilder.Append retire_pay&", "
			objBuilder.Append sort_seq&") "
		Else
			objBuilder.Append "UPDATE org_cost SET "
			objBuilder.Append "cost_amt_"&cost_month&"="&retire_pay&", "
			objBuilder.Append "sort_seq="&sort_seq&" "
			objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND emp_company = '"&org_company&"' "
			objBuilder.Append "	AND bonbu ='"&org_bonbu&"' "
			objBuilder.Append "	AND saupbu ='"&org_saupbu&"' "
			objBuilder.Append "	AND team ='"&org_team&"' "
			objBuilder.Append "	AND org_name ='"&org_name&"' "
			objBuilder.Append "	AND cost_id ='인건비' "
			objBuilder.Append "	AND cost_detail ='퇴직충당금' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rs_retireCost.Close() : Set rs_retireCost = Nothing
	Next
End If
%>