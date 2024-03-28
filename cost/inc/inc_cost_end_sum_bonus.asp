<%
'상여 SUM

objBuilder.Append "SELECT eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_name, "
objBuilder.Append "	eomt.org_name, pmgt.pmg_id, "
objBuilder.Append "	SUM(pmgt.pmg_give_total) AS cost "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&end_month&"' "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month ='"&end_month&"' "
objBuilder.Append "WHERE eomt.org_bonbu = '"&deptName&"' "
objBuilder.Append "	AND pmgt.pmg_yymm ='"&end_month&"' "
objBuilder.Append "	AND pmgt.pmg_id ='2' "
objBuilder.Append "GROUP BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name "

Set rsBunus = Server.CreateObject("ADODB.RecordSet")
rsBunus.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsBunus.EOF
	sort_seq = 1
	cost_detail = "상여"

	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsBunus("org_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsBunus("org_bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsBunus("org_saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsBunus("org_team")&"' "
	objBuilder.Append "	AND org_name ='"&rsBunus("org_name")&"' "
	objBuilder.Append "	AND cost_id ='인건비' "
	objBuilder.Append "	AND cost_detail ='"&cost_detail&"' "

	Set rs_bonus = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_bonus.EOF Or rs_bonus.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "	org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsBunus("org_company")&"', "
		objBuilder.Append "'"&rsBunus("org_bonbu")&"', "
		objBuilder.Append "'"&rsBunus("org_saupbu")&"', "
		objBuilder.Append "'"&rsBunus("org_team")&"', "
		objBuilder.Append "'"&rsBunus("org_name")&"', "
		objBuilder.Append "'인건비', "
		objBuilder.Append "'"&cost_detail&"', "
		objBuilder.Append rsBunus("cost")&", "
		objBuilder.Append sort_seq&");"
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "cost_amt_"&cost_month&"="&rsBunus("tot_cost")&", "
		objBuilder.Append "sort_seq="&sort_seq&" "
		objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsBunus("org_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsBunus("org_bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsBunus("org_saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsBunus("org_team")&"' "
		objBuilder.Append "	AND org_name ='"&rsBunus("org_name")&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='"&cost_detail&"'; "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_bonus.Close()

	rsBunus.MoveNext()
Loop
Set rs_bonus = Nothing
rsBunus.Close() : Set rsBunus = Nothing

'알바비
objBuilder.Append "SELECT company, bonbu, saupbu, team, org_name, SUM(alba_give_total) AS cost "
objBuilder.Append "	FROM pay_alba_cost "
objBuilder.Append "WHERE bonbu = '"&deptName&"' "
objBuilder.Append "	AND rever_yymm ='"&end_month&"' "
objBuilder.Append "GROUP BY company, bonbu, saupbu, team, org_name "

Set rsAlba = Server.CreateObject("ADODB.RecordSet")
rsAlba.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsAlba.EOF
	sort_seq = 8

	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsAlba("company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsAlba("bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsAlba("saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsAlba("team")&"' "
	objBuilder.Append "	AND org_name ='"&rsAlba("org_name")&"' "
	objBuilder.Append "	AND cost_id ='인건비' "
	objBuilder.Append "	AND cost_detail ='알바비' "

	Set rs_alba = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_alba.EOF Or rs_alba.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "	org_name, cost_id, cost_detail, cost_amt_"&cost_month&", sort_seq)"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsAlba("company")&"', "
		objBuilder.Append "'"&rsAlba("bonbu")&"', "
		objBuilder.Append "'"&rsAlba("saupbu")&"', "
		objBuilder.Append "'"&rsAlba("team")&"', "
		objBuilder.Append "'"&rsAlba("org_name")&"', "
		objBuilder.Append "'인건비', "
		objBuilder.Append "'알바비', "
		objBuilder.Append rsAlba("cost")&", "
		objBuilder.Append sort_seq&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&" = "&rsAlba("cost")&", "
		objBuilder.Append "	sort_seq = "&sort_seq&" "
		objBuilder.Append "WHERE  cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsAlba("company")&"' "
		objBuilder.Append "	AND bonbu = '"&rsAlba("bonbu")&"' "
		objBuilder.Append "	AND saupbu = '"&rsAlba("saupbu")&"' "
		objBuilder.Append "	AND team = '"&rsAlba("team")&"' "
		objBuilder.Append "	AND org_name = '"&rsAlba("org_name")&"' "
		objBuilder.Append "	AND cost_id = '인건비' "
		objBuilder.Append "	AND cost_detail = '알바비' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_alba.Close()

	rsAlba.MoveNext()
Loop
Set rs_alba = Nothing
rsAlba.Close() : Set rsAlba = Nothing
%>