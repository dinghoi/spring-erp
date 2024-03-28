<%
Dim rsGeneral, rs_gCost
Dim rs_endGeneral

'야특근 마감 처리
objBuilder.Append "UPDATE overtime SET "
objBuilder.Append "	end_yn = 'Y' "
objBuilder.Append "WHERE (work_date >= '"&from_date&"' AND work_date <= '"&to_date&"') "
objBuilder.Append "	AND bonbu = '' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'일반 경비 마감 처리
objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	end_yn = 'Y' "
objBuilder.Append "WHERE (slip_date >= '"&from_date&"' AND slip_date <= '"&to_date&"') "
objBuilder.Append "	AND bonbu = '' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'DB SUM 처리(비용)
objBuilder.Append "SELECT emp_company, bonbu, saupbu, team, org_name, account, SUM(cost) AS cost "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE slip_gubun = '비용' AND cancel_yn = 'N' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND bonbu = '' "
objBuilder.Append "GROUP BY emp_company, bonbu, saupbu, team, org_name, account "

Set rsGeneral = Server.CreateObject("ADODB.RecordSet")
rsGeneral.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsGeneral.EOF
	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsGeneral("emp_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsGeneral("bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsGeneral("saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsGeneral("team")&"' "
	objBuilder.Append "	AND org_name ='"&rsGeneral("org_name")&"' "
	objBuilder.Append "	AND cost_id ='일반경비' "
	objBuilder.Append "	AND cost_detail ='"&rsGeneral("account")&"' "

	Set rs_gCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_gCost.EOF Or rs_gCost.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "	org_name, cost_id, cost_detail, cost_amt_"&cost_month&")VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsGeneral("emp_company")&"', "
		objBuilder.Append "'"&rsGeneral("bonbu")&"', "
		objBuilder.Append "'"&rsGeneral("saupbu")&"', "
		objBuilder.Append "'"&rsGeneral("team")&"', "
		objBuilder.Append "'"&rsGeneral("org_name")&"', "
		objBuilder.Append "'일반경비', "
		objBuilder.Append "'"&rsGeneral("account")&"', "
		objBuilder.Append rsGeneral("cost")&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsGeneral("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsGeneral("emp_company")&"' AND bonbu ='"&rsGeneral("bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsGeneral("saupbu")&"' AND team ='"&rsGeneral("team")&"' "
		objBuilder.Append "	AND org_name ='"&rsGeneral("org_name")&"' AND cost_id ='일반경비' "
		objBuilder.Append "	AND cost_detail ='"&rsGeneral("account")&"' "
	End If	
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_gCost.Close()

	rsGeneral.MoveNext()
Loop
Set rs_gCost = Nothing
rsGeneral.Close() : Set rsGeneral = Nothing

Dim rsEctCost, rs_nCost

'DB SUM 처리(비용 외)
objBuilder.Append "SELECT emp_company, bonbu, saupbu, team, org_name, slip_gubun, account, SUM(cost) AS cost "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE slip_gubun <> '비용' AND cancel_yn = 'N'  "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"')"
objBuilder.Append "	AND bonbu = '' "
objBuilder.Append "GROUP BY slip_gubun, emp_company, bonbu, saupbu, team, org_name, account "

Set rsEctCost = Server.CreateObject("ADODB.RecordSet")
rsEctCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsEctCost.EOF
	cost_id = rsEctCost("slip_gubun")

	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsEctCost("emp_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsEctCost("bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsEctCost("saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsEctCost("team")&"' "
	objBuilder.Append "	AND org_name ='"&rsEctCost("org_name")&"' "
	objBuilder.Append "	AND cost_id ='"&cost_id&"' "
	objBuilder.Append "	AND cost_detail ='"&rsEctCost("account")&"' "

	Set rs_nCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_nCost.EOF Or rs_nCost.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "	org_name, cost_id, cost_detail, cost_amt_"&cost_month&")VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsEctCost("emp_company")&"', "
		objBuilder.Append "'"&rsEctCost("bonbu")&"', "
		objBuilder.Append "'"&rsEctCost("saupbu")&"', "
		objBuilder.Append "'"&rsEctCost("team")&"', "
		objBuilder.Append "'"&rsEctCost("org_name")&"', "
		objBuilder.Append "'"&cost_id&"', "
		objBuilder.Append "'"&rsEctCost("account")&"', "
		objBuilder.Append rsEctCost("cost")&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsEctCost("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsEctCost("emp_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsEctCost("bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsEctCost("saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsEctCost("team")&"' "
		objBuilder.Append "	AND org_name ='"&rsEctCost("org_name")&"' "
		objBuilder.Append "	AND cost_id ='"&cost_id&"' "
		objBuilder.Append "	AND cost_detail ='"&rsEctCost("account")&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_nCost.Close()

	rsEctCost.MoveNext()
Loop
Set rs_nCost = Nothing
rsEctCost.Close() : Set rsEctCost = Nothing
%>