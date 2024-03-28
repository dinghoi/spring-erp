<%
Dim rsTransit, rs_orgTran, rsRepair, rs_orgRepair

'교통비 마감
objBuilder.Append "Update transit_cost SET "
objBuilder.Append "	end_yn='Y' "
objBuilder.Append "WHERE (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
objBuilder.Append "AND bonbu = '' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "SELECT emp_company, bonbu, saupbu, team, org_name, car_owner, SUM(somopum + oil_price + fare + parking + toll) AS cost "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE cancel_yn = 'N' "
objBuilder.Append "	AND (run_date >='"&from_date&"' AND run_date <='"&to_date&"') "
objBuilder.Append "	AND bonbu = '' "
objBuilder.Append "GROUP BY emp_company, bonbu, saupbu, team, org_name, car_owner "

Set rsTransit = Server.CreateObject("ADODB.RecordSet")
rsTransit.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTransit.EOF
	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsTransit("emp_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsTransit("bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsTransit("saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsTransit("team")&"' "
	objBuilder.Append "	AND org_name ='"&rsTransit("org_name")&"' "
	objBuilder.Append "	AND cost_id ='교통비' "
	objBuilder.Append "	AND cost_detail ='"&rsTransit("car_owner")&"' "

	Set rs_orgTran = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_orgTran.EOF Or rs_orgTran.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&")"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsTransit("emp_company")&"', "
		objBuilder.Append "'"&rsTransit("bonbu")&"', "
		objBuilder.Append "'"&rsTransit("saupbu")&"', "
		objBuilder.Append "'"&rsTransit("team")&"', "
		objBuilder.Append "'"&rsTransit("org_name")&"', "
		objBuilder.Append "'교통비', "
		objBuilder.Append "'"&rsTransit("car_owner")&"', "
		objBuilder.Append rsTransit("cost")&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "cost_amt_"&cost_month&"="&rsTransit("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsTransit("emp_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsTransit("bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsTransit("saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsTransit("team")&"' "
		objBuilder.Append "	AND org_name ='"&rsTransit("org_name")&"' "
		objBuilder.Append "	AND cost_id ='교통비' "
		objBuilder.Append "	AND cost_detail ='"&rsTransit("car_owner")&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_orgTran.Close()

	rsTransit.MoveNext()
Loop
Set rs_orgTran = Nothing
rsTransit.Close() : Set rsTransit = Nothing

'DB SUM 교통비(차량수리비)
objBuilder.Append "SELECT emp_company, bonbu, saupbu, team, org_name, SUM(repair_cost) AS cost "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE cancel_yn = 'N' "
objBuilder.Append "	AND repair_cost > 0 "
objBuilder.Append "	AND (run_date >= '"&from_date&"' AND run_date <='"&to_date&"') "
objBuilder.Append "	AND bonbu = '' "
objBuilder.Append "GROUP BY emp_company, bonbu, saupbu, team, org_name "

Set rsRepair = Server.CreateObject("ADODB.RecordSet")
rsRepair.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsRepair.EOF
	objBuilder.Append "SELECT cost_yar "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsRepair("emp_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsRepair("bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsRepair("saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsRepair("team")&"' "
	objBuilder.Append "	AND org_name ='"&rsRepair("org_name")&"' "
	objBuilder.Append "	AND cost_id ='교통비' "
	objBuilder.Append "	AND cost_detail ='차량수리비' "

	Set rs_orgRepair = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_orgRepair.EOF Or rs_orgRepair.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "	org_name, cost_id, cost_detail, cost_amt_"&cost_month&")"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsRepair("emp_company")&"', "
		objBuilder.Append "'"&rsRepair("bonbu")&"', "
		objBuilder.Append "'"&rsRepair("saupbu")&"', "
		objBuilder.Append "'"&rsRepair("team")&"', "
		objBuilder.Append "'"&rsRepair("org_name")&"', "
		objBuilder.Append "'교통비', "
		objBuilder.Append "'차량수리비', "
		objBuilder.Append rsRepair("cost")&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsRepair("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsRepair("emp_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsRepair("bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsRepair("saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsRepair("team")&"' "
		objBuilder.Append "	AND org_name ='"&rsRepair("org_name")&"' "
		objBuilder.Append "	AND cost_id ='교통비' "
		objBuilder.Append "	AND cost_detail ='차량수리비' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rs_orgRepair.Close()

	rsRepair.MoveNext()
Loop

Set rs_orgRepair = Nothing
rsRepair.Close() : Set rsRepair = Nothing
%>