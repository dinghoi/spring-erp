<%
Dim rsCardTran, rsCardSlip, rs_cost

' 회사 차량 운행 주유카드 셋팅
objBuilder.Append "SELECT mg_ce_id "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE (run_date >='"&from_date&"' AND run_date <='"&to_date&"') "
objBuilder.Append "	AND car_owner = '회사' "
objBuilder.Append "	AND bonbu = '"&deptName&"' "
objBuilder.Append "GROUP BY mg_ce_id "

Set rsCardTran = Server.CreateObject("ADODB.RecordSet")
rsCardTran.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCardTran.EOF
	objBuilder.Append "UPDATE card_slip SET "
	objBuilder.Append "	com_drv_yn = 'Y' "
	objBuilder.Append "WHERE (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
	objBuilder.Append "	AND emp_no='"&rsCardTran("mg_ce_id")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsCardTran.MoveNext()
Loop

rsCardTran.Close() : Set rsCardTran = Nothing

' 카드비용 집계
objBuilder.Append "SELECT owner_company as emp_company, bonbu, saupbu, team, org_name, account, SUM(cost) AS cost "
objBuilder.Append "FROM card_slip "
objBuilder.Append "WHERE (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND (card_type NOT LIKE '%주유%' OR com_drv_yn = 'Y') "
objBuilder.Append "	AND bonbu = '"&deptName&"'"
objBuilder.Append "GROUP BY owner_company, bonbu, team, org_name, account "

Set rsCardSlip = Server.CreateObject("ADODB.RecordSet")
rsCardSlip.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCardSlip.EOF
	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM org_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND emp_company ='"&rsCardSlip("emp_company")&"' "
	objBuilder.Append "	AND bonbu ='"&rsCardSlip("bonbu")&"' "
	objBuilder.Append "	AND saupbu ='"&rsCardSlip("saupbu")&"' "
	objBuilder.Append "	AND team ='"&rsCardSlip("team")&"' "
	objBuilder.Append "	AND org_name ='"&rsCardSlip("org_name")&"' "
	objBuilder.Append "	AND cost_id ='법인카드' "
	objBuilder.Append "	AND cost_detail ='"&rsCardSlip("account")&"' "

	Set rs_cost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_cost.EOF Or rs_cost.BOF Then
		objBuilder.Append "INSERT INTO org_cost(cost_year, emp_company, bonbu, saupbu, team, "
		objBuilder.Append "org_name, cost_id, cost_detail, cost_amt_"&cost_month&")"
		objBuilder.Append "VALUES("
		objBuilder.Append "'"&cost_year&"',"
		objBuilder.Append "'"&rsCardSlip("emp_company")&"', "
		objBuilder.Append "'"&rsCardSlip("bonbu")&"', "
		objBuilder.Append "'"&rsCardSlip("saupbu")&"', "
		objBuilder.Append "'"&rsCardSlip("team")&"', "
		objBuilder.Append "'"&rsCardSlip("org_name")&"', "
		objBuilder.Append "'법인카드', "
		objBuilder.Append "'"&rsCardSlip("account")&"', "
		objBuilder.Append rsCardSlip("cost")&") "
	Else
		objBuilder.Append "UPDATE org_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsCardSlip("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND emp_company = '"&rsCardSlip("emp_company")&"' "
		objBuilder.Append "	AND bonbu ='"&rsCardSlip("bonbu")&"' "
		objBuilder.Append "	AND saupbu ='"&rsCardSlip("saupbu")&"' "
		objBuilder.Append "	AND team ='"&rsCardSlip("team")&"' "
		objBuilder.Append "	AND org_name ='"&rsCardSlip("org_name")&"' "
		objBuilder.Append "	AND cost_id ='법인카드' "
		objBuilder.Append "	AND cost_detail ='"&rsCardSlip("account")&"' "
	End If

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsCardSlip.MoveNext()
Loop
rsCardSlip.Close() : Set rsCardSlip = Nothing
%>