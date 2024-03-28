<%
'상주처가 있는 경우 상주직접비 지정
'sql = "update card_slip set cost_center = '상주직접비' where (pl_yn = 'Y') and (reside_company <> '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"

objBuilder.Append "UPDATE card_slip SET "
objBuilder.Append "	cost_center = '상주직접비' "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND reside_company <> '' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'상주처가 없는 경우
'sql = "select org_name from card_slip where (pl_yn = 'Y') and (reside_company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by org_name"

'objBuilder.Append "SELECT org_name "
objBuilder.Append "SELECT emp_no "
objBuilder.Append "FROM card_slip "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND reside_company = '' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
'objBuilder.Append "GROUP BY org_name "
objBuilder.Append "GROUP BY emp_no "

Set rsCard = Server.CreateObject("ADODB.RecordSet")
rsCard.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCard.EOF
	'sql = "select org_cost_center from emp_org_mst_month where org_month = '"&end_month&"' and org_name = '"&rs("org_name")&"' group by org_name"
	'objBuilder.Append "SELECT org_cost_center "
	'objBuilder.Append "FROM emp_org_mst_month "
	'objBuilder.Append "WHERE org_month = '"&end_month&"' "
	'objBuilder.Append "	AND org_name = '"&rsCard("org_name")&"' "
	'objBuilder.Append "GROUP BY org_name "
	objBuilder.Append "SELECT cost_center "
	objBuilder.Append "FROM emp_master_month "
	objBuilder.Append "WHERE emp_month = '"&end_month&"' "
	objBuilder.Append "	AND emp_no = '"&rsCard("emp_no")&"' "

	Set rsCardOrg = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not(rsCardOrg.BOF Or rsCardOrg.EOF) Then

		'sql = "update card_slip set cost_center = '"&rs_org("org_cost_center")&"' where (pl_yn = 'Y') and (reside_company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and org_name = '"&rs("org_name")&"'"
		objBuilder.Append "UPDATE card_slip SET "

		'ICT본부 프로젝트 관련 직접비 강제 지정
		Select Case rsCard("emp_no")
			Case "102764", "102765", "102766", "102767", "102768", "102769"
				objBuilder.Append "	cost_center = '직접비' "
			Case Else
				objBuilder.Append "	cost_center = '"&rsCardOrg("cost_center")&"' "
		End Select

		objBuilder.Append "WHERE pl_yn = 'Y' "
		objBuilder.Append "	AND reside_company = '' "
		objBuilder.Append "	AND slip_date >= '"&from_date&"' AND slip_date <= '"&to_date&"' "

		objBuilder.Append "	AND emp_no = '"&rsCard("emp_no")&"' "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If
	rsCardOrg.Close()

	rsCard.MoveNext()
Loop
Set rsCardOrg = Nothing
rsCard.Close() : Set rsCard = Nothing

' 카드사용 직접비 관리사업부 지정
'sql = "select saupbu from card_slip where (pl_yn = 'Y') and (cost_center = '직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by saupbu"

objBuilder.Append "SELECT bonbu "
objBuilder.Append "FROM card_slip "
objBuilder.Append "WHERE pl_yn = 'Y' AND cost_center = '직접비' "
objBuilder.Append "AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "GROUP BY bonbu "

Set rsCardCost = Server.CreateObject("ADODB.RecordSet")
rsCardCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCardCost.EOF
	'sql = "update card_slip set mg_saupbu = '"&rsCardList("org_bonbu")&"' where (pl_yn = 'Y') and (cost_center = '직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') AND (emp_no = '"&rsCardList("emp_no")&"')"

	objBuilder.Append "UPDATE card_slip SET "

	'ICT본부 프로젝트 관리사업부 강제 지정
	If rsCardCost("bonbu") = "기술연구소" Then
		objBuilder.Append "	mg_saupbu = 'ICT본부' "
	Else
		objBuilder.Append "	mg_saupbu = '"&rsCardCost("bonbu")&"' "
	End If

	objBuilder.Append "WHERE pl_yn = 'Y' AND cost_center = '직접비' "
	objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
	objBuilder.Append "	AND bonbu = '"&rsCardCost("bonbu")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsCardCost.MoveNext()
Loop
rsCardCost.Close() : Set rsCardCost = Nothing

' 카드사용 상주직접비 관리사업부 지정
'sql = "select reside_company from card_slip where (pl_yn = 'Y') and (cost_center = '상주직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by reside_company"

'objBuilder.Append "SELECT calt.emp_no, emmt.mg_saupbu "
'objBuilder.Append "SELECT calt.reside_company "
objBuilder.Append "SELECT calt.emp_no, calt.bonbu, calt.reside_company "
objBuilder.Append "FROM card_slip AS calt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON calt.emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&end_month&"' "
objBuilder.Append "WHERE calt.pl_yn = 'Y' "
objBuilder.Append "	AND calt.cost_center = '상주직접비' "
objBuilder.Append "	AND (calt.slip_date >='"&from_date&"' AND calt.slip_date <='"&to_date&"') "
'objBuilder.Append "GROUP BY calt.emp_no, emmt.mg_saupbu "
'objBuilder.Append "GROUP BY calt.reside_company "

Set rsCardOutCost = Server.CreateObject("ADODB.RecordSet")
rsCardOutCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCardOutCost.EOF
	'sql = "select saupbu from trade where trade_name = '"&rs("reside_company")&"'"
	objBuilder.Append "SELECT saupbu "
	objBuilder.Append "FROM trade "
	objBuilder.Append "WHERE trade_name = '"&rsCardOutCost("reside_company")&"' "

	Set rsCardOutCostTrade = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsCardOutCostTrade.EOF Or rsCardOutCostTrade.BOF Then
		deptName = "Error"
	Else

		If rsCardOutCost("reside_company") = "" Then
			deptName = rsCardOutCost("bonbu")
		Else
			deptName = rsCardOutCostTrade("saupbu")
		End If
	End If
	rsCardOutCostTrade.Close()

	'sql = "update card_slip set mg_saupbu = '"&deptName&"' where (pl_yn = 'Y') and (cost_center = '상주직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (reside_company = '"&rsCardList("reside_company")&"')"
	objBuilder.Append "UPDATE card_slip SET "
	objBuilder.Append "	mg_saupbu = '"&deptName&"' "
	objBuilder.Append "WHERE pl_yn = 'Y' AND cost_center = '상주직접비' "
	objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"')"
	objBuilder.Append "	AND (reside_company = '"&rsCardOutCost("reside_company")&"')"

	objBuilder.Append "	AND emp_no = '"&rsCardOutCost("emp_no")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsCardOutCost.MoveNext()
Loop
rsCardOutCost.Close() : Set rsCardOutCost = Nothing
%>