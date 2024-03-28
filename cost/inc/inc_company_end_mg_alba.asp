<%
' 초기값 Clear
'sql = "UPDATE pay_alba_cost SET mg_saupbu = '', cost_center = '' WHERE rever_yymm ='"&end_month&"' "

'sql = "update pay_alba_cost set cost_center = '상주직접비' where (cost_company <> '공통' and cost_company <> '전사' and cost_company <> '부문' and cost_company <> '기타' and cost_company <> '본사' and cost_company <> '케이원정보통신' and cost_company <> '') and (rever_yymm ='"&end_month&"')"

'sql = "select company,org_name from pay_alba_cost where (cost_company = '공통' or cost_company <> '전사' or cost_company <> '부문' or cost_company = '기타' or cost_company = '본사' or cost_company = '케이원정보통신' or cost_company = '') and (rever_yymm ='"&end_month&"') group by company,org_name"

objBuilder.Append "CALL USP_COMPANY_END_ALBA_INIT('"&end_month&"');"

dbconn.rollbacktrans
Response.write objBuilder.ToString()
Response.end

Set rsAlba = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsAlba.EOF Then
	arrAlba = rsAlba.getRows()
End If
rsAlba.Close() : Set rsAlba = Nothing

If IsArray(arrAlba) Then
	Do Until rsAlba.EOF
		'sql = "select org_cost_center from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("company")&"' and org_name = '"&rs("org_name")&"'"
		objBuilder.Append "SELECT org_cost_center "
		objBuilder.Append "FROM emp_org_mst_month "
		objBuilder.Append "WHERE org_month = '"&end_month&"' "
		objBuilder.Append "	AND org_company = '"&rsAlba("company")&"' "
		objBuilder.Append "	AND org_name = '"&rsAlba("org_name")&"' "

		Set rsAlbaOrg = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsAlbaOrg.EOF Or rsAlbaOrg.BOF Then
			cost_center = "전사공통비"
			cost_company = ""
			group_name = ""
			bill_trade_name = ""
		Else
			cost_center = rsAlbaOrg("org_cost_center")
			cost_company = ""
			group_name = ""
			bill_trade_name = ""
		End If
		rsAlbaOrg.Close()

		'sql = "update pay_alba_cost set cost_center = '"&cost_center&"' where (cost_company = '공통' or cost_company = '기타' or cost_company = '본사' or cost_company = '케이원정보통신' or cost_company = '') and (rever_yymm ='"&end_month&"') and org_name = '"&rs("org_name")&"'"
		objBuilder.Append "UPDATE pay_alba_cost SET "
		objBuilder.Append "	cost_center = '"&cost_center&"' "
		objBuilder.Append "WHERE rever_yymm ='"&end_month&"' "
		objBuilder.Append "	AND org_name = '"&rsAlba("org_name")&"' "
		objBuilder.Append "	AND cost_company IN ('공통', '전사', '기타', '본사', '케이원정보통신', '케이원', '') "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		rsAlba.MoveNext()
	Loop
End If

'알바비용 관리사업부 지정
'sql = "SELECT saupbu, cost_company FROM pay_alba_cost WHERE cost_center = '상주직접비' AND (rever_yymm ='"&end_month&"') GROUP BY saupbu, cost_company"
objBuilder.Append "SELECT bonbu, cost_company "
objBuilder.Append "FROM pay_alba_cost "
objBuilder.Append "WHERE cost_center = '상주직접비' "
objBuilder.Append "	AND rever_yymm ='"&end_month&"' "
objBuilder.Append "GROUP BY bonbu, cost_company "

Set rsAlbaOutCost = Server.CreateObject("ADODB.RecordSet")
rsAlbaOutCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsAlbaOutCost.EOF
	alba_bonbu = rsAlbaOutCost("bonbu")

	'sql = "SELECT sort_seq FROM sales_org WHERE saupbu = '"&saupbu&"' AND sales_year='" & cost_year & "' "
	objBuilder.Append "SELECT sort_seq "
	objBuilder.Append "FROM sales_org "
	objBuilder.Append "WHERE saupbu = '"&alba_bonbu&"' "
	objBuilder.Append "	AND sales_year='" & cost_year & "' "

	Set rsAlbaOutCostSales = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsAlbaOutCostSales.EOF Or rsAlbaOutCostSales.BOF Then
		If rsAlbaOutCost("cost_company") = "" Or IsNull(rsAlbaOutCost("cost_company")) Then
			alba_bonbu = ""
		Else
			'sql = "SELECT saupbu FROM trade WHERE trade_name = '"&rs("cost_company")&"'"
			objBuilder.Append "SELECT org_bonbu"
			objBuilder.Append "FROM emp_org_mst_month "
			objBuilder.Append "WHERE org_month = '"&end_month&"' AND org_name = '"&alba_bonbu&"' "
			objBuilder.Append "GROUP BY org_bonbu "

			Set rsAlbaOutCostTrade = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsAlbaOutCostTrade.EOF Or rsAlbaOutCostTrade.BOF Then
				alba_bonbu = "Error"
			Else
				'alba_bonbu = rsAlbaOutCostTrade("saupbu")
				alba_bonbu = rsAlbaOutCostTrade("org_bonbu")
			End If
			rsAlbaOutCostTrade.Close()
		End If
	End If
	rsAlbaOutCostSales.Close()

	'sql = "update pay_alba_cost set mg_saupbu = '"&saupbu&"' where (cost_center = '상주직접비') and (rever_yymm ='"&end_month&"') and (saupbu = '"&rs("saupbu")&"') and (cost_company = '"&rs("cost_company")&"') "
	objBuilder.Append "UPDATE pay_alba_cost SET "
	objBuilder.Append "	mg_saupbu = '"&alba_bonbu&"' "
	objBuilder.Append "WHERE cost_center = '상주직접비' "
	objBuilder.Append "	AND rever_yymm ='"&end_month&"' "
	objBuilder.Append "	AND bonbu = '"&rsAlbaOutCost("bonbu")&"' "
	objBuilder.Append "	AND cost_company = '"&rsAlbaOutCost("cost_company")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsAlbaOutCost.MoveNext()
Loop
rsAlbaOutCost.Close() : Set rsAlbaOutCost = Nothing

'알바비용 직접비 관리사업부 지정
'sql = "select saupbu from pay_alba_cost where (cost_center = '직접비') and (rever_yymm ='"&end_month&"') group by saupbu"
objBuilder.Append "SELECT bonbu "
objBuilder.Append "FROM pay_alba_cost "
objBuilder.Append "WHERE cost_center = '직접비' "
objBuilder.Append "	AND rever_yymm ='"&end_month&"' "
objBuilder.Append "GROUP BY bonbu "

Set rsAlbaCost = Server.CreateObject("ADODB.RecordSet")
rsAlbaCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsAlbaCost.EOF
	'sql = "update pay_alba_cost set mg_saupbu = '"&rs("saupbu")&"' where (cost_center = '직접비') and (rever_yymm ='"&end_month&"') and (saupbu = '"&rs("saupbu")&"')"
	objBuilder.Append "UPDATE pay_alba_cost SET "
	objBuilder.Append "	mg_saupbu = '"&rsAlbaCost("bonbu")&"' "
	objBuilder.Append "WHERE cost_center = '직접비' "
	objBuilder.Append "	AND rever_yymm ='"&end_month&"' "
	objBuilder.Append "	AND bonbu = '"&rsAlbaCost("bonbu")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsAlbaCost.MoveNext()
Loop
rsAlbaCost.Close() : Set rsAlbaCost = Nothing
%>