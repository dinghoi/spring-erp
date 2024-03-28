<%
'sql = "select mg_saupbu,cost_center,reside_company as company,account,sum(cost) as cost "&_
'		 "  from card_slip "&_
'		 " where (pl_yn = 'Y') "&_
'		 "    and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "&_
'		 "    and (card_type not like '%주유%' or com_drv_yn = 'Y') "&_
'		 "  group by  mg_saupbu,cost_center,reside_company,account"
objBuilder.Append "SELECT mg_saupbu, cost_center, reside_company AS company, account, SUM(cost) AS cost "
objBuilder.Append "FROM card_slip "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND (card_type NOT LIKE '%주유%' OR com_drv_yn = 'Y') "
objBuilder.Append "GROUP BY mg_saupbu, cost_center, reside_company, account"

Set rsCardMg = Server.CreateObject("ADODB.RecordSet")
rsCardMg.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCardMg.EOF
	group_name = ""
	bill_trade_name = ""

	If rsCardMg("cost_center") = "상주직접비" Then
		'sql = "select * from trade where trade_name = '"&rs("company")&"'"
		objBuilder.Append "SELECT group_name, bill_trade_name "
		objBuilder.Append "FROM trade "
		objBuilder.Append "WHERE trade_name = '"&rsCardMg("company")&"' "

		Set rsCardMgTrade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsCardMgTrade.EOF Or rsCardMgTrade.BOF Then
			group_name = "Error"
			bill_trade_name = "Error"
		Else
			group_name = rsCardMgTrade("group_name")
			bill_trade_name = rsCardMgTrade("bill_trade_name")
		End If
		rsCardMgTrade.Close()
	End If

	'sql = "select * from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='법인카드' and cost_detail ='"&rs("account")&"' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
	objBuilder.Append "	AND company = '"&rsCardMg("company")&"' "
	objBuilder.Append "	AND group_name = '"&group_name&"' "
	objBuilder.Append "	AND bill_trade_name = '"&bill_trade_name&"' "
	objBuilder.Append "	AND cost_id = '법인카드' "
	objBuilder.Append "	AND cost_detail = '"&rsCardMg("account")&"' "
	objBuilder.Append "	AND cost_center = '"&rsCardMg("cost_center")&"' "
	objBuilder.Append "	AND saupbu = '"&rsCardMg("mg_saupbu")&"'"

	Set rsCardCompanyCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsCardCompanyCost.EOF Or rsCardCompanyCost.BOF Then
		'sql = "insert into company_cost "&_
		'		 " (   cost_year "&_
		'		 "   , cost_center "&_
		'		 "   , company "&_
		'		 "   , group_name "&_
		'		 "   , bill_trade_name "&_
		'		 "   , cost_id "&_
		'		 "   , cost_detail "&_
		'		 "   , saupbu "&_
		'		 "   , cost_amt_"&cost_month&_
		'		 ") "&_
		'		 " values "&_
		'		 " (   '"&cost_year&"' "&_
		'		 "   , '"&rs("cost_center")&"' "&_
		'		 "   , '"&rs("company")&"' "&_
		'		 "   , '"&group_name&"' "&_
		'		 "   , '"&bill_trade_name&"' "&_
		'		 "   , '법인카드' "&_
		'		 "   , '"&rs("account")&"' "&_
		'		 "   , '"&rs("mg_saupbu")&"' "&_
		'		 "   , "&rs("cost") &_
		'		 "   )"
		objBuilder.Append "INSERT INTO company_cost("
		objBuilder.Append "cost_year, cost_center, company, "
		objBuilder.Append "group_name, bill_trade_name, cost_id, "
		objBuilder.Append "cost_detail, saupbu, cost_amt_"&cost_month&")VALUES("
		objBuilder.Append "'"&cost_year&"', '"&rsCardMg("cost_center")&"', '"&rsCardMg("company")&"', "
		objBuilder.Append "'"&group_name&"', '"&bill_trade_name&"', '법인카드', "
		objBuilder.Append "'"&rsCardMg("account")&"', '"&rsCardMg("mg_saupbu")&"', "&rsCardMg("cost")&")"
	  Else
		'sql = "update company_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='법인카드' and cost_detail ='"&rs("account")&"' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"

		objBuilder.Append "UPDATE company_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsCardMg("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND company ='"&rsCardMg("company")&"' "
		objBuilder.Append "	AND group_name ='"&group_name&"' "
		objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
		objBuilder.Append "	AND cost_id ='법인카드' "
		objBuilder.Append "	AND cost_detail ='"&rsCardMg("account")&"' "
		objBuilder.Append "	AND cost_center ='"&rsCardMg("cost_center")&"' "
		objBuilder.Append "	AND saupbu ='"&rsCardMg("mg_saupbu")&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rsCardCompanyCost.Close()

	rsCardMg.MoveNext()
Loop
rsCardMg.Close() : Set rsCardMg = Nothing
%>