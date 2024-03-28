<%
' 알바비
sort_seq = 8

'sql = "select cost_center,mg_saupbu,cost_company,sum(alba_give_total) as cost from pay_alba_cost where (rever_yymm ='"&end_month&"') group by cost_center,mg_saupbu,cost_company"
objBuilder.Append "SELECT cost_center, mg_saupbu, cost_company, SUM(alba_give_total) AS cost "
objBuilder.Append "FROM pay_alba_cost "
objBuilder.Append "WHERE rever_yymm ='"&end_month&"' "
objBuilder.Append "GROUP BY cost_center, mg_saupbu, cost_company "

Set rsAlbaTot = Server.CreateObject("ADODB.RecordSet")
rsAlbaTot.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsAlbaTot.EOF
	group_name = ""
	bill_trade_name = ""

	If rsAlbaTot("cost_center") = "상주직접비" Then
		'sql = "select * from trade where trade_name = '"&rs("cost_company")&"'"
		objBuilder.Append "SELECT group_name, bill_trade_name "
		objBuilder.Append "FROM trade "
		objBuilder.Append "WHERE trade_name = '"&rsAlbaTot("cost_company")&"' "

		Set rsAlbaTotTrade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsAlbaTotTrade.EOF Or rsAlbaTotTrade.BOF Then
			group_name = "Error"
			bill_trade_name = "Error"
		Else
			group_name = rsAlbaTotTrade("group_name")
			bill_trade_name = rsAlbaTotTrade("bill_trade_name")
		End If
		rsAlbaTotTrade.Close()
	End If

	'sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and cost_center ='"&rs("cost_center")&"' and company ='"&rs("cost_company")&"' and cost_id ='인건비' and cost_detail ='알바비' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"'"
	objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND cost_center ='"&rsAlbaTot("cost_center")&"' "
	objBuilder.Append "	AND company ='"&rsAlbaTot("cost_company")&"' "
	objBuilder.Append "	AND cost_id ='인건비' "
	objBuilder.Append "	AND cost_detail ='알바비' "
	objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
	objBuilder.Append "	AND group_name ='"&group_name&"' "
	objBuilder.Append "	AND saupbu ='"&rsAlbaTot("mg_saupbu")&"' "

	Set rsAlbaCompanyCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsAlbaCompanyCost.EOF Or rsAlbaCompanyCost.BOF Then
		'sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("cost_company")&"','"&bill_trade_name&"','"&group_name&"','인건비','알바비','"&rs("mg_saupbu")&"',"&rs("cost")&","&sort_seq&")"
		objBuilder.Append "INSERT INTO company_cost(cost_year,cost_center,company, "
		objBuilder.Append "bill_trade_name,group_name,cost_id, "
		objBuilder.Append "cost_detail,saupbu,cost_amt_"&cost_month&", "
		objBuilder.Append "sort_seq)values("
		objBuilder.Append "'"&cost_year&"', '"&rsAlbaTot("cost_center")&"', '"&rsAlbaTot("cost_company")&"', "
		objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '인건비', "
		objBuilder.Append "'알바비','"&rsAlbaTot("mg_saupbu")&"',"&rsAlbaTot("cost")&", "
		objBuilder.Append sort_seq&")"
	Else
		sum_cost = CLng(rsAlbaCompanyCost("cost")) + CLng(rsAlbaTot("cost"))

		'sql = "update company_cost set cost_amt_"&cost_month&"="&sum_cost&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and cost_center ='"&rs("cost_center")&"' and company ='"&rs("cost_company")&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and cost_id ='인건비' and cost_detail ='알바비' and saupbu ='"&rs("mg_saupbu")&"'"
		objBuilder.Append "UPDATE company_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&sum_cost&", "
		objBuilder.Append "	sort_seq="&sort_seq&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND cost_center ='"&rsAlbaTot("cost_center")&"' "
		objBuilder.Append "	AND company ='"&rsAlbaTot("cost_company")&"' "
		objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
		objBuilder.Append "	AND group_name ='"&group_name&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='알바비' "
		objBuilder.Append "	AND saupbu ='"&rsAlbaTot("mg_saupbu")&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rsAlbaCompanyCost.Close()

	rsAlbaTot.MoveNext()
Loop
rsAlbaTot.Close() : Set rsAlbaTot = Nothing
' 알바비 종료
%>