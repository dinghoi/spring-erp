<%
sort_seq = 8

'sql = "select slip_gubun,cost_center,mg_saupbu,company,account,sum(cost) as cost from general_cost where (pl_yn = 'Y') and (cancel_yn = 'N') and (skip_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by slip_gubun,cost_center,mg_saupbu,company,account"
objBuilder.Append "SELECT slip_gubun, cost_center, mg_saupbu, company, account, SUM(cost) AS cost "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND cancel_yn = 'N' "
objBuilder.Append "	AND skip_yn = 'N' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "GROUP BY slip_gubun, cost_center, mg_saupbu, company, account "

Set rsCostSum = Server.CreateObject("ADODB.RecordSet")
rsCostSum.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCostSum.EOF
	cost_id = rsCostSum("slip_gubun")

	If cost_id = "비용" Then
		cost_id = "일반경비"
	End If

	group_name = ""
	bill_trade_name = ""

	If rsCostSum("cost_center") = "상주직접비" Then
		'sql = "select * from trade where trade_name = '"&rs("company")&"'"
		objBuilder.Append "SELECT group_name, bill_trade_name "
		objBuilder.Append "FROM trade "
		objBuilder.Append "WHERE trade_name = '"&rsCostSum("company")&"' "

		Set rsCostSumTrade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsCostSumTrade.EOF Or rsCostSumTrade.BOF Then
			group_name = "Error"
			bill_trade_name = "Error"
		Else
			group_name = rsCostSumTrade("group_name")
			bill_trade_name = rsCostSumTrade("bill_trade_name")
		End If

		rsCostSumTrade.Close()
	End If

	'sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and cost_center ='"&rs("cost_center")&"' and company ='"&rs("company")&"' and cost_id ='"&cost_id&"' and cost_detail ='"&rs("account")&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"'"
	objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND cost_center ='"&rsCostSum("cost_center")&"' "
	objBuilder.Append "	AND company ='"&rsCostSum("company")&"' "
	objBuilder.Append "	AND cost_id ='"&cost_id&"' "
	objBuilder.Append "	AND cost_detail ='"&rsCostSum("account")&"' "
	objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
	objBuilder.Append "	AND group_name ='"&group_name&"' "
	objBuilder.Append "	AND saupbu ='"&rsCostSum("mg_saupbu")&"' "

	Set rsCompanyCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsCompanyCost.EOF Or rsCompanyCost.BOF Then
		'sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("company")&"','"&bill_trade_name&"','"&group_name&"','"&cost_id&"','"&rs("account")&"','"&rs("mg_saupbu")&"',"&rs("cost")&","&sort_seq&")"
		objBuilder.Append "INSERT INTO company_cost(cost_year, cost_center, company,"
		objBuilder.Append "bill_trade_name, group_name, cost_id, "
		objBuilder.Append "cost_detail, saupbu, cost_amt_"&cost_month&", "
		objBuilder.Append "sort_seq)VALUES("
		objBuilder.Append "'"&cost_year&"', '"&rsCostSum("cost_center")&"', '"&rsCostSum("company")&"', "
		objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '"&cost_id&"', "
		objBuilder.Append "'"&rsCostSum("account")&"', '"&rsCostSum("mg_saupbu")&"', "&rsCostSum("cost")&", "
		objBuilder.Append sort_seq&")"
	Else
		sum_cost = CLng(rsCompanyCost("cost")) + Cdbl(rsCostSum("cost"))

		'sql = "update company_cost set cost_amt_"&cost_month&"="&sum_cost&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and cost_center ='"&rs("cost_center")&"' and company ='"&rs("company")&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and cost_id ='"&cost_id&"' and cost_detail ='"&rs("account")&"' and saupbu ='"&rs("mg_saupbu")&"'"
		objBuilder.Append "UPDATE company_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&sum_cost&", "
		objBuilder.Append "	sort_seq="&sort_seq&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND cost_center ='"&rsCostSum("cost_center")&"' "
		objBuilder.Append "	AND company ='"&rsCostSum("company")&"' "
		objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"'"
		objBuilder.Append "	AND group_name ='"&group_name&"' "
		objBuilder.Append "	AND cost_id ='"&cost_id&"' "
		objBuilder.Append "	AND cost_detail ='"&rsCostSum("account")&"' "
		objBuilder.Append "	AND saupbu ='"&rsCostSum("mg_saupbu")&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rsCompanyCost.Close()

	rsCostSum.MoveNext()
Loop
rsCostSum.close() : Set rsCostSum = Nothing
%>