<%
' 檬扁拳
'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='惑林流立厚' or cost_center ='流立厚') "
objBuilder.Append "UPDATE saupbu_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&" = '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND (cost_center ='惑林流立厚' OR cost_center ='流立厚')"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 惑林流立厚 客 流立厚 诀单捞飘
'sql = "select saupbu,cost_center,cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '惑林流立厚' or cost_center = '流立厚') and cost_year ='"&cost_year&"' group by saupbu,cost_center,cost_id,cost_detail"
objBuilder.Append "SELECT saupbu, cost_center, cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE (cost_center = '惑林流立厚' OR cost_center = '流立厚') "
objBuilder.Append "	AND cost_year ='"&cost_year&"' "
objBuilder.Append "GROUP BY saupbu ,cost_center, cost_id, cost_detail "

Set rsCostCompany = Server.CreateObject("ADODB.RecordSet")
rsCostCompany.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCostCompany.EOF
	'sql = "select * from saupbu_profit_loss where cost_year ='"&cost_year&"' and saupbu ='"&rs("saupbu")&"' and cost_center ='"&rs("cost_center")&"' and cost_id ='"&rs("cost_id")&"' and cost_detail ='"&rs("cost_detail")&"'"
	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM saupbu_profit_loss "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND saupbu ='"&rsCostCompany("saupbu")&"' "
	objBuilder.Append "	AND cost_center ='"&rsCostCompany("cost_center")&"' "
	objBuilder.Append "	AND cost_id ='"&rsCostCompany("cost_id")&"' "
	objBuilder.Append "	AND cost_detail ='"&rsCostCompany("cost_detail")&"' "

	Set rsCostProfit = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsCostProfit.EOF Or rsCostProfit.BOF Then
		'sql = "insert into saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("saupbu")&"','"&rs("cost_center")&"','"&rs("cost_id")&"','"&rs("cost_detail")&"',"&rs("cost")&")"
		objBuilder.Append "INSERT INTO saupbu_profit_loss(cost_year,saupbu,cost_center, "
		objBuilder.Append "cost_id,cost_detail,cost_amt_"&cost_month&")VALUES("
		objBuilder.Append "'"&cost_year&"','"&rsCostCompany("saupbu")&"','"&rsCostCompany("cost_center")&"',"
		objBuilder.Append "'"&rsCostCompany("cost_id")&"','"&rsCostCompany("cost_detail")&"',"&rsCostCompany("cost")&")"
	Else
		'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and saupbu ='"&rs("saupbu")&"' and cost_center ='"&rs("cost_center")&"' and cost_id ='"&rs("cost_id")&"' and cost_detail ='"&rs("cost_detail")&"'"
		objBuilder.Append "UPDATE saupbu_profit_loss SET "
		objBuilder.Append "	cost_amt_"&cost_month&" = "&rsCostCompany("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND saupbu ='"&rsCostCompany("saupbu")&"' "
		objBuilder.Append "	AND cost_center ='"&rsCostCompany("cost_center")&"' "
		objBuilder.Append "	AND cost_id ='"&rsCostCompany("cost_id")&"' "
		objBuilder.Append "	AND cost_detail ='"&rsCostCompany("cost_detail")&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rsCostProfit.Close()

	rsCostCompany.MoveNext()
Loop
rsCostCompany.Close() : Set rsCostCompany = Nothing
' 荤诀何喊 颊劳 磊丰 积己 辆丰

' 雀荤喊喊 颊劳 磊丰 积己
' 贸府傈 zero
'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='惑林流立厚') "
objBuilder.Append "UPDATE company_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&"= '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_center ='惑林流立厚' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 惑林流立厚 诀单捞飘
'sql = "select company,group_name,cost_center,cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '惑林流立厚') and cost_year ='"&cost_year&"' group by company,group_name,cost_center,cost_id,cost_detail"
objBuilder.Append "SELECT company, group_name, cost_center, cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_center = '惑林流立厚' "
objBuilder.Append "	AND cost_year ='"&cost_year&"' "
objBuilder.Append "GROUP BY company, group_name, cost_center, cost_id, cost_detail "

Set rsCompanyOutCost = Server.CreateObject("ADODB.RecordSet")
rsCompanyOutCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCompanyOutCost.EOF
	'sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&rs("group_name")&"' and cost_center ='"&rs("cost_center")&"' and cost_id ='"&rs("cost_id")&"' and cost_detail ='"&rs("cost_detail")&"'"
	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM company_profit_loss "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND company ='"&rsCompanyOutCost("company")&"' "
	objBuilder.Append "	AND group_name ='"&rsCompanyOutCost("group_name")&"' "
	objBuilder.Append "	AND cost_center ='"&rsCompanyOutCost("cost_center")&"' "
	objBuilder.Append "	AND cost_id ='"&rsCompanyOutCost("cost_id")&"' "
	objBuilder.Append "	AND cost_detail ='"&rsCompanyOutCost("cost_detail")&"' "

	Set rsProfitCostList = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsProfitCostList.eof Or rsProfitCostList.bof Then
		'sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&rs("group_name")&"','"&rs("cost_center")&"','"&rs("cost_id")&"','"&rs("cost_detail")&"',"&rs("cost")&")"
		objBuilder.Append "INSERT INTO company_profit_loss(cost_year,company,group_name,"
		objBuilder.Append "cost_center,cost_id,cost_detail, "
		objBuilder.Append "cost_amt_"&cost_month&")VALUES("
		objBuilder.Append "'"&cost_year&"','"&rsCompanyOutCost("company")&"','"&rsCompanyOutCost("group_name")&"', "
		objBuilder.Append "'"&rsCompanyOutCost("cost_center")&"','"&rsCompanyOutCost("cost_id")&"','"&rsCompanyOutCost("cost_detail")&"', "
		objBuilder.Append rsCompanyOutCost("cost")&")"
	Else
		'sql = "update company_profit_loss set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&rs("group_name")&"' and cost_center ='"&rs("cost_center")&"' and cost_id ='"&rs("cost_id")&"' and cost_detail ='"&rs("cost_detail")&"'"
		objBuilder.Append "UPDATE company_profit_loss SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsCompanyOutCost("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND company ='"&rsCompanyOutCost("company")&"' "
		objBuilder.Append "	AND group_name ='"&rsCompanyOutCost("group_name")&"' "
		objBuilder.Append "	AND cost_center ='"&rsCompanyOutCost("cost_center")&"'"
		objBuilder.Append "	AND cost_id ='"&rsCompanyOutCost("cost_id")&"' "
		objBuilder.Append "	AND cost_detail ='"&rsCompanyOutCost("cost_detail")&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rsProfitCostList.Close()

	rsCompanyOutCost.MoveNext()
Loop
rsCompanyOutCost.Close() : Set rsCompanyOutCost = Nothing
%>