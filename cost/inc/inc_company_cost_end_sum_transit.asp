<%
' 유류비,주차비,톨비,대중교통비
'sql = "select mg_saupbu,cost_center,company,car_owner,sum(somopum+oil_price+fare+parking+toll) as cost from transit_cost where (cancel_yn = 'N') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by mg_saupbu,cost_center,company,car_owner"
objBuilder.Append "SELECT mg_saupbu, cost_center, company, car_owner, "
objBuilder.Append "	SUM(somopum+oil_price + fare + parking+toll) AS cost "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE cancel_yn = 'N' "
objBuilder.Append "	AND (run_date >='"&from_date&"' AND run_date <='"&to_date&"') "
objBuilder.Append "GROUP BY mg_saupbu, cost_center, company, car_owner "

Set rsTranMg = Server.CreateObject("ADODB.RecordSet")
rsTranMg.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTranMg.EOF
	group_name = ""
	bill_trade_name = ""

	If rsTranMg("cost_center") = "상주직접비" Then
		'sql = "select * from trade where trade_name = '"&rs("company")&"'"
		objBuilder.Append "SELECT group_name, bill_trade_name "
		objBuilder.Append "FROM trade "
		objBuilder.Append "WHERE trade_name = '"&rsTranMg("company")&"' "

		Set rsTranMgTrade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsTranMgTrade.EOF Or rsTranMgTrade.BOF Then
			group_name = "Error"
			bill_trade_name = "Error"
		Else
			group_name = rsTranMgTrade("group_name")
			bill_trade_name = rsTranMgTrade("bill_trade_name")
		End If
		rsTranMgTrade.Close()
	End If

	'sql = "select * from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='교통비' and cost_detail ='"&rs("car_owner")&"' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND company ='"&rsTranMg("company")&"'"
	objBuilder.Append "	AND group_name ='"&group_name&"' "
	objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"'"
	objBuilder.Append "	AND cost_id ='교통비' "
	objBuilder.Append "	AND cost_detail ='"&rsTranMg("car_owner")&"'"
	objBuilder.Append "	AND cost_center ='"&rsTranMg("cost_center")&"' "
	objBuilder.Append "	AND saupbu ='"&rsTranMg("mg_saupbu")&"' "

	Set rsTranCompanyCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsTranCompanyCost.EOF Or rsTranCompanyCost.BOF Then
		'sql = "insert into company_cost (cost_year,cost_center,company,group_name,bill_trade_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("company")&"','"&group_name&"','"&bill_trade_name&"','교통비','"&rs("car_owner")&"','"&rs("mg_saupbu")&"',"&rs("cost")&")"
		objBuilder.Append "INSERT INTO company_cost(cost_year, cost_center, company, "
		objBuilder.Append "group_name, bill_trade_name, cost_id, "
		objBuilder.Append "cost_detail, saupbu, cost_amt_"&cost_month&")VALUES("
		objBuilder.Append "'"&cost_year&"', '"&rsTranMg("cost_center")&"', '"&rsTranMg("company")&"', "
		objBuilder.Append "'"&group_name&"', '"&bill_trade_name&"', '교통비', "
		objBuilder.Append "'"&rsTranMg("car_owner")&"', '"&rsTranMg("mg_saupbu")&"', "&rsTranMg("cost")&") "
	Else
		'sql = "update company_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='교통비' and cost_detail ='"&rs("car_owner")&"' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
		objBuilder.Append "UPDATE company_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsTranMg("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND company ='"&rsTranMg("company")&"' "
		objBuilder.Append "	AND group_name ='"&group_name&"' "
		objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
		objBuilder.Append "	AND cost_id ='교통비' "
		objBuilder.Append "	AND cost_detail ='"&rsTranMg("car_owner")&"' "
		objBuilder.Append "	AND cost_center ='"&rsTranMg("cost_center")&"' "
		objBuilder.Append "	AND saupbu ='"&rsTranMg("mg_saupbu")&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rsTranCompanyCost.Close()

	rsTranMg.MoveNext()
Loop
rsTranMg.Close() : Set rsTranMg = Nothing

' 차량수리비
'sql = "select mg_saupbu,cost_center,company,car_owner,sum(repair_cost) as cost from transit_cost where (cancel_yn = 'N') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by mg_saupbu,cost_center,company,car_owner"
objBuilder.Append "SELECT mg_saupbu, cost_center, company, car_owner, SUM(repair_cost) AS cost "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE cancel_yn = 'N' "
objBuilder.Append "	AND (run_date >='"&from_date&"' AND run_date <='"&to_date&"') "
objBuilder.Append "GROUP BY mg_saupbu, cost_center, company, car_owner "

Set rsRepair = Server.CreateObject("ADODB.RecordSet")
rsRepair.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsRepair.EOF
	group_name = ""
	bill_trade_name = ""

	If rsRepair("cost_center") = "상주직접비" Then
		'sql = "select * from trade where trade_name = '"&rs("company")&"'"
		objBuilder.Append "SELECT group_name, bill_trade_name "
		objBuilder.Append "FROM trade "
		objBuilder.Append "WHERE trade_name = '"&rsRepair("company")&"' "

		Set rsRepairTrade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsRepairTrade.EOF Or rsRepairTrade.BOF Then
			group_name = "Error"
			bill_trade_name = "Error"
		Else
			group_name = rsRepairTrade("group_name")
			bill_trade_name = rsRepairTrade("bill_trade_name")
		End If
		rsRepairTrade.Close()
	End If

	'sql = "select * from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='교통비' and cost_detail ='차량수리비' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
	objBuilder.Append "SELECT cost_year "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
	objBuilder.Append "	AND company ='"&rsRepair("company")&"'"
	objBuilder.Append "	AND group_name ='"&group_name&"' "
	objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"'"
	objBuilder.Append "	AND cost_id ='교통비' "
	objBuilder.Append "	AND cost_detail ='차량수리비'"
	objBuilder.Append "	AND cost_center ='"&rsRepair("cost_center")&"' "
	objBuilder.Append "	AND saupbu ='"&rsRepair("mg_saupbu")&"' "

	Set rsRepairCompanyCost = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsRepairCompanyCost.EOF Or rsRepairCompanyCost.BOF Then
		'sql = "insert into company_cost (cost_year,cost_center,company,group_name,bill_trade_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("company")&"','"&group_name&"','"&bill_trade_name&"','교통비','차량수리비','"&rs("mg_saupbu")&"',"&rs("cost")&")"
		objBuilder.Append "INSERT INTO company_cost(cost_year, cost_center, company, "
		objBuilder.Append "group_name, bill_trade_name, cost_id, "
		objBuilder.Append "cost_detail, saupbu, cost_amt_"&cost_month&")VALUES("
		objBuilder.Append "'"&cost_year&"','"&rsRepair("cost_center")&"','"&rsRepair("company")&"', "
		objBuilder.Append "'"&group_name&"','"&bill_trade_name&"','교통비',"
		objBuilder.Append "'차량수리비','"&rsRepair("mg_saupbu")&"',"&rsRepair("cost")&")"
	Else
		'sql = "update company_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='교통비' and cost_detail ='차량수리비' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
		objBuilder.Append "UPDATE company_cost SET "
		objBuilder.Append "	cost_amt_"&cost_month&"="&rsRepair("cost")&" "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND company ='"&rsRepair("company")&"' "
		objBuilder.Append "	AND group_name ='"&group_name&"' "
		objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"'"
		objBuilder.Append "	AND cost_id ='교통비' "
		objBuilder.Append "	AND cost_detail ='차량수리비' "
		objBuilder.Append "	AND cost_center ='"&rsRepair("cost_center")&"' "
		objBuilder.Append "	AND saupbu ='"&rsRepair("mg_saupbu")&"'"
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rsRepairCompanyCost.Close()

	rsRepair.MoveNext()
Loop
rsRepair.Close() : Set rsRepair = Nothing
%>