<%
'설치/공사
sort_seq = 9

Dim rsComCost, tot_part_cost, rsAsSum, as_set_sum, set_time_sum, total_time_sum
Dim dist_part, dist_cost
Dim rsAsTot, rsAsTotTrade
Dim rsAsCompanyCost

'부문공통비 전체 비용
objBuilder.Append "SELECT SUM(cost_amt_"& cost_month &") AS tot_cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_year ='"& cost_year &"' "
objBuilder.Append "AND cost_center = '부문공통비' "

Set rsComCost = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsComCost("tot_cost")) Then
	tot_part_cost = 0
Else
	tot_part_cost = CLng(rsComCost("tot_cost"))
End If

rsComCost.Close() : Set rsComCost = Nothing

'AS 현황 집계
objBuilder.Append "SELECT SUM(as_set) AS 'as_set_sum', SUM(set_time) AS 'set_time_sum', SUM(total_time) AS 'total_time_sum' "
objBuilder.Append "FROM as_acpt_status "
objBuilder.Append "WHERE as_month = '"&end_month&"' "

Set rsAsSum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

as_set_sum = CLng(f_toString(rsAsSum("as_set_sum"), 0))	'설치공사 총 건수
set_time_sum = CLng(f_toString(rsAsSum("set_time_sum"), 0))	'설치공사 총 시간
total_time_sum = CLng(f_toString(rsAsSum("total_time_sum"), 0)) '총 시간

rsAsSum.Close() : Set rsAsSum = Nothing

If as_set_sum > 0 Then
	'설치공사 비율 = 설치공사 총 시간 / 총 시간 * 100
	dist_part = FormatNumber(set_time_sum / total_time_sum * 100, 1)

	'설치공사 비중  = 총 부문공통비 * 설치공사 비율(%)
	dist_cost = CDbl(FormatNumber(tot_part_cost * dist_part / 100, 0))

	'AS 현황 > 설치/공사 조회
	objBuilder.Append "SELECT saupbu, as_company, as_set, ("&dist_cost&" / "&as_set_sum&" * as_set) AS 'cost' "
	objBuilder.Append "FROM as_acpt_status AS aast "
	objBuilder.Append "INNER JOIN trade AS trat ON aast.as_company = trat.trade_name AND trade_id = '매출' "
	objBuilder.Append "WHERE as_month ='"&end_month&"' AND as_set > 0 "

	Set rsAsTot = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Do Until rsAsTot.EOF
		'company_part_cost = dist_cost / as_set_sum * rsAsTot("as_set")

		group_name = ""
		bill_trade_name = ""

		objBuilder.Append "SELECT group_name, bill_trade_name "
		objBuilder.Append "FROM trade "
		objBuilder.Append "WHERE trade_name = '"&rsAsTot("as_company")&"' "

		Set rsAsTotTrade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsAsTotTrade.EOF Or rsAsTotTrade.BOF Then
			group_name = "Error"
			bill_trade_name = "Error"
		Else
			group_name = rsAsTotTrade("group_name")
			bill_trade_name = rsAsTotTrade("bill_trade_name")
		End If
		rsAsTotTrade.Close()

		objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
		objBuilder.Append "FROM company_cost "
		objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND cost_center ='상주직접비' "
		objBuilder.Append "	AND company ='"&rsAsTot("as_company")&"' "
		objBuilder.Append "	AND cost_id ='인건비' "
		objBuilder.Append "	AND cost_detail ='설치공사' "
		objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
		objBuilder.Append "	AND group_name ='"&group_name&"' "
		objBuilder.Append "	AND saupbu ='"&rsAsTot("saupbu")&"' "

		Set rsAsCompanyCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsAsCompanyCost.EOF Or rsAsCompanyCost.BOF Then
			objBuilder.Append "INSERT INTO company_cost(cost_year,cost_center,company, "
			objBuilder.Append "bill_trade_name,group_name,cost_id, "
			objBuilder.Append "cost_detail,saupbu,cost_amt_"&cost_month&", "
			objBuilder.Append "sort_seq)values("
			objBuilder.Append "'"&cost_year&"', '상주직접비', '"&rsAsTot("as_company")&"', "
			objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '인건비', "
			objBuilder.Append "'설치공사','"&rsAsTot("saupbu")&"',"&rsAsTot("cost")&", "
			objBuilder.Append sort_seq&")"
		Else
			sum_cost = CLng(rsAsCompanyCost("cost")) + CLng(rsAsTot("cost"))

			objBuilder.Append "UPDATE company_cost SET "
			objBuilder.Append "	cost_amt_"&cost_month&"="&sum_cost&", "
			objBuilder.Append "	sort_seq="&sort_seq&" "
			objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND cost_center ='상주직접비' "
			objBuilder.Append "	AND company ='"&rsAsTot("as_company")&"' "
			objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
			objBuilder.Append "	AND group_name ='"&group_name&"' "
			objBuilder.Append "	AND cost_id ='인건비' "
			objBuilder.Append "	AND cost_detail ='설치공사' "
			objBuilder.Append "	AND saupbu ='"&rsAsTot("saupbu")&"' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rsAsCompanyCost.Close()

		rsAsTot.MoveNext()
	Loop
	Set rsAsTotTrade = Nothing
	Set rsAsCompanyCost = Nothing
	rsAsTot.Close() : Set rsAsTot = Nothing
End If

'설치/공사 END	============================================

'협업	============================================
sort_seq = 10

Dim rsCowork, arrCowork, person_cost, cw_saupbu, cw_company, cw_cost, rsCoworkCost
Dim rsCworkTrade, i

'협업 건수 인당 비용
person_cost = 30000

objBuilder.Append "SELECT saupbu, as_company, (sum(as_give_cowork * "&person_cost&" * -1) + sum(as_get_cowork * "&person_cost&")) as 'cowork_cost' "
objBuilder.Append "FROM as_acpt_status as aast "
objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
objBuilder.Append "WHERE as_month = '"&end_month&"' "
objBuilder.Append "GROUP BY saupbu, as_company "

Set rsCowork = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsCowork.EOF Then
	arrCowork = rsCowork.getRows()
End If

rsCowork.Close() : Set rsCoWork = Nothing

If IsArray(arrCowork) Then
	For i = LBound(arrCowork) To UBound(arrCowork, 2)
		cw_saupbu = arrCowork(0, i)
		cw_company = arrCowork(1, i)
		cw_cost = CDbl(arrCowork(2, i))

		group_name = ""
		bill_trade_name = ""

		objBuilder.Append "SELECT group_name, bill_trade_name "
		objBuilder.Append "FROM trade "
		objBuilder.Append "WHERE trade_name = '"&cw_company&"' "

		Set rsCworkTrade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsCworkTrade.EOF Or rsCworkTrade.BOF Then
			group_name = "Error"
			bill_trade_name = "Error"
		Else
			group_name = rsCworkTrade("group_name")
			bill_trade_name = rsCworkTrade("bill_trade_name")
		End If
		rsCworkTrade.Close()

		'If cw_cost > 0 Or cw_cost < 0 Then
			objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
			objBuilder.Append "FROM company_cost "
			objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND cost_center ='상주직접비' "
			objBuilder.Append "	AND company ='"&cw_company&"' "
			objBuilder.Append "	AND cost_id ='인건비' "
			objBuilder.Append "	AND cost_detail ='협업' "
			objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
			objBuilder.Append "	AND group_name ='"&group_name&"' "
			objBuilder.Append "	AND saupbu ='"&cw_saupbu&"' "

			Set rsCoworkCost = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsCoworkCost.EOF Or rsCoworkCost.BOF Then
				objBuilder.Append "INSERT INTO company_cost(cost_year,cost_center,company, "
				objBuilder.Append "bill_trade_name,group_name,cost_id, "
				objBuilder.Append "cost_detail,saupbu,cost_amt_"&cost_month&", "
				objBuilder.Append "sort_seq)values("
				objBuilder.Append "'"&cost_year&"', '상주직접비', '"&cw_company&"', "
				objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '인건비', "
				objBuilder.Append "'협업','"&cw_saupbu&"',"&cw_cost&", "
				objBuilder.Append sort_seq&")"
			Else
				objBuilder.Append "UPDATE company_cost SET "
				objBuilder.Append "	cost_amt_"&cost_month&"="&cw_cost&", "
				objBuilder.Append "	sort_seq="&sort_seq&" "
				objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
				objBuilder.Append "	AND cost_center ='상주직접비' "
				objBuilder.Append "	AND company ='"&cw_company&"' "
				objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
				objBuilder.Append "	AND group_name ='"&group_name&"' "
				objBuilder.Append "	AND cost_id ='인건비' "
				objBuilder.Append "	AND cost_detail ='협업' "
				objBuilder.Append "	AND saupbu ='"&cw_saupbu&"' "
			End If
			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
			rsCoworkCost.Close()
		'End If
	Next
	Set rsCworkTrade = Nothing
	Set rsCoworkCost = Nothing
End If

'협업 END	=============================================

' 사업부 별 초기화
'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='상주직접비' or cost_center ='직접비') "
objBuilder.Append "UPDATE saupbu_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&" = '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND (cost_center ='상주직접비' OR cost_center ='직접비')"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 상주직접비 와 직접비 업데이트
'sql = "select saupbu,cost_center,cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '상주직접비' or cost_center = '직접비') and cost_year ='"&cost_year&"' group by saupbu,cost_center,cost_id,cost_detail"
objBuilder.Append "SELECT saupbu, cost_center, cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE (cost_center = '상주직접비' OR cost_center = '직접비') "
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
' 사업부별 손익 자료 생성 종료

' 회사별별 손익 자료 생성
' 처리전 zero
'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='상주직접비') "
objBuilder.Append "UPDATE company_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&"= '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_center ='상주직접비' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 상주직접비 업데이트
'sql = "select company,group_name,cost_center,cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '상주직접비') and cost_year ='"&cost_year&"' group by company,group_name,cost_center,cost_id,cost_detail"
objBuilder.Append "SELECT company, group_name, cost_center, cost_id, cost_detail, SUM(cost_amt_"&cost_month&") AS cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_center = '상주직접비' "
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