<%
' 원격 5%, 방문 95% -> 원격 0, 방문 100% 조정
'won_per = 5
'bang_per = 95
won_per = 0
bang_per = 100

'sql = "select sum(cost_amt_"&mm&") as tot_cost from company_cost where cost_year ='"&cost_year&"' and cost_center = '부문공통비'"
objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS tot_cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_center = '부문공통비' "

Set rsCostAmtTot = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'총 부문 공통비
tot_cost = CLng(rsCostAmtTot("tot_cost"))

rsCostAmtTot.Close() : Set rsCostAmtTot = Nothing

'==========================================================
'AS 현황 및 부문 공통비 배분 프로세스 추가[허정호_20210503]
'==========================================================

'해당 월 AS 데이타 복사(AS 현황 용) - 부분 복사되는 것으로 확인되어 로직 제외 [허정호_20210505]
'objBuilder.Append "INSERT INTO as_acpt_end "
'objBuilder.Append "SELECT * FROM as_acpt "
'objBuilder.Append "WHERE REPLACE(SUBSTRING(acpt_date, 1, 7), '-', '') = '"&end_month&"' "
'objBuilder.Append "ORDER BY acpt_no ASC "

'DBConn.Execute(objBuilder.ToString())
'objBuilder.Clear()

'Dim rsSaupbuSalesTotal, saupbuSalesTotal

'4개 사업부 총 매출액(si1본부, si2본부, NI본부, 공공본부)
'objBuilder.Append "SELECT SUM(cost_amt) AS tot_sale "
'objBuilder.Append "FROM saupbu_sales "
'objBuilder.Append "WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' "
'objBuilder.Append "	AND saupbu IN ('SI1본부', 'SI2본부', 'NI본부', '공공본부') "

'Set rsSaupbuSalesTotal = DBConn.Execute(objBuilder.ToString())
'objBuilder.Clear()

'4개 사업부 총 매출액
'saupbuSalesTotal = CDbl(rsSaupbuSalesTotal("tot_sale"))

'rsSaupbuSalesTotal.Close() : Set rsSaupbuSalesTotal = Nothing

'objBuilder.Append "SELECT company AS as_company, as_saupbu AS saupbu, as_cnt, "
'objBuilder.Append "	ROUND(std_cost, 3) AS divide_amt_1, /*1차배부금액*/"
'objBuilder.Append "	ROUND(saupbu_sales / "&saupbuSalesTotal&" * ("&tot_cost&" - std_cost), 3) AS divide_amt_2, /*2차배부금액*/"
'objBuilder.Append "	ROUND((std_cost + (saupbu_sales / "&saupbuSalesTotal&" * ("&tot_cost&" - std_cost))) / "&tot_cost&", 3) AS charge_per, /*차지율*/"
'objBuilder.Append "	ROUND((std_cost + (saupbu_sales / "&saupbuSalesTotal&" * ("&tot_cost&" - std_cost))), 3) AS cost_amt /*부문공통비*/"
'objBuilder.Append "FROM ( "
'objBuilder.Append "	SELECT asat.company, trat.saupbu AS as_saupbu, "
'objBuilder.Append "		COUNT(*) AS as_cnt, "
'objBuilder.Append "		SUM(asat.as_standard_money) AS std_cost, "
'objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) FROM saupbu_sales "
'objBuilder.Append "		WHERE company = asat.company AND saupbu = trat.saupbu "
'objBuilder.Append "			AND REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"') AS saupbu_sales "
'objBuilder.Append "	FROM as_acpt AS asat "
'objBuilder.Append "	INNER JOIN emp_master_month AS emmt ON asat.mg_ce_id = emmt.emp_no "
'objBuilder.Append "		AND emmt.emp_month = '"&end_month&"' "
'objBuilder.Append "	LEFT OUTER JOIN trade AS trat ON asat.company = trat.trade_name "
'objBuilder.Append "	WHERE asat.as_type NOT IN ('원격처리', '야특근') "
'objBuilder.Append "		AND asat.as_process <> '취소' "
'objBuilder.Append "		AND asat.reside = '0' "
'objBuilder.Append "		AND asat.reside_place = '' "
'objBuilder.Append "		AND (CAST(asat.visit_date AS DATE) >= '"&from_date&"' AND CAST(asat.visit_date AS DATE) <= '"&to_date&"') "
'objBuilder.Append "		AND emmt.cost_center = '부문공통비' "
'objBuilder.Append "	GROUP BY asat.company "
'objBuilder.Append ") r1 "

Dim rsAsTot, tot_part_cnt, rsCompanyAs
'A/S 전체 카운트
objBuilder.Append "SELECT COUNT(*) AS tot_cnt "
objBuilder.Append "FROM as_acpt AS asat "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON asat.mg_ce_id = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&end_month&"'"
objBuilder.Append "INNER JOIN trade AS trat ON asat.company = trat.trade_name "
objBuilder.Append "WHERE asat.as_type NOT IN ('원격처리', '야특근')"
objBuilder.Append "	AND asat.as_process <> '취소'"
objBuilder.Append "	AND asat.reside = '0'"
objBuilder.Append "	AND asat.reside_place = ''"
objBuilder.Append "	AND (CAST(asat.visit_date AS DATE) >= '"&from_date&"' AND CAST(asat.visit_date AS DATE) <= '"&to_date&"') "
objBuilder.Append "	AND emmt.cost_center = '부문공통비' "

Set rsAsTot = DBconn.Execute(objBuilder.ToString())
objBuilder.Clear()

tot_part_cnt = rsAsTot("tot_cnt")

rsAsTot.Close() : Set rsAsTot = Nothing 

objBuilder.Append "SELECT company, bonbu, cnt, "
'objBuilder.Append "	SUM(IF(as_type = '기타' OR as_type = '방문처리', cnt, 0)) AS 'fault', "
'objBuilder.Append "	SUM(IF(as_type = '신규설치' OR as_type = '신규설치공사' OR as_type = '이전설치' "
'objBuilder.Append "		OR as_type = '이전설치공사' OR as_type = '랜공사' OR as_type = '이전랜공사', cnt, 0)) AS 'setting', "
'objBuilder.Append "	SUM(IF(as_type = '예방점검', cnt, 0)) AS 'testing', "
'objBuilder.Append "	SUM(IF(as_type = '장비회수', cnt, 0)) AS 'collect', "
objBuilder.Append	tot_cost&" / "&tot_part_cnt&" * SUM(cnt) AS 'as_cost' "	'/*부문공통비 전체 비용 / as 전체 건수 * 사이트별 AS 건수*/
objBuilder.Append "FROM ( "
objBuilder.Append "	SELECT asat.company, trat.saupbu AS bonbu, COUNT(*) AS cnt, SUM(as_standard_money) AS std_cost "
objBuilder.Append "	FROM as_acpt AS asat "
objBuilder.Append "	INNER JOIN emp_master_month AS emmt ON asat.mg_ce_id = emmt.emp_no "
objBuilder.Append "		AND emmt.emp_month = '"&end_month&"' "
objBuilder.Append "	INNER JOIN trade AS trat ON asat.company = trat.trade_name "
objBuilder.Append "	WHERE asat.as_type NOT IN ('원격처리', '야특근') "
objBuilder.Append "		AND asat.as_process <> '취소' "
objBuilder.Append "		AND asat.reside = '0' "
objBuilder.Append "		AND asat.reside_place = '' "
objBuilder.Append "		AND (CAST(asat.visit_date as date) >= '"&from_date&"' AND CAST(asat.visit_date as date) <= '"&to_date&"') "
objBuilder.Append "		AND emmt.cost_center = '부문공통비' "
objBuilder.Append "	GROUP BY asat.company, as_type "
objBuilder.Append ") r1 "
objBuilder.Append "GROUP BY company "

Set rsCompanyAs = Server.CreateObject("ADODB.RecordSet")
rsCompanyAs.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCompanyAs.EOF
	objBuilder.Append "INSERT INTO company_asunit(as_month, as_company, saupbu, as_cnt, divide_amt_1, divide_amt_2, charge_per, "
	objBuilder.Append "cost_amt, reg_id, reg_name, reg_date)VALUES("
	objBuilder.Append "'"&end_month&"', '"&rsCompanyAs("company")&"', '"&rsCompanyAs("bonbu")&"', '"&rsCompanyAs("cnt")&"', 0, 0, 0, "
	objBuilder.Append "'"&rsCompanyAs("as_cost")&"', '"&user_id&"', '"&user_name&"', NOW()"
	objBuilder.Append ")"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsCompanyAs.MoveNext()
Loop

rsCompanyAs.Close() : Set rsCompanyAs = Nothing

'AS 현황 및 부문 공통비 배분 프로세스 END

'원격
'sql = " select count(*) as tot_cnt from as_acpt Where acpt_man in ('조민순','주영미','한수정','안태환') and (Cast(visit_date as date) >= '" + from_date + "' and Cast(visit_date as date) <= '"+to_date+"')and company not in('코웨이','웅진씽크빅','웅진식품','롯데렌탈')"
objBuilder.Append "SELECT COUNT(*) AS tot_cnt "
objBuilder.Append "FROM as_acpt "
objBuilder.Append "WHERE acpt_man IN ('조민순', '주영미', '한수정', '안태환') "
objBuilder.Append "	AND (CAST(visit_date AS date) >= '" & from_date & "' "
objBuilder.Append "		AND CAST(visit_date AS date) <= '" & to_date & "') "
objBuilder.Append "	AND company NOT IN ('코웨이','웅진씽크빅','웅진식품','롯데렌탈') "

Set rsAsCnt = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

won_cnt = CLng(rsAsCnt("tot_cnt"))

If won_cnt = "" Or IsNull(won_cnt) Then
	won_cnt = 0
End If

rsAsCnt.Close() : Set rsAsCnt = Nothing

'sql = "select company, count(*) as as_cnt from as_acpt Where acpt_man in ('조민순','주영미','한수정','안태환') and (Cast(visit_date as date) >= '" + from_date + "' and Cast(visit_date as date) <= '"+to_date+"') and company not in('코웨이','웅진씽크빅','웅진식품','롯데렌탈') GROUP BY company Order By company Asc"
objBuilder.Append "SELECT company, COUNT(*) AS as_cnt "
objBuilder.Append "FROM as_acpt "
objBuilder.Append "WHERE acpt_man IN ('조민순', '주영미', '한수정', '안태환') "
objBuilder.Append "	AND (CAST(visit_date AS date) >= '" & from_date & "' "
objBuilder.Append "		AND CAST(visit_date AS date) <= '"&to_date&"') "
objBuilder.Append "	AND company NOT IN ('코웨이', '웅진씽크빅', '웅진식품', '롯데렌탈') "
objBuilder.Append "GROUP BY company "
objBuilder.Append "Order By company Asc "

Set rsRemoteCnt = Server.CreateObject("ADODB.RecordSet")
rsRemoteCnt.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsRemoteCnt.EOF
	'sql = "select saupbu from trade where trade_name = '"&rs("company")&"'"
	objBuilder.Append "SELECT saupbu "
	objBuilder.Append "FROM trade "
	objBuilder.Append "WHERE trade_name = '"&rsRemoteCnt("company")&"' "

	Set rsRemoteTrade = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsRemoteTrade.EOF Or rsRemoteTrade.BOF Then
		trade_bonbu = "Error"
	Else
		trade_bonbu = rsRemoteTrade("saupbu")
	End If

	rsRemoteTrade.Close()

	charge_per = CLng(rsRemoteCnt("as_cnt")) / won_cnt * won_per / 100
	cost_amt = Int(charge_per * tot_cost)

	'sql="insert into company_as (as_month,as_company,saupbu,remote_cnt,charge_per,cost_amt,reg_id,reg_name,reg_date) values ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&charge_per&"',"&cost_amt&",'"&user_id&"','"&user_name&"',now())"
	objBuilder.Append "INSERT INTO company_as(as_month, as_company, saupbu,"
	objBuilder.Append "remote_cnt, charge_per, cost_amt,"
	objBuilder.Append "reg_id, reg_name, reg_date)VALUES("
	objBuilder.Append "'"&end_month&"', '"&rsRemoteCnt("company")&"', '"&saupbu&"',"
	objBuilder.Append "'"&rsRemoteCnt("as_cnt")&"', '"&charge_per&"', "&cost_amt&","
	objBuilder.Append "'"&user_id&"', '"&user_name&"', NOW())"

	'DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsRemoteCnt.MoveNext()
Loop
rsRemoteCnt.Close() : Set rsRemoteCnt = Nothing

' 원격외
'sql = " select count(*) as tot_cnt "
'sql = sql & " from as_acpt a inner join emp_master_month b on a.mg_ce_id=b.emp_no and b.emp_month='" & end_month & "'"
'sql = sql & " Where (as_type <> '원격처리' and as_process <> '취소' and as_type <> '야특근') "
'sql = sql & " and reside='0'  and reside_place=' ' "
'sql = sql & " and (Cast(visit_date as date) >= '" + from_date + "' and Cast(visit_date as date) <= '"+to_date+"')"
'sql = sql & " and b.cost_center='부문공통비' "

objBuilder.Append "SELECT COUNT(*) AS tot_cnt "
objBuilder.Append "FROM as_acpt AS asat  "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON asat.mg_ce_id = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&end_month& "' "
objBuilder.Append "WHERE (asat.as_type <> '원격처리' AND asat.as_process <> '취소' "
objBuilder.Append "		AND asat.as_type <> '야특근') "
objBuilder.Append "	AND asat.reside = '0' "
objBuilder.Append "	AND asat.reside_place = ' ' "
objBuilder.Append "	AND (CAST(asat.visit_date AS date) >= '"& from_date&"' "
objBuilder.Append "		AND CAST(asat.visit_date AS date) <= '"&to_date&"') "
objBuilder.Append "	AND emmt.cost_center = '부문공통비' "

Set rsNoRemote = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

bang_cnt = CLng(rsNoRemote("tot_cnt"))

If bang_cnt = "" Or IsNull(bang_cnt) Then
	bang_cnt = 0
End If

rsNoRemote.Close() : Set rsNoRemote = Nothing

'sql = " select company, count(*) as as_cnt "
'sql = sql & " from as_acpt a inner join emp_master_month b on a.mg_ce_id=b.emp_no and b.emp_month='" & end_month & "'"
'sql = sql & " Where (as_type <> '원격처리' and as_process <> '취소' and as_type <> '야특근') "
'sql = sql & " and reside='0' and reside_place=' ' "
'sql = sql & " and (Cast(visit_date as date) >= '" + from_date + "' and Cast(visit_date as date) <= '"+to_date+"') "
'sql = sql & " and b.cost_center='부문공통비' "
'sql = sql & " GROUP BY company Order By company Asc"

objBuilder.Append "SELECT asat.company, COUNT(*) AS as_cnt "
objBuilder.Append "FROM as_acpt AS asat "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON asat.mg_ce_id = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '" & end_month & "' "
objBuilder.Append "WHERE (asat.as_type <> '원격처리' AND asat.as_process <> '취소' AND asat.as_type <> '야특근') "
objBuilder.Append "	AND asat.reside = '0' "
objBuilder.Append "	AND asat.reside_place = ' ' "
objBuilder.Append "	AND (CAST(asat.visit_date AS date) >= '"&from_date&"' AND CAST(asat.visit_date AS date) <= '"&to_date&"') "
objBuilder.Append "	AND emmt.cost_center = '부문공통비' "
objBuilder.Append "GROUP BY asat.company "
objBuilder.Append "ORDER BY asat.company ASC "

Set rsNoRemoteCnt = Server.CreateObject("ADODB.RecordSet")
rsNoRemoteCnt.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsNoRemoteCnt.EOF
	'sql = "select saupbu from trade where trade_name = '"&rs("company")&"'"
	objBuilder.Append "SELECT saupbu "
	objBuilder.Append "FROM trade "
	objBuilder.Append "WHERE trade_name = '"&rsNoRemoteCnt("company")&"' "

	Set rsNoRemoteTrade = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsNoRemoteTrade.EOF Or rsNoRemoteTrade.BOF Then
		trade_bonbu = "Error"
	Else
	  	trade_bonbu = rsNoRemoteTrade("saupbu")
	End If
	rsNoRemoteTrade.Close()

	'sql = "select * from company_as where as_month = '"&end_month&"' and as_company = '"&rs("company")&"'"
	objBuilder.Append "SELECT as_month, charge_per "
	objBuilder.Append "FROM company_as "
	objBuilder.Append "WHERE as_month = '"&end_month&"' "
	objBuilder.Append "	AND as_company = '"&rsNoRemoteCnt("company")&"' "

	Set rsCompAsEtc = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsCompAsEtc.EOF Or rsCompAsEtc.BOF Then
		'고객사별 건수 / 전체 건수 * AS 비율(원격, 방문) /100
		'charge_per = CLng(rsNoRemoteCnt("as_cnt")) / bang_cnt * bang_per / 100

		charge_per = CLng(rsNoRemoteCnt("as_cnt")) / bang_cnt
		cost_amt = Int(charge_per * tot_cost)

		'sql="INSERT INTO company_as (as_month,as_company,saupbu,visit_cnt,charge_per,cost_amt,reg_id,reg_name,reg_date) "&_
		'    " VALUES ('"&end_month&"','"&rs("company")&"','"&trade_bonbu&"','"&rs("as_cnt")&"','"&charge_per&"',"&cost_amt&",'"&user_id&"','"&user_name&"',now())"
		objBuilder.Append "INSERT INTO company_as(as_month, as_company, saupbu, "
		objBuilder.Append "visit_cnt, charge_per, cost_amt, "
		objBuilder.Append "reg_id, reg_name, reg_date)VALUES("
		objBuilder.Append "'"&end_month&"', '"&rsNoRemoteCnt("company")&"', '"&trade_bonbu&"', "
		objBuilder.Append "'"&rsNoRemoteCnt("as_cnt")&"', '"&charge_per&"', "&cost_amt&", "
		objBuilder.Append "'"&user_id&"', '"&user_name&"', NOW()) "
	Else
		charge_per = CLng(rsNoRemoteCnt("as_cnt")) / bang_cnt * bang_per / 100 + rsCompAsEtc("charge_per")
		cost_amt = Int(charge_per * tot_cost)

		'sql = "UPDATE company_as SET visit_cnt='"&rs("as_cnt")&"', charge_per='"&charge_per&"', cost_amt="&cost_amt&_
		      '" WHERE as_company='" &rs("company")& "' and as_month = '" &end_month& "'"
		objBuilder.Append "UPDATE company_as SET "
		objBuilder.Append "	visit_cnt='"&rsNoRemoteCnt("as_cnt")&"', "
		objBuilder.Append "	charge_per='"&charge_per&"', "
		objBuilder.Append "	cost_amt="&cost_amt&" "
		objBuilder.Append "WHERE as_company='"&rsNoRemoteCnt("company")&"' "
		objBuilder.Append "	AND as_month = '"&end_month&"' "
	End If
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
	rsCompAsEtc.Close()

	rsNoRemoteCnt.MoveNext()
Loop
rsNoRemoteCnt.Close() : Set rsNoRemoteCnt = Nothing



%>