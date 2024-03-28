<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim sum_amt(20, 3, 13)
Dim saupbu_tab(20)

Dim cost_year, base_year, view_sw, be_year
Dim title_line, savefilename, i, j, k
Dim rsSalesDept, arrSalesDept, rsCostStats, rsSaleStats
Dim rsKsysPart, ksysPartCost
Dim exceptDate

cost_year = f_Request("cost_year")	'조회 년도

title_line = cost_year & "년" & " 사업부별 손익 총괄 현황"
savefilename = title_line & ".xls"

'엑셀 다운로드 설정
Call ViewExcelType(savefilename)

If cost_year = "" Then
	cost_year = Mid(CStr(Now()), 1, 4)
	base_year = cost_year
	view_sw = "0"
End If

be_year = Int(cost_year) - 1

For i = 1 To 20
	saupbu_tab(i) = ""
Next

For i = 1 To 20
	For j = 1 To 3
		For k = 1 To 13
			sum_amt(i, j, k) = 0
		Next
	Next
Next

' 영업조직 발췌
objBuilder.Append "SELECT saupbu FROM sales_org "
objBuilder.Append "WHERE sales_year='" & cost_year & "' AND sort_seq <> '31' "	'OA수행본부는 제외

If team <> "회계재무" And user_id <> "102592" Then
	objBuilder.Append "	AND saupbu <> '기타사업부' "
End If

'보안 사항으로 소속 부서 제한 열람 조건 추가
If empProfitGrade = "N" Then
	objBuilder.Append "	AND saupbu = '"&bonbu&"' "
End If

objBuilder.Append "ORDER BY sort_seq ASC "

Set rsSalesDept = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsSalesDept.EOF Then
	arrSalesDept = rsSalesDept.getRows()
End If
rsSalesDept.Close() : Set rsSalesDept = Nothing

If IsArray(arrSalesDept) Then
	For i = LBound(arrSalesDept) To UBound(arrSalesDept, 2)
		saupbu_tab(i + 1) = arrSalesDept(0, i)
	Next
End If

'---------------------------------------------------------------------------------------------------------------
'// 2017-09-15 회계재무 팀만 기타사업부,회사간거래 조회 가능하게 수정
'---------------------------------------------------------------------------------------------------------------
If team="회계재무" Or user_id = "102592"  Then
	'i = i + 1
	'saupbu_tab(i) = "기타사업부"
	'i = i + 1
	'saupbu_tab(i) = "회사간거래"

	' 회사간거래
	'sql = "select cost_center,sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from company_cost where cost_year = '"&cost_year&"' and (cost_center = '회사간거래') group by cost_center"
	'rs.Open sql, Dbconn, 1
	'do until rs.eof
	'	for k = 1 to 12
	'		sum_amt(i,2,k) = sum_amt(i,2,k) + cdbl(rs(k))
	'	next
	'	rs.movenext()
	'loop
	'rs.close()

	objBuilder.Append "SELECT cost_center, SUM(cost_amt_01), SUM(cost_amt_02), "
	objBuilder.Append "	SUM(cost_amt_03), SUM(cost_amt_04), SUM(cost_amt_05), "
	objBuilder.Append "	SUM(cost_amt_06), SUM(cost_amt_07), SUM(cost_amt_08), "
	objBuilder.Append "	SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "
	objBuilder.Append "	SUM(cost_amt_12) "
	objBuilder.Append "FROM company_cost "
	objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
	objBuilder.Append "	AND cost_center = '회사간거래' "
	objBuilder.Append "GROUP BY cost_center "

	Set rsCostStats = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Do Until rsCostStats.EOF
		For k = 1 To 12
			sum_amt(i, 2, k) = sum_amt(i, 2, k) + CDbl(rsCostStats(k))
		Next

		rsCostStats.MoveNext()
	Loop
	rsCostStats.close() : Set rsCostStats = Nothing
End If
'---------------------------------------------------------------------------------------------------------------

' 매출 집계
'sql = "select substring(sales_date,1,7) as sales_month,saupbu,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,4) = '"&cost_year&"' group by substring(sales_date,1,7), saupbu"
'rs.Open sql, Dbconn, 1
'do until rs.eof
'	for i = 1 to 20
'		if saupbu_tab(i) = rs("saupbu") then
'			j = 1
'			k = int(mid(rs("sales_month"),6,2))
'			sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs("cost"))
'			exit for
'		end if
'	next
'	rs.movenext()
'loop
'rs.close()

objBuilder.Append "SELECT SUBSTRING(sales_date, 1, 7) AS sales_month, "
objBuilder.Append "	saupbu,	SUM(cost_amt) AS cost  "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(sales_date,1, 4) = '"&cost_year&"' "
objBuilder.Append "GROUP BY SUBSTRING(sales_date, 1, 7), saupbu "

Set rsSaleStats = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsSaleStats.EOF
	For i = 1 To 20
		If saupbu_tab(i) = rsSaleStats("saupbu") Then
			j = 1
			k = Int(Mid(rsSaleStats("sales_month"), 6, 2))

			sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsSaleStats("cost"))

			Exit For
		End If
	Next

	rsSaleStats.MoveNext()
Loop

rsSaleStats.Close() : Set rsSaleStats = Nothing

'202204월부터 전사공통비 SI1본부 고객사 삼성생명보험(주) 매출 제외 처리(재무 요청)[허정호_20220511]
exceptDate = "202204"

' 비용 집계
'sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' group by saupbu"
'rs.Open sql, Dbconn, 1

'do until rs.eof
'	for i = 1 to 20
'		if saupbu_tab(i) = rs("saupbu") then
'			j = 2
'			for k = 1 to 12
'				sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs(k))
'			next
'			exit for
'		end if
'	next
'	rs.movenext()
'loop
'rs.close()
Dim rsProfitStats, mm, end_month, rsManage, manageCost, rsPart, part_tot_cost
Dim as_tot_cnt, rsSaupbuPart, part_cnt, partCost

objBuilder.Append "SELECT saupbu, SUM(cost_amt_01), SUM(cost_amt_02), "
objBuilder.Append "	SUM(cost_amt_03), SUM(cost_amt_04), SUM(cost_amt_05), "

'분기별 비용 차이 차감(6,12월 실제에만 적용)
'objBuilder.Append "	SUM(cost_amt_06), "
objBuilder.Append "	(SUM(cost_amt_06) "
objBuilder.Append "	- (SELECT SUM(cost_amt_06) FROM saupbu_profit_loss "
objBuilder.Append "		WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비' AND cost_id = '일반경비' AND cost_detail = '급여' "
objBuilder.Append "		AND saupbu = splt.saupbu)), "

objBuilder.Append "	SUM(cost_amt_07), SUM(cost_amt_08), "
objBuilder.Append "	SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "

'objBuilder.Append "	SUM(cost_amt_12) "
objBuilder.Append "	(SUM(cost_amt_12) "
objBuilder.Append "	- (SELECT SUM(cost_amt_12) FROM saupbu_profit_loss "
objBuilder.Append "		WHERE cost_year = '"&cost_year&"' AND cost_center = '직접비' AND cost_id = '일반경비' AND cost_detail = '급여' "
objBuilder.Append "		AND saupbu = splt.saupbu)) "

objBuilder.Append "FROM saupbu_profit_loss AS splt "
objBuilder.Append "WHERE cost_year = '"&cost_year&"' "

'보안 사항으로 소속 부서 제한 열람 조건 추가
If empProfitGrade = "Y" Then
	objBuilder.Append "	AND saupbu IN (SELECT saupbu FROM sales_org WHERE sales_year = '"&cost_year&"' AND sort_seq <> '9') "
Else
	objBuilder.Append "	AND saupbu = '"&bonbu&"' "
End If

objBuilder.Append "	AND cost_center NOT IN ('전사공통비', '부문공통비', '부문공통비(2)') "
objBuilder.Append "GROUP BY saupbu "

Set rsProfitStats = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsProfitStats.EOF
	For i = 1 To 20
		If saupbu_tab(i) = rsProfitStats("saupbu") Then
			j = 2

			For k = 1 To 12
				If CInt(k) < 10 Then
					mm = "0" & k
				Else
					mm = k
				End If

				end_month = cost_year & mm

				'전사공통비
				objBuilder.Append "SELECT ROUND((tot_cost_amt * 0.5 / tot_person * saupbu_person) "
				objBuilder.Append "	+ (tot_cost_amt * 0.5 / tot_sale * saupbu_sale), 1) AS tot_amt "
				objBuilder.Append "FROM ( "
				objBuilder.Append "	SELECT mgct.saupbu, mgct.tot_cost_amt, mgct.saupbu_person, mgct.tot_person, "
				objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
				objBuilder.Append "		FROM saupbu_sales "
				objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' "
				objBuilder.Append "			AND mgct.saupbu = saupbu "

				If end_month >= exceptDate Then
					objBuilder.Append "		AND company <> '삼성생명보험(주)' "
				End If

				objBuilder.Append "		) AS saupbu_sale, "
				objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
				objBuilder.Append "		FROM saupbu_sales "
				objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' "
				objBuilder.Append "			AND saupbu <> '기타사업부' "

				If end_month >= exceptDate Then
					objBuilder.Append "		AND company <> '삼성생명보험(주)' "
				End If

				objBuilder.Append "		) AS tot_sale "
				objBuilder.Append "	FROM management_cost AS mgct "
				objBuilder.Append "	WHERE cost_month = '"&end_month&"' "
				objBuilder.Append "		AND saupbu = '"&saupbu_tab(i)&"' "
				objBuilder.Append "	GROUP BY saupbu"
				objBuilder.Append ") r1 "

				Set rsManage = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				If Not (rsManage.BOF Or rsManage.EOF) Then
					manageCost = rsManage("tot_amt")
				Else
					manageCost = 0
				End If
				rsManage.Close()

				'부문공통비
				'objBuilder.Append "SELECT IFNULL(SUM(cost_amt), 0) AS tot_amt "
				'objBuilder.Append "FROM company_asunit "
				'objBuilder.Append "WHERE as_month = '"&end_month&"' "
				'objBuilder.Append "	AND saupbu = '"&saupbu_tab(i)&"' "

				'Set rsPart = DBConn.Execute(objBuilder.ToString())
				'objBuilder.Clear()

				'If Not (rsPart.BOF Or rsPart.EOF) Then
				'	partCost = rsPart("tot_amt")
				'Else
				'	partCost = 0
				'End If
				'rsPart.Close()

				'부문공통비(배분)
				objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
				objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
				objBuilder.Append "	AND cost_detail = '설치공사')) AS 'part_tot_cost', "
				objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&end_month&"') AS 'as_tot_cnt' "
				objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '부문공통비' "

				Set rsPart = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				part_tot_cost = CDbl(f_toString(rsPart("part_tot_cost"), 0))	'부문공통비(배분)
				as_tot_cnt = CInt(f_toString(rsPart("as_tot_cnt"), 0))	'AS 총 건수

				rsPart.Close() : Set rsPart = Nothing

				'사업부 별 AS 총 건수 조회
				objBuilder.Append "SELECT SUM(as_total - as_set) AS as_cnt "
				objBuilder.Append "FROM as_acpt_status AS aast "
				objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
				objBuilder.Append "	AND trdt.trade_id = '매출' "
				objBuilder.Append "WHERE as_month = '"&end_month&"' "
				objBuilder.Append "	AND trdt.saupbu = '"&saupbu_tab(i)&"' "

				Set rsSaupbuPart = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				part_cnt = CInt(f_toString(rsSaupbuPart(0), 0))	'사업부 AS 총 건수

				rsSaupbuPart.Close() : Set rsSaupbuPart = Nothing

				'사업부별 배분 부분공통비
				If part_cnt > 0 Then
					partCost = part_tot_cost / as_tot_cnt * part_cnt
				Else
					partCost = 0
				End If

				'사업부별 배분 부문공통비(2)
				objBuilder.Append "SELECT ROUND((part_tot * 0.5 / tot_person * saupbu_person) + (part_tot * 0.5 / tot_sale * saupbu_sale), 1) FROM ("
				objBuilder.Append "	SELECT mgct.saupbu, mgct.saupbu_person, "
				objBuilder.Append "		(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '부문공통비(2)') AS 'part_tot',"
				objBuilder.Append "		(SELECT count(*) FROM pay_month_give AS pmgt "
				objBuilder.Append "		INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no AND emp_month = '"&end_month&"' "
				objBuilder.Append "		WHERE pmg_yymm = '"&end_month&"' AND pmgt.mg_saupbu IN ('금융SI본부', '공공SI본부', 'DI사업부문') "
				objBuilder.Append "			AND pmg_id = '1' AND pmg_emp_type = '정직' AND emmt.cost_except IN ('0', '1')) AS tot_person, "
				objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt FROM saupbu_sales "
				objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' AND mgct.saupbu = saupbu) AS saupbu_sale, "
				objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt FROM saupbu_sales "
				objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' AND saupbu IN ('금융SI본부', '공공SI본부', 'DI사업부문')) AS tot_sale"
				objBuilder.Append "	FROM management_cost AS mgct "
				objBuilder.Append "	WHERE cost_month = '"&end_month&"' AND saupbu IN ('금융SI본부', '공공SI본부', 'DI사업부문') "
				objBuilder.Append "	GROUP BY saupbu "
				objBuilder.Append ") r1 WHERE r1.saupbu= '"&saupbu_tab(i)&"' "

				Set rsKsysPart = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				If rsKsysPart.EOF Or rsKsysPart.BOF Then
					ksysPartCost = 0
				Else
					ksysPartCost = f_toString(rsKsysPart(0), 0)
				End If
				rsKsysPart.Close()

				'sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsProfitStats(k)) + CDbl(manageCost) + CDbl(partCost) + CDbl(ksysPartCost)
				sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(f_toString(rsProfitStats(k), 0)) + CDbl(manageCost) + CDbl(partCost) + CDbl(ksysPartCost)
			Next

			Exit For
		End If
	Next

	rsProfitStats.MoveNext()
Loop
Set rsManage = Nothing
Set rsPart = Nothing
rsProfitStats.Close() : Set rsProfitStats = Nothing

' 비용 집계 (기타사업부)
Dim rsEtcStats, cost_saupbu, rsPartEtc, part_etc_tot_cost, as_etc_tot_cnt,  rsSaupbuPartEtc
Dim part_etc_cnt, partEtcCost

'sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' and saupbu = '' group by saupbu"
'rs.Open sql, Dbconn, 1
'do until rs.eof
'	for i = 1 to 20
'		if saupbu_tab(i) = "기타사업부" then
'			j = 2
'			for k = 1 to 12
'				sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs(k))
'			next
'			exit for
'		end if
'	next
'	rs.movenext()
'loop
'rs.close()

' 비용이 없으면 매출도 표기 하지 않음
'for i = 1 to 20
'	if saupbu_tab(i) = "" then
'		exit for
'	end if
'	for k = 1 to 12
'		if sum_amt(i,2,k) = 0 then
'			sum_amt(i,1,k) = 0
'		end if
'	next
'next
objBuilder.Append "SELECT saupbu, SUM(cost1), SUM(cost2), SUM(cost3), SUM(cost4), SUM(cost5), "
objBuilder.Append "	SUM(cost6), SUM(cost7), SUM(cost8), SUM(cost9), SUM(cost10), SUM(cost11), SUM(cost12) "
objBuilder.Append "FROM( "
objBuilder.Append "	SELECT CASE WHEN saupbu = '' THEN '기타사업부' ELSE saupbu END AS saupbu, "
objBuilder.Append "		SUM(cost_amt_01) AS cost1, SUM(cost_amt_02) AS cost2, "
objBuilder.Append "		SUM(cost_amt_03) AS cost3, SUM(cost_amt_04) AS cost4, SUM(cost_amt_05) AS cost5, "
objBuilder.Append "		SUM(cost_amt_06) AS cost6, SUM(cost_amt_07) AS cost7, SUM(cost_amt_08) AS cost8, "
objBuilder.Append "		SUM(cost_amt_09) AS cost9, SUM(cost_amt_10) AS cost10, SUM(cost_amt_11) AS cost11, "
objBuilder.Append "		SUM(cost_amt_12) AS cost12 "
objBuilder.Append "	FROM saupbu_profit_loss "
objBuilder.Append "	WHERE cost_year = '"&cost_year&"' "
objBuilder.Append "		AND (saupbu = '' OR saupbu = '기타사업부') "
objBuilder.Append "		AND cost_center NOT IN ('전사공통비', '부문공통비', '부문공통비(2)') "
objBuilder.Append "	GROUP BY saupbu "
objBuilder.Append ") r1 "
objBuilder.Append "GROUP BY r1.saupbu "

Set rsEtcStats = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsEtcStats.EOF
	cost_saupbu = Trim(rsEtcStats("saupbu")&"")

	If cost_saupbu = "" Then
		cost_saupbu = "기타사업부"
	End If

	For i = 1 To 20
		If saupbu_tab(i) = cost_saupbu Then
			j = 2

			For k = 1 To 12

				If CInt(k) < 10 Then
					mm = "0" & k
				Else
					mm = k
				End If

				end_month = cost_year & mm

				'부문공통비(기타사업부)
				'objBuilder.Append "SELECT IFNULL(SUM(cost_amt), 0) AS tot_amt "
				'objBuilder.Append "FROM company_asunit "
				'objBuilder.Append "WHERE as_month = '"&end_month&"' "
				'objBuilder.Append "	AND saupbu = '"&rsEtcStats("saupbu")&"' "

				'Set rsPartEtc = DBConn.Execute(objBuilder.ToString())
				'objBuilder.Clear()

				'If Not (rsPartEtc.BOF Or rsPartEtc.EOF) Then
				'	partEtcCost = rsPartEtc("tot_amt")
				'Else
				'	partEtcCost = 0
				'End If
				'rsPartEtc.Close()

				'부문공통비(배분)
				objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
				objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
				objBuilder.Append "	AND cost_detail = '설치공사')) AS 'part_tot_cost', "
				objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&end_month&"') AS 'as_tot_cnt' "
				objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '부문공통비' "

				Set rsPartEtc = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				part_etc_tot_cost = CDbl(f_toString(rsPartEtc("part_tot_cost"), 0))	'부문공통비(배분)
				as_etc_tot_cnt = CInt(f_toString(rsPartEtc("as_tot_cnt"), 0))	'AS 총 건수

				rsPartEtc.Close() : Set rsPartEtc = Nothing

				'사업부 별 AS 총 건수 조회
				objBuilder.Append "SELECT SUM(as_total - as_set) AS as_cnt "
				objBuilder.Append "FROM as_acpt_status AS aast "
				objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
				objBuilder.Append "	AND trdt.trade_id = '매출' "
				objBuilder.Append "WHERE as_month = '"&end_month&"' "
				objBuilder.Append "	AND trdt.saupbu = '' "

				Set rsSaupbuPartEtc = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				part_etc_cnt = CInt(f_toString(rsSaupbuPartEtc(0), 0))	'사업부 AS 총 건수

				rsSaupbuPartEtc.Close() : Set rsSaupbuPartEtc = Nothing

				'사업부별 배분 부분공통비
				If part_etc_cnt > 0 Then
					partEtcCost = part_etc_tot_cost / as_etc_tot_cnt * part_etc_cnt
				Else
					partEtcCost = 0
				End If

				sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsEtcStats(k)) + CDbl(partEtcCost)
			Next

			Exit For
		End If
	Next

	rsEtcStats.MoveNext()
Loop
Set rsPartEtc = Nothing
rsEtcStats.Close() : Set rsEtcStats = Nothing

' 손익계산
For i = 1 To 20
	If saupbu_tab(i) = "" Then
		Exit For
	End If

	j = 3
	For k = 1 To 12
		sum_amt(i, j, k) = sum_amt(i, 1, k) - sum_amt(i, 2, k)
	Next
Next

' 년 합계
For i = 1 To 20
	If saupbu_tab(i) = "" Then
		Exit For
	End If

	For j = 1 To 3
		For k = 1 To 12
			sum_amt(i, j, 13) = sum_amt(i, j, 13) + sum_amt(i, j, k)
		Next
	Next
Next

' 총계
For i = 1 To 20
	If saupbu_tab(i) = "" Then
		Exit For
	End If

	For j = 1 To 3
		For k = 1 To 13
			sum_amt(0, j, k) = sum_amt(0, j, k) + sum_amt(i, j, k)
		Next
	Next
Next
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
                <div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
						</colgroup>
						<thead>
							<tr>
							  <th class="first" scope="col">사업부</th>
							  <th scope="col">구분</th>
						<% For i = 1 To  12	%>
							  <th scope="col"><%=i%>월</th>
						<% Next	%>
							  <th scope="col">합계</th>
                          </tr>
						</thead>
						<tbody>
					<%
						For i = 1 To 20
							If saupbu_tab(i) = "" Then
								Exit For
							End If
					%>
							<tr>
							  	<td rowspan="3" class="first"><%=saupbu_tab(i)%></td>
								<td>매출</td>
						<%
							For k = 1 To 13
						%>
								<td class="right"><%=FormatNumber(sum_amt(i, 1, k), 0)%></td>
						<%
							Next
						%>
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">비용</td>
						<%
							For k = 1 To 13
						%>
								<td class="right">
								<%=FormatNumber(sum_amt(i, 2, k), 0)%>
                                </td>
						<%
							next
						%>
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">손익</td>
						<%
							For k = 1 To 13
						%>
								<td class="right"><%=FormatNumber(sum_amt(i, 3, k), 0)%></td>
						<%
							Next
						%>
			              </tr>
					<%
						Next
					%>
							<tr>
							  	<td rowspan="3" class="first" bgcolor="#CCFFFF"><strong>계</strong></td>
								<td>매출</td>
						<%
							For k = 1 To 13
						%>
								<td class="right"><%=FormatNumber(sum_amt(0, 1, k), 0)%></td>
						<%
							Next
						%>
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">비용</td>
						<%
							For k = 1 To 13
						%>
								<td class="right"><%=FormatNumber(sum_amt(0, 2, k), 0)%></td>
						<%
							Next
						%>
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">손익</td>
						<%
							For k = 1 To 13
						%>
								<td class="right"><%=FormatNumber(sum_amt(0, 3, k), 0)%></td>
						<%
							Next
						%>
			              </tr>
						</tbody>
					</table>
				<br>
			</form>
		</div>
	</div>
	</body>
</html>
<!--#include virtual="/common/log_sales_profit.asp" -->
<%
DBConn.Close() : Set DBConn = Nothing
%>