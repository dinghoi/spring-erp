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

cost_year = f_Request("cost_year")	'��ȸ �⵵

title_line = cost_year & "��" & " ����κ� ���� �Ѱ� ��Ȳ"
savefilename = title_line & ".xls"

'���� �ٿ�ε� ����
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

' �������� ����
objBuilder.Append "SELECT saupbu FROM sales_org "
objBuilder.Append "WHERE sales_year='" & cost_year & "' AND sort_seq <> '31' "	'OA���ົ�δ� ����

If team <> "ȸ���繫" And user_id <> "102592" Then
	objBuilder.Append "	AND saupbu <> '��Ÿ�����' "
End If

'���� �������� �Ҽ� �μ� ���� ���� ���� �߰�
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
'// 2017-09-15 ȸ���繫 ���� ��Ÿ�����,ȸ�簣�ŷ� ��ȸ �����ϰ� ����
'---------------------------------------------------------------------------------------------------------------
If team="ȸ���繫" Or user_id = "102592"  Then
	'i = i + 1
	'saupbu_tab(i) = "��Ÿ�����"
	'i = i + 1
	'saupbu_tab(i) = "ȸ�簣�ŷ�"

	' ȸ�簣�ŷ�
	'sql = "select cost_center,sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from company_cost where cost_year = '"&cost_year&"' and (cost_center = 'ȸ�簣�ŷ�') group by cost_center"
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
	objBuilder.Append "	AND cost_center = 'ȸ�簣�ŷ�' "
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

' ���� ����
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

'202204������ �������� SI1���� ���� �Ｚ������(��) ���� ���� ó��(�繫 ��û)[����ȣ_20220511]
exceptDate = "202204"

' ��� ����
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

'�б⺰ ��� ���� ����(6,12�� �������� ����)
'objBuilder.Append "	SUM(cost_amt_06), "
objBuilder.Append "	(SUM(cost_amt_06) "
objBuilder.Append "	- (SELECT SUM(cost_amt_06) FROM saupbu_profit_loss "
objBuilder.Append "		WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' "
objBuilder.Append "		AND saupbu = splt.saupbu)), "

objBuilder.Append "	SUM(cost_amt_07), SUM(cost_amt_08), "
objBuilder.Append "	SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "

'objBuilder.Append "	SUM(cost_amt_12) "
objBuilder.Append "	(SUM(cost_amt_12) "
objBuilder.Append "	- (SELECT SUM(cost_amt_12) FROM saupbu_profit_loss "
objBuilder.Append "		WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' "
objBuilder.Append "		AND saupbu = splt.saupbu)) "

objBuilder.Append "FROM saupbu_profit_loss AS splt "
objBuilder.Append "WHERE cost_year = '"&cost_year&"' "

'���� �������� �Ҽ� �μ� ���� ���� ���� �߰�
If empProfitGrade = "Y" Then
	objBuilder.Append "	AND saupbu IN (SELECT saupbu FROM sales_org WHERE sales_year = '"&cost_year&"' AND sort_seq <> '9') "
Else
	objBuilder.Append "	AND saupbu = '"&bonbu&"' "
End If

objBuilder.Append "	AND cost_center NOT IN ('��������', '�ι������', '�ι������(2)') "
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

				'��������
				objBuilder.Append "SELECT ROUND((tot_cost_amt * 0.5 / tot_person * saupbu_person) "
				objBuilder.Append "	+ (tot_cost_amt * 0.5 / tot_sale * saupbu_sale), 1) AS tot_amt "
				objBuilder.Append "FROM ( "
				objBuilder.Append "	SELECT mgct.saupbu, mgct.tot_cost_amt, mgct.saupbu_person, mgct.tot_person, "
				objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
				objBuilder.Append "		FROM saupbu_sales "
				objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' "
				objBuilder.Append "			AND mgct.saupbu = saupbu "

				If end_month >= exceptDate Then
					objBuilder.Append "		AND company <> '�Ｚ������(��)' "
				End If

				objBuilder.Append "		) AS saupbu_sale, "
				objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
				objBuilder.Append "		FROM saupbu_sales "
				objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' "
				objBuilder.Append "			AND saupbu <> '��Ÿ�����' "

				If end_month >= exceptDate Then
					objBuilder.Append "		AND company <> '�Ｚ������(��)' "
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

				'�ι������
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

				'�ι������(���)
				objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
				objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
				objBuilder.Append "	AND cost_detail = '��ġ����')) AS 'part_tot_cost', "
				objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&end_month&"') AS 'as_tot_cnt' "
				objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '�ι������' "

				Set rsPart = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				part_tot_cost = CDbl(f_toString(rsPart("part_tot_cost"), 0))	'�ι������(���)
				as_tot_cnt = CInt(f_toString(rsPart("as_tot_cnt"), 0))	'AS �� �Ǽ�

				rsPart.Close() : Set rsPart = Nothing

				'����� �� AS �� �Ǽ� ��ȸ
				objBuilder.Append "SELECT SUM(as_total - as_set) AS as_cnt "
				objBuilder.Append "FROM as_acpt_status AS aast "
				objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
				objBuilder.Append "	AND trdt.trade_id = '����' "
				objBuilder.Append "WHERE as_month = '"&end_month&"' "
				objBuilder.Append "	AND trdt.saupbu = '"&saupbu_tab(i)&"' "

				Set rsSaupbuPart = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				part_cnt = CInt(f_toString(rsSaupbuPart(0), 0))	'����� AS �� �Ǽ�

				rsSaupbuPart.Close() : Set rsSaupbuPart = Nothing

				'����κ� ��� �κа����
				If part_cnt > 0 Then
					partCost = part_tot_cost / as_tot_cnt * part_cnt
				Else
					partCost = 0
				End If

				'����κ� ��� �ι������(2)
				objBuilder.Append "SELECT ROUND((part_tot * 0.5 / tot_person * saupbu_person) + (part_tot * 0.5 / tot_sale * saupbu_sale), 1) FROM ("
				objBuilder.Append "	SELECT mgct.saupbu, mgct.saupbu_person, "
				objBuilder.Append "		(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '�ι������(2)') AS 'part_tot',"
				objBuilder.Append "		(SELECT count(*) FROM pay_month_give AS pmgt "
				objBuilder.Append "		INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no AND emp_month = '"&end_month&"' "
				objBuilder.Append "		WHERE pmg_yymm = '"&end_month&"' AND pmgt.mg_saupbu IN ('����SI����', '����SI����', 'DI����ι�') "
				objBuilder.Append "			AND pmg_id = '1' AND pmg_emp_type = '����' AND emmt.cost_except IN ('0', '1')) AS tot_person, "
				objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt FROM saupbu_sales "
				objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' AND mgct.saupbu = saupbu) AS saupbu_sale, "
				objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt FROM saupbu_sales "
				objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&end_month&"' AND saupbu IN ('����SI����', '����SI����', 'DI����ι�')) AS tot_sale"
				objBuilder.Append "	FROM management_cost AS mgct "
				objBuilder.Append "	WHERE cost_month = '"&end_month&"' AND saupbu IN ('����SI����', '����SI����', 'DI����ι�') "
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

' ��� ���� (��Ÿ�����)
Dim rsEtcStats, cost_saupbu, rsPartEtc, part_etc_tot_cost, as_etc_tot_cnt,  rsSaupbuPartEtc
Dim part_etc_cnt, partEtcCost

'sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' and saupbu = '' group by saupbu"
'rs.Open sql, Dbconn, 1
'do until rs.eof
'	for i = 1 to 20
'		if saupbu_tab(i) = "��Ÿ�����" then
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

' ����� ������ ���⵵ ǥ�� ���� ����
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
objBuilder.Append "	SELECT CASE WHEN saupbu = '' THEN '��Ÿ�����' ELSE saupbu END AS saupbu, "
objBuilder.Append "		SUM(cost_amt_01) AS cost1, SUM(cost_amt_02) AS cost2, "
objBuilder.Append "		SUM(cost_amt_03) AS cost3, SUM(cost_amt_04) AS cost4, SUM(cost_amt_05) AS cost5, "
objBuilder.Append "		SUM(cost_amt_06) AS cost6, SUM(cost_amt_07) AS cost7, SUM(cost_amt_08) AS cost8, "
objBuilder.Append "		SUM(cost_amt_09) AS cost9, SUM(cost_amt_10) AS cost10, SUM(cost_amt_11) AS cost11, "
objBuilder.Append "		SUM(cost_amt_12) AS cost12 "
objBuilder.Append "	FROM saupbu_profit_loss "
objBuilder.Append "	WHERE cost_year = '"&cost_year&"' "
objBuilder.Append "		AND (saupbu = '' OR saupbu = '��Ÿ�����') "
objBuilder.Append "		AND cost_center NOT IN ('��������', '�ι������', '�ι������(2)') "
objBuilder.Append "	GROUP BY saupbu "
objBuilder.Append ") r1 "
objBuilder.Append "GROUP BY r1.saupbu "

Set rsEtcStats = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsEtcStats.EOF
	cost_saupbu = Trim(rsEtcStats("saupbu")&"")

	If cost_saupbu = "" Then
		cost_saupbu = "��Ÿ�����"
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

				'�ι������(��Ÿ�����)
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

				'�ι������(���)
				objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
				objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
				objBuilder.Append "	AND cost_detail = '��ġ����')) AS 'part_tot_cost', "
				objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&end_month&"') AS 'as_tot_cnt' "
				objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '�ι������' "

				Set rsPartEtc = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				part_etc_tot_cost = CDbl(f_toString(rsPartEtc("part_tot_cost"), 0))	'�ι������(���)
				as_etc_tot_cnt = CInt(f_toString(rsPartEtc("as_tot_cnt"), 0))	'AS �� �Ǽ�

				rsPartEtc.Close() : Set rsPartEtc = Nothing

				'����� �� AS �� �Ǽ� ��ȸ
				objBuilder.Append "SELECT SUM(as_total - as_set) AS as_cnt "
				objBuilder.Append "FROM as_acpt_status AS aast "
				objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
				objBuilder.Append "	AND trdt.trade_id = '����' "
				objBuilder.Append "WHERE as_month = '"&end_month&"' "
				objBuilder.Append "	AND trdt.saupbu = '' "

				Set rsSaupbuPartEtc = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				part_etc_cnt = CInt(f_toString(rsSaupbuPartEtc(0), 0))	'����� AS �� �Ǽ�

				rsSaupbuPartEtc.Close() : Set rsSaupbuPartEtc = Nothing

				'����κ� ��� �κа����
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

' ���Ͱ��
For i = 1 To 20
	If saupbu_tab(i) = "" Then
		Exit For
	End If

	j = 3
	For k = 1 To 12
		sum_amt(i, j, k) = sum_amt(i, 1, k) - sum_amt(i, 2, k)
	Next
Next

' �� �հ�
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

' �Ѱ�
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
		<title>��� ���� �ý���</title>
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
							  <th class="first" scope="col">�����</th>
							  <th scope="col">����</th>
						<% For i = 1 To  12	%>
							  <th scope="col"><%=i%>��</th>
						<% Next	%>
							  <th scope="col">�հ�</th>
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
								<td>����</td>
						<%
							For k = 1 To 13
						%>
								<td class="right"><%=FormatNumber(sum_amt(i, 1, k), 0)%></td>
						<%
							Next
						%>
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">���</td>
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
							  <td style="border-left:1px solid #e3e3e3;">����</td>
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
							  	<td rowspan="3" class="first" bgcolor="#CCFFFF"><strong>��</strong></td>
								<td>����</td>
						<%
							For k = 1 To 13
						%>
								<td class="right"><%=FormatNumber(sum_amt(0, 1, k), 0)%></td>
						<%
							Next
						%>
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">���</td>
						<%
							For k = 1 To 13
						%>
								<td class="right"><%=FormatNumber(sum_amt(0, 2, k), 0)%></td>
						<%
							Next
						%>
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">����</td>
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