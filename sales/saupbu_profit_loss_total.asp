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
Dim cost_year, base_year, be_year
Dim view_sw, i, j, k

Dim year_tab(15)	'�˻��⵵
Dim sum_amt(20, 3, 13)
Dim saupbu_tab(20)

Dim rsSalesDept, rsCostStats, rsSaleStats, rsProfitStats, rsEtcStats
Dim title_line, cost_saupbu
Dim ii, arrSalesDept
Dim mm, rsManage, rsPart, manageCost, partCost, end_month
Dim part_tot_cost, as_tot_cnt, part_cnt, rsSaupbuPart
Dim rsKsysPart, ksysPartCost, exceptDate

cost_year = f_Request("cost_year")	'��ȸ �⵵

title_line = "����κ� ���� �Ѱ� ��Ȳ"

If cost_year = "" Then
	cost_year = Mid(CStr(Now()),1 , 4)
	base_year = cost_year
	view_sw = "0"
End If

be_year = Int(cost_year) - 1

'�˻� ��ȸ �⵵
For i = 1 To 15
	year_tab(i) = Int(Mid(CStr(Now()), 1, 4)) - i + 1
Next

'For i = 0 To 4
'	year_tab(i) = Int(Mid(CStr(Now()), 1, 4)) + i
'Next

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

' 2019.02.22 ������ ��û '����κ� �����Ѱ�'���� �ش�⵵�� ����θ� �����ϸ��
' �������� ����
objBuilder.Append "SELECT saupbu FROM sales_org "
objBuilder.Append "WHERE sales_year='" & cost_year & "' AND sort_seq <> '31' "	'OA���ົ�δ� ����

'If team = "ȸ���繫" Or user_id = "102592" Then
If team <> "ȸ���繫" And user_id <> "102592" Then
'	objBuilder.Append "ORDER BY sort_seq ASC "
'Else
	' ȸ���繫 �϶��� ��Ÿ����ΰ� ������ ����..
	objBuilder.Append "	AND saupbu NOT IN ('��Ÿ�����') "
'	objBuilder.Append "ORDER BY sort_seq ASC "
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
	ii = 0
	For i = LBound(arrSalesDept) To UBound(arrSalesDept, 2)

'Do Until rsSalesDept.EOF
	ii = ii + 1
'	saupbu_tab(i) = rsSalesDept("saupbu")
	saupbu_tab(ii) = arrSalesDept(0, i)

'	rsSalesDept.MoveNext()
'Loop
	Next
End If

'---------------------------------------------------------------------------------------------------------------
'// 2017-09-15 ȸ���繫 ���� ��Ÿ�����,ȸ�簣�ŷ� ��ȸ �����ϰ� ����
'---------------------------------------------------------------------------------------------------------------

If team="ȸ���繫" Or user_id = "102592" Then
	'i = i + 1
	'saupbu_tab(i) = "��Ÿ�����"
	'i = i + 1
	'saupbu_tab(i) = "ȸ�簣�ŷ�"
	'i = i + 1
'	saupbu_tab(i) = "�ַ�ǻ����"

	' ȸ�簣�ŷ�
	'sql = "select cost_center,sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from company_cost where cost_year = '"&cost_year&"' and (cost_center = 'ȸ�簣�ŷ�') group by cost_center"
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

	rsCostStats.Close() : Set rsCostStats = Nothing
End If

'---------------------------------------------------------------------------------------------------------------
' ���� ����
objBuilder.Append "SELECT SUBSTRING(sales_date, 1, 7) AS sales_month, "
objBuilder.Append "	saupbu,	SUM(cost_amt) AS cost  "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(sales_date,1,4) = '"&cost_year&"' "
objBuilder.Append "GROUP BY SUBSTRING(sales_date,1,7), saupbu "

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

objBuilder.Append "SELECT saupbu, SUM(cost_amt_01), SUM(cost_amt_02), "
objBuilder.Append "	SUM(cost_amt_03), SUM(cost_amt_04), SUM(cost_amt_05), "

'�б⺰ ��� ���� ����(6,12�� �������� ����)
objBuilder.Append "	SUM(cost_amt_06), "
'objBuilder.Append "	(SUM(cost_amt_06) "
'objBuilder.Append "	- (SELECT SUM(cost_amt_06) FROM saupbu_profit_loss "
'objBuilder.Append "		WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' "
'objBuilder.Append "		AND saupbu = splt.saupbu)), "

objBuilder.Append "	SUM(cost_amt_07), SUM(cost_amt_08), "
objBuilder.Append "	SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "

objBuilder.Append "	SUM(cost_amt_12) "
'objBuilder.Append "	(SUM(cost_amt_12) "
'objBuilder.Append "	- (SELECT SUM(cost_amt_12) FROM saupbu_profit_loss "
'objBuilder.Append "		WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' "
'objBuilder.Append "		AND saupbu = splt.saupbu)) "

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

				sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(f_toString(rsProfitStats(k), 0)) + CDbl(manageCost) + CDbl(partCost) + CDbl(ksysPartCost)
			Next

			Exit For
		End If
	Next

	rsProfitStats.MoveNext()
Loop
Set rsManage = Nothing
Set rsPart = Nothing
Set rsKsysPart = Nothing
rsProfitStats.Close() : Set rsProfitStats = Nothing

Dim rsPartEtc, partEtcCost, rsSaupbuPartEtc, part_etc_cnt, part_etc_tot_cost, as_etc_tot_cnt

' ��� ���� (��Ÿ�����)
'sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' and (saupbu = '' or saupbu = '��Ÿ�����') group by saupbu"

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

' ���Ͱ��(i:����, j:����, k:��)
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
		For k = 1 To  12
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
			sum_amt(0,j,k) = sum_amt(0,j,k) + sum_amt(i,j,k)
		Next
	Next
Next
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�������� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
  	    <script src="/java/jquery-1.9.1.js"></script>
  	    <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}

			function frmcheck(){
				var c_year = parseInt($('#cost_year').val());

				if(c_year < 2021){
					$('#frm').attr('action', '/sales/old/saupbu_profit_loss_total_old.asp').submit();
				}else{
					document.frm.submit();
				}
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/sales/saupbu_profit_loss_total.asp" method="post" name="frm" id="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
						<dd>
							<p>
								<label>
									&nbsp;&nbsp;<strong>��ȸ�⵵&nbsp;</strong> :
									<select name="cost_year" id="cost_year" style="width:70px">
									<%
									For i = 1 To 15
									%>
										<option value="<%=year_tab(i)%>" <%If CInt(cost_year) = CInt(year_tab(i)) Then%>selected<%End If %>>&nbsp;<%=year_tab(i)%></option>
									<%
									Next
									%>
									</select>
								</label>
								<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
							</p>
						</dd>
					</dl>
				</fieldset>
				<div  style="text-align:right"><strong>�ݾ״��� : õ��</strong></div>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
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
							  <th class="first" scope="col">����</th>
							  <th scope="col">����</th>
							  <%For i = 1 To 12	%>
							  <th scope="col"><%=i%>��</th>
							  <%Next%>
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
								<td class="right"><%=FormatNumber(sum_amt(i, 1, k)/1000, 0)%></td>
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
								<%
								'If(k < 13 And sum_amt(i, 2, k) > 0) And (saupbu_tab(i) <> "ȸ�簣�ŷ�") Then
								If k < 13 And saupbu_tab(i) <> "ȸ�簣�ŷ�" Then
								%>
									<a href="#" onClick="pop_Window('/sales/saupbu_profit_loss_report2.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>&sales_saupbu=<%=saupbu_tab(i)%>','saupbu_profit_loss_report_pop','scrollbars=yes,width=1230,height=650')"><%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%></a>
								<%
								Else 'ȸ�簣 �ŷ�
									If(k < 13 And sum_amt(i, 2, k) > 0) And (saupbu_tab(i) = "ȸ�簣�ŷ�") Then
								%>
									<a href="#" onClick="pop_Window('/company_deal_detail_view.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>','company_deal_detail_view_pop','scrollbars=yes,width=1000,height=600')"><%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%></a>
								<%
									Else	'�հ�
								%>
									<%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%>
								<%
									End If
								End If
								%>
								</td>
								<%Next%>
			              	</tr>

							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">����</td>
								<%
								For k = 1 To 13
								%>
								<td class="right"><%=FormatNumber(sum_amt(i, 3, k)/1000, 0)%></td>
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
								<td class="right"><%=FormatNumber(sum_amt(0, 1, k)/1000, 0)%></td>
							<%
							Next
							%>
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">���</td>
							<%
							For k = 1 To 13
							%>
								<td class="right"><%=FormatNumber(sum_amt(0, 2 ,k)/1000, 0)%></td>
							<%
							Next
							%>
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">����</td>
							<%
							For k = 1 To 13
							%>
								<td class="right"><%=FormatNumber(sum_amt(0, 3, k)/1000, 0)%></td>
							<%
							Next
							%>
			              </tr>
						</tbody>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="/sales/saupbu_profit_loss_total_excel.asp?cost_year=<%=cost_year%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
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