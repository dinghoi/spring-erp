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
Dim sum_amt(10)
Dim tot_amt(10)
Dim detail_tab(30)
Dim cost_amt(30,10)

Dim cost_tab

Dim sales_saupbu, cost_year, cost_mm, cost_month
Dim before_year, before_mm, before_month, c_month, b_month
Dim condi_sql

Dim i
Dim rsPreCostSum, before_sales_amt
Dim rsCurrCostSum, curr_sales_amt

Dim title_line
Dim exceptDate

cost_tab = Array("�ΰǺ�","��Ư��","�Ϲݰ��","�����","����ī��","������","���ֺ�","����","���","��ݺ�","�󰢺�")

sales_saupbu = f_Request("sales_saupbu")
cost_year = f_Request("cost_year")
cost_mm = Right("0" & CStr(f_Request("cost_mm")), 2)
cost_month = CStr(cost_year) & CStr(cost_mm)

title_line = sales_saupbu & " ���� ��Ȳ"

If sales_saupbu = "" Then
	title_line = "��Ÿ����� ���� ��Ȳ"
End If

If cost_mm = "01" Then
	before_year = CStr(Int(cost_year) - 1)
	before_mm = "12"
Else
	before_year = cost_year
	before_mm = Right("0" & CStr(Int(cost_mm) - 1),2)
End If

before_month = CStr(before_year) & CStr(before_mm)	'���� �⵵(yyyyMM)
c_month = CStr(cost_year) & "-" & CStr(cost_mm)		'��� �⵵(yyyy-MM)
b_month = CStr(before_year) & "-" & CStr(before_mm)	'���� �⵵(yyyy-MM)

'If sales_saupbu = "��ü" Then
'	condi_sql = ""
'Else
'	condi_sql = " AND saupbu ='"&sales_saupbu&"' "
'End If

'If sales_saupbu = "��Ÿ�����" Then
'	condi_sql = " AND (saupbu ='' OR saupbu = '��Ÿ�����') "
'End If

'If sales_saupbu = "����" OR sales_saupbu = "�����׷�" Then
'	condi_sql = " AND saupbu IN ('����', '�����׷�') "
'End If

Select Case sales_saupbu
	Case "��ü"
		condi_sql = ""
	Case "��Ÿ�����"
		condi_sql = " AND (saupbu ='' OR saupbu = '��Ÿ�����') "
	Case "����", "�����׷�"
		condi_sql = " AND saupbu IN ('����', '�����׷�') "
	Case Else
		condi_sql = " AND saupbu ='"&sales_saupbu&"' "
End Select

For i = 0 To 10
	sum_amt(i) = 0
	tot_amt(i) = 0
Next

'202204������ �������� SI1���� ���� �Ｚ������(��) ���� ���� ó��(�繫 ��û)[����ȣ_20220511]
exceptDate = "202204"

'�����(����)
'sql = "SELECT SUM(cost_amt) AS sales_amt "&_
'	  "  FROM saupbu_sales "&_
'	  " WHERE SUBSTRING(SALES_DATE,1,7) = '"&b_month&"'"&condi_sql
'Set rs_sum = Dbconn.Execute(sql)
objBuilder.Append "SELECT SUM(cost_amt) AS sales_amt "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(SALES_DATE, 1, 7) = '"&b_month&"'"&condi_sql

Set rsPreCostSum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsPreCostSum(0)) Then
	before_sales_amt = 0
Else
	before_sales_amt = CDbl(rsPreCostSum(0))
End If

rsPreCostSum.Close()
Set rsPreCostSum = Nothing

Dim rsBeforeManage, beforeManageCost, rsBeforePart, beforePartCost
Dim rsCurrentManage, currentManageCost, rsCurrentPart, currentPartCost

beforeManageCost = 0
currentManageCost = 0
beforePartCost = 0
currentPartCost = 0

'�������� �հ�(����)
objBuilder.Append "SELECT ROUND((tot_cost_amt * 0.5 / tot_person * saupbu_person) "
objBuilder.Append "	+ (tot_cost_amt * 0.5 / tot_sale * saupbu_sale), 1) AS tot_amt "
objBuilder.Append "FROM ( "
objBuilder.Append "	select mgct.saupbu, mgct.tot_cost_amt, mgct.saupbu_person, mgct.tot_person, "
objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "		FROM saupbu_sales "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&Replace(b_month, "-", "")&"' "
objBuilder.Append "			AND mgct.saupbu = saupbu "

If Replace(b_month, "-", "") >= exceptDate Then
	objBuilder.Append "		AND company <> '�Ｚ������(��)' "
End If

objBuilder.Append "		) AS saupbu_sale, "

objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "		FROM saupbu_sales "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&Replace(b_month, "-", "")&"' "
objBuilder.Append "			AND saupbu <> '��Ÿ�����' "

If Replace(b_month, "-", "") >= exceptDate Then
	objBuilder.Append "		AND company <> '�Ｚ������(��)' "
End If

objBuilder.Append "		) AS tot_sale "

objBuilder.Append "	FROM management_cost AS mgct "
objBuilder.Append "	WHERE cost_month = '"&Replace(b_month, "-", "")&"' "
objBuilder.Append "		AND saupbu = '"&sales_saupbu&"' "
objBuilder.Append "	GROUP BY saupbu "
objBuilder.Append ") r1 "

Set rsBeforeManage = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not (rsBeforeManage.BOF Or rsBeforeManage.EOF) Then
	beforeManageCost = rsBeforeManage(0)
End If

rsBeforeManage.Close() : Set rsBeforeManage = Nothing

'�ι������ �հ�(����)
'objBuilder.Append "SELECT IFNULL(SUM(cost_amt), 0) "
'objBuilder.Append "FROM company_asunit "
'objBuilder.Append "WHERE as_month = '"&Replace(b_month, "-", "")&"' "

'If sales_saupbu = "��Ÿ�����" Then
'	objBuilder.Append "	AND saupbu = '' "
'Else
'	objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
'End If

'Set rsBeforePart = DBConn.Execute(objBuilder.ToString())
'objBuilder.Clear()

'beforePartCost = rsBeforePart(0)

'rsBeforePart.Close() : Set rsBeforePart = Nothing

Dim rsPart, part_tot_cost, as_tot_cnt, rsSaupbuPart, part_cnt
'�ι������(���) - ����
objBuilder.Append "SELECT (SUM(cost_amt_"&before_mm&") - "
objBuilder.Append "(SELECT SUM(cost_amt_"&before_mm&") FROM company_cost WHERE cost_year ='"&before_year&"' "
objBuilder.Append "	AND cost_detail = '��ġ����')) AS 'part_tot_cost', "
objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&before_year&before_mm&"') AS 'as_tot_cnt' "
objBuilder.Append "FROM company_cost WHERE cost_year = '"&before_year&"' AND cost_center = '�ι������' "

Set rsPart = DBConn.Execute(objBuilder.ToString())

part_tot_cost = CDbl(f_toString(rsPart("part_tot_cost"), 0))	'�ι������(���)
as_tot_cnt = CInt(f_toString(rsPart("as_tot_cnt"), 0))	'AS �� �Ǽ�

objBuilder.Clear()
rsPart.Close() : Set rsPart = Nothing

'����� �� AS �� �Ǽ� ��ȸ
objBuilder.Append "SELECT SUM(as_total - as_set) AS as_cnt "
objBuilder.Append "FROM as_acpt_status AS aast "
objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
objBuilder.Append "	AND trdt.trade_id = '����' "
objBuilder.Append "WHERE as_month = '"&before_year&before_mm&"' "

If sales_saupbu = "��Ÿ�����" Then
	objBuilder.Append "	AND trdt.saupbu = '' "
Else
	objBuilder.Append "	AND trdt.saupbu = '"&sales_saupbu&"' "
End If

Set rsSaupbuPart = DBConn.Execute(objBuilder.ToString())

part_cnt = CInt(f_toString(rsSaupbuPart(0), 0))	'����� AS �� �Ǽ�

objBuilder.Clear()
rsSaupbuPart.Close() : Set rsSaupbuPart = Nothing

'����κ� ��� �κа����
If part_cnt > 0 Then
	beforePartCost = part_tot_cost / as_tot_cnt * part_cnt
End If

Dim rsKsysPart, beforeKsysPartCost, currentKsysPartCost

'����κ� ��� �ι������(2)(����)
objBuilder.Append "SELECT ROUND((part_tot * 0.5 / tot_person * saupbu_person) + (part_tot * 0.5 / tot_sale * saupbu_sale), 1) FROM ("
objBuilder.Append "	SELECT mgct.saupbu, mgct.saupbu_person, "
objBuilder.Append "		(SELECT SUM(cost_amt_"&before_mm&") FROM company_cost WHERE cost_year = '"&before_year&"' AND cost_center = '�ι������(2)') AS 'part_tot',"
objBuilder.Append "		(SELECT count(*) FROM pay_month_give AS pmgt "
objBuilder.Append "		INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no AND emp_month = '"&before_year&before_mm&"' "
objBuilder.Append "		WHERE pmg_yymm = '"&before_year&before_mm&"' AND pmgt.mg_saupbu IN ('����SI����', '����SI����', 'DI����ι�') "
objBuilder.Append "			AND pmg_id = '1' AND pmg_emp_type = '����' AND emmt.cost_except IN ('0', '1')) AS tot_person, "
objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt FROM saupbu_sales "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&before_year&before_mm&"' AND mgct.saupbu = saupbu) AS saupbu_sale, "
objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt FROM saupbu_sales "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&before_year&before_mm&"' AND saupbu IN ('����SI����', '����SI����', 'DI����ι�')) AS tot_sale"
objBuilder.Append "	FROM management_cost AS mgct "
objBuilder.Append "	WHERE cost_month = '"&before_year&before_mm&"' AND saupbu IN ('����SI����', '����SI����', 'DI����ι�') "
objBuilder.Append "	GROUP BY saupbu "
objBuilder.Append ") r1 WHERE r1.saupbu= '"&sales_saupbu&"' "

Set rsKsysPart = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsKsysPart.EOF Or rsKsysPart.BOF Then
	beforeKsysPartCost = 0
Else
	beforeKsysPartCost = f_toString(rsKsysPart(0), 0)
End If
rsKsysPart.Close() : Set rsKsysPart = Nothing

'================================================================

'�����(���)
'sql = "SELECT SUM(cost_amt) AS sales_amt "&_
'	  "  FROM saupbu_sales "&_
'	  " WHERE SUBSTRING(sales_date,1,7) = '"&c_month&"'"&condi_sql
'Set rs_sum = Dbconn.Execute (sql)
objBuilder.Append "SELECT SUM(cost_amt) AS sales_amt "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&c_month&"'"&condi_sql

Set rsCurrCostSum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsCurrCostSum(0)) Then
	curr_sales_amt = 0
Else
	curr_sales_amt = CDbl(rsCurrCostSum(0))
End If

rsCurrCostSum.Close()
Set rsCurrCostSum = Nothing

'�������� �հ�(���)
objBuilder.Append "SELECT ROUND((tot_cost_amt * 0.5 / tot_person * saupbu_person) "
objBuilder.Append "	+ (tot_cost_amt * 0.5 / tot_sale * saupbu_sale), 1) AS tot_amt "
objBuilder.Append "FROM ( "
objBuilder.Append "	select mgct.saupbu, mgct.tot_cost_amt, mgct.saupbu_person, mgct.tot_person, "
objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "		FROM saupbu_sales "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&Replace(c_month, "-", "")&"' "
objBuilder.Append "			AND mgct.saupbu = saupbu "

If Replace(c_month, "-", "") >= exceptDate Then
	objBuilder.Append "		AND company <> '�Ｚ������(��)' "
End If

objBuilder.Append "		) AS saupbu_sale, "
objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "		FROM saupbu_sales "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&Replace(c_month, "-", "")&"' "
objBuilder.Append "			AND saupbu <> '��Ÿ�����' "

If Replace(c_month, "-", "") >= exceptDate Then
	objBuilder.Append "		AND company <> '�Ｚ������(��)' "
End If

objBuilder.Append "		) AS tot_sale "
objBuilder.Append "	FROM management_cost AS mgct "
objBuilder.Append "	WHERE cost_month = '"&Replace(c_month, "-", "")&"' "
objBuilder.Append "		AND saupbu = '"&sales_saupbu&"' "
objBuilder.Append "	GROUP BY saupbu "
objBuilder.Append ") r1 "

Set rsCurrentManage = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not (rsCurrentManage.EOF Or rsCurrentManage.BOF) Then
	currentManageCost = rsCurrentManage(0)
End If

rsCurrentManage.Close() : Set rsCurrentManage = Nothing

'�ι������ �հ�(���)
'objBuilder.Append "SELECT IFNULL(SUM(cost_amt), 0) "
'objBuilder.Append "FROM company_asunit "
'objBuilder.Append "WHERE as_month = '"&Replace(c_month, "-", "")&"' "

'If sales_saupbu = "��Ÿ�����" Then
'	objBuilder.Append "	AND saupbu = '' "
'Else
'	objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
'End If

'Set rsCurrentPart = DBConn.Execute(objBuilder.ToString())
'objBuilder.Clear()

'currentPartCost = rsCurrentPart(0)

'rsCurrentPart.Close() : Set rsCurrentPart = Nothing

'�ι������(���) - ���
objBuilder.Append "SELECT (SUM(cost_amt_"&cost_mm&") - "
objBuilder.Append "(SELECT SUM(cost_amt_"&cost_mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_detail = '��ġ����')) AS 'part_tot_cost', "
objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&cost_year&cost_mm&"') AS 'as_tot_cnt' "
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
objBuilder.Append "WHERE as_month = '"&cost_year&cost_mm&"' "

If sales_saupbu = "��Ÿ�����" Then
	objBuilder.Append "	AND trdt.saupbu = '' "
Else
	objBuilder.Append "	AND trdt.saupbu = '"&sales_saupbu&"' "
End If

Set rsSaupbuPart = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

part_cnt = CInt(f_toString(rsSaupbuPart(0), 0))	'����� AS �� �Ǽ�

rsSaupbuPart.Close() : Set rsSaupbuPart = Nothing

'����κ� ��� �κа����
currentPartCost = part_tot_cost / as_tot_cnt * part_cnt

'����κ� ��� �ι������(2)(����)
objBuilder.Append "SELECT ROUND((part_tot * 0.5 / tot_person * saupbu_person) + (part_tot * 0.5 / tot_sale * saupbu_sale), 1) FROM ("
objBuilder.Append "	SELECT mgct.saupbu, mgct.saupbu_person, "
objBuilder.Append "		(SELECT SUM(cost_amt_"&cost_mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '�ι������(2)') AS 'part_tot',"
objBuilder.Append "		(SELECT count(*) FROM pay_month_give AS pmgt "
objBuilder.Append "		INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no AND emp_month = '"&cost_year&cost_mm&"' "
objBuilder.Append "		WHERE pmg_yymm = '"&cost_year&cost_mm&"' AND pmgt.mg_saupbu IN ('����SI����', '����SI����', 'DI����ι�') "
objBuilder.Append "			AND pmg_id = '1' AND pmg_emp_type = '����' AND emmt.cost_except IN ('0', '1')) AS tot_person, "
objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt FROM saupbu_sales "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&cost_year&cost_mm&"' AND mgct.saupbu = saupbu) AS saupbu_sale, "
objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt FROM saupbu_sales "
objBuilder.Append "		WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&cost_year&cost_mm&"' AND saupbu IN ('����SI����', '����SI����', 'DI����ι�')) AS tot_sale"
objBuilder.Append "	FROM management_cost AS mgct "
objBuilder.Append "	WHERE cost_month = '"&cost_year&cost_mm&"' AND saupbu IN ('����SI����', '����SI����', 'DI����ι�') "
objBuilder.Append "	GROUP BY saupbu "
objBuilder.Append ") r1 WHERE r1.saupbu= '"&sales_saupbu&"' "

Set rsKsysPart = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsKsysPart.EOF Or rsKsysPart.BOF Then
	currentKsysPartCost = 0
Else
	currentKsysPartCost = f_toString(rsKsysPart(0), 0)
End If
rsKsysPart.Close() : Set rsKsysPart = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
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
				if (chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if (document.frm.cost_month.value == ""){
					alert ("��ȸ���� �Է��ϼ���.");
					return false;
				}
				return true;
			}

			function scrollAll(){
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td>
							<div id="topLine2" style="width:1200px;overflow:hidden;">
								<div class="gView">
									<table cellpadding="0" cellspacing="0" class="tableList">
										<colgroup>
											<col width="4%" >
											<col width="*" >
											<col width="8%" >
											<col width="6%" >
											<col width="6%" >
											<col width="7%" >
											<col width="9%" >
											<col width="8%" >
											<col width="6%" >
											<col width="6%" >
											<col width="7%" >
											<col width="9%" >
											<col width="7%" >
											<col width="5%" >
											<col width="1%" >
										</colgroup>
										<thead>
											<tr>
											  <th rowspan="2" class="first" scope="col">����׸�</th>
											  <th rowspan="2" scope="col">���γ���</th>
											  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� ��&nbsp;(<%=before_year%>��<%=before_mm%>��)</th>
											  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� ��&nbsp;(<%=cost_year%>��<%=cost_mm%>��)</th>
											  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">����</th>
											  <th rowspan="2" scope="col"></th>
										  	</tr>
											<tr>
											  <th scope="col" style="border-left:1px solid #e3e3e3;">����������</th>
											  <th scope="col">������</th>
											  <th scope="col">��������</th>
											  <th scope="col">�ι������</th>
											  <th scope="col">��</th>
											  <th scope="col">����������</th>
											  <th scope="col">������</th>
											  <th scope="col">��������</th>
											  <th scope="col">�ι������</th>
											  <th scope="col">��</th>
											  <th scope="col">�ݾ�</th>
											  <th scope="col">��</th>
				              				</tr>
										</thead>
									</table>
								</div>
							</div>
						</td>
					</tr>
					<tr>
          				<td valign="top">
				    	<DIV id="mainDisplay2" style="width:1200;height:470px;overflow:scroll" onscroll="scrollAll()">
				    	<table cellpadding="0" cellspacing="0" class="scrollList">
				    		<colgroup>
								<col width="6%" >
								<col width="*" >
								<col width="8%" >
								<col width="6%" >
								<col width="6%" >
								<col width="7%" >
								<col width="9%" >
								<col width="8%" >
								<col width="6%" >
								<col width="6%" >
								<col width="7%" >
								<col width="9%" >
								<col width="7%" >
								<col width="5%" >
								<col width="1%" >
							</colgroup>
							<tbody>
								<tr bgcolor="#FFFFCC">
									<td colspan="2" class="first" scope="col"><strong>�����</strong></td>
									<td colspan="5" scope="col"><strong><%=FormatNumber(before_sales_amt, 0)%></strong></td>
									<td colspan="5" scope="col"><strong><%=FormatNumber(curr_sales_amt, 0)%></strong></td>
									<%
									Dim incr_amt, incr_per

									incr_amt = curr_sales_amt - before_sales_amt

									If before_sales_amt = 0 And curr_sales_amt = 0 Then
										incr_per = 0
							  		ElseIf before_sales_amt = 0 Then
										incr_per = 100
							  		Else
						   				incr_per = incr_amt / before_sales_amt * 100
						   			End If
									%>
									<td scope="col" class="right"><%=FormatNumber(incr_amt, 0)%></td>
									<td scope="col" class="right"><%=FormatNumber(incr_per, 2)%>%</td>
									<td scope="col" class="right">&nbsp;</td>
                    			</tr>
								<%
								Dim jj, rec_cnt, j
								Dim rsCostDetail, rsCostSum

								For jj = 0 To 10
									rec_cnt = 0

									For i = 1 To 30
										detail_tab(i) = ""

										For j = 1 To 10
											cost_amt(i, j) = 0
											sum_amt(j) = 0
										Next
									Next

									If cost_tab(jj) = "�ΰǺ�" Then
										'sql = "   SELECT cost_detail "&_
										'	  "     FROM SAUPBU_COST_ACCOUNT "&_
										'	  "    WHERE cost_id = '�ΰǺ�' "&_
										'	  " ORDER BY view_seq"
										'rs.Open sql, Dbconn, 1
										objBuilder.Append "SELECT cost_detail "
										objBuilder.Append "FROM SAUPBU_COST_ACCOUNT "
										objBuilder.Append "WHERE cost_id = '�ΰǺ�' "
										objBuilder.Append "ORDER BY view_seq "

										Set rsCostDetail = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										Do Until rsCostDetail.EOF
											rec_cnt = rec_cnt + 1
											detail_tab(rec_cnt) = rsCostDetail("cost_detail")

											rsCostDetail.MoveNext()
										Loop

										rsCostDetail.Close()
									Else
										'sql = "   SELECT cost_detail "&_
										'	  "     FROM SAUPBU_PROFIT_LOSS "&_
										'	  "    WHERE (cost_year = '"& cost_year &"' OR cost_year = '"& before_year &"') "&_
										'	  "      AND cost_id ='"& cost_tab(jj) &"'"& condi_sql &_
										'	  " GROUP BY cost_detail "&_
										'	  " ORDER BY cost_detail"
										'rs.Open sql, Dbconn, 1
										objBuilder.Append "SELECT cost_detail "
										objBuilder.Append "FROM SAUPBU_PROFIT_LOSS "
										objBuilder.Append "WHERE (cost_year = '"& cost_year &"' OR cost_year = '"& before_year &"') "
										objBuilder.Append "	AND cost_id ='"& cost_tab(jj) &"'"& condi_sql
										objBuilder.Append "GROUP BY cost_detail "
										objBuilder.Append "ORDER BY cost_detail "

										Set rsCostDetail = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										Do Until rsCostDetail.EOF
											rec_cnt = rec_cnt + 1
											detail_tab(rec_cnt) = rsCostDetail("cost_detail")

											rsCostDetail.MoveNext()
										Loop

										rsCostDetail.Close()
									End If

									If rec_cnt <> 0 Then
										' ���� �ݾ� SUM
										'sql = "  SELECT cost_center "&_
										' 	  "       , cost_detail "&_
										'	  "       , SUM(cost_amt_"& before_mm &") AS cost " &_
										'	  "    FROM SAUPBU_PROFIT_LOSS "&_
										'	  "   WHERE cost_year = '"& before_year &"' "&_
										'	  "     AND cost_id   = '"& cost_tab(jj) &"'"&condi_sql &_
										'	  " GROUP BY cost_center, cost_detail "&_
										'	  " ORDER BY cost_center, cost_detail"
										'rs.Open sql, Dbconn, 1

										'�б⺰ ��� ���� ����(6��, 12��, �������� ����)
										'If (cost_mm = "06" Or cost_mm = "12") AND cost_tab(jj) = "�Ϲݰ��" Then
										'	objBuilder.Append "SELECT cost_center, cost_detail, "
										'	objBuilder.Append "	CASE WHEN cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' THEN "
										'	objBuilder.Append "		SUM(cost_amt_"&before_mm&") - (SELECT SUM(cost_amt_"&before_mm&") FROM saupbu_profit_loss WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' AND saupbu = splt.saupbu) "
										'	objBuilder.Append "	ELSE SUM(cost_amt_"&before_mm&") END AS 'cost' "
										'Else
											objBuilder.Append "SELECT cost_center, cost_detail, SUM(cost_amt_"&before_mm&") AS cost "
										'End If

										objBuilder.Append "FROM SAUPBU_PROFIT_LOSS AS splt "
										objBuilder.Append "WHERE cost_year = '"& before_year &"' "
										objBuilder.Append "	AND cost_id = '"& cost_tab(jj) &"'"&condi_sql
										objBuilder.Append "	AND cost_center NOT IN ('�ι������', '��������') "
										objBuilder.Append "GROUP BY cost_center, cost_detail "
										objBuilder.Append "ORDER BY cost_center, cost_detail "

										Set rsPreCostSum = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										Do Until rsPreCostSum.EOF
											For i = 1 To 30
												' �������� ������ detail_tab�� ���ٸ� cost_detail�� ������ �ʴ´�..
												If rsPreCostSum("cost_detail") = detail_tab(i) Then
													Select Case rsPreCostSum("cost_center")
														Case "����������" : j = 1
														Case "������"     : j = 2
														Case "��������" : j = 3
														Case "�ι������" : j = 4
													End Select

													cost_amt(i, j) = cost_amt(i, j) + CDbl(rsPreCostSum("cost"))
													cost_amt(i, 5) = cost_amt(i, 5) + CDbl(rsPreCostSum("cost"))
													sum_amt(j) = sum_amt(j) + CDbl(rsPreCostSum("cost"))
													sum_amt(5) = sum_amt(5) + CDbl(rsPreCostSum("cost"))
													tot_amt(j) = tot_amt(j) + CDbl(rsPreCostSum("cost"))
													tot_amt(5) = tot_amt(5) + CDbl(rsPreCostSum("cost"))

													Exit For
												End If
											Next

											rsPreCostSum.MoveNext()
										Loop
										rsPreCostSum.Close()

										' ��� �ݾ� SUM
										'sql = "    SELECT cost_center "&_
										'	  "         , cost_detail "&_
										'	  "         , SUM(cost_amt_"&cost_mm&") AS cost "&_
										'	  "      FROM  SAUPBU_PROFIT_LOSS "&_
										'	  "     WHERE  cost_year ='"& cost_year &"' "&_
										'	  "       AND  cost_id   ='"& cost_tab(jj) &"'"&condi_sql&" "&_
										'	  " GROUP  BY cost_center, cost_detail "&_
										'	  " ORDER  BY cost_center, cost_detail"
										'rs.Open sql, Dbconn, 1

										'�б⺰ ��� ���� ����(6��, 12��, �������� ����)
										'If (cost_mm = "06" Or cost_mm = "12") AND cost_tab(jj) = "�Ϲݰ��" Then
										'	objBuilder.Append "SELECT cost_center, cost_detail, "
										'	objBuilder.Append "	CASE WHEN cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' THEN "
										'	objBuilder.Append "		SUM(cost_amt_"&cost_mm&") - (SELECT SUM(cost_amt_"&cost_mm&") FROM saupbu_profit_loss WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' AND saupbu = splt.saupbu) "
										'	objBuilder.Append "	ELSE SUM(cost_amt_"&cost_mm&") END AS 'cost' "

										'Else
											objBuilder.Append "SELECT cost_center, cost_detail, SUM(cost_amt_"&cost_mm&") AS cost "
										'End If
										objBuilder.Append "FROM  SAUPBU_PROFIT_LOSS AS splt "
										objBuilder.Append "WHERE  cost_year ='"& cost_year &"' "
										objBuilder.Append "	AND cost_id ='"& cost_tab(jj) &"' "&condi_sql
										objBuilder.Append "	AND cost_center NOT IN ('�ι������', '��������') "
										objBuilder.Append "GROUP  BY cost_center, cost_detail "
										objBuilder.Append "ORDER  BY cost_center, cost_detail "

										Set rsCurrCostSum = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										Do Until rsCurrCostSum.EOF
											For i = 1 To 30
												' �������� ������ detail_tab�� ���ٸ� cost_detail�� ������ �ʴ´�..
												If rsCurrCostSum("cost_detail") = detail_tab(i) Then
													Select Case rsCurrCostSum("cost_center")
														Case "����������"	: j = 6
														Case "������"	    : j = 7
														Case "��������"	: j = 8
														Case "�ι������"	: j = 9
													End Select

													cost_amt(i, j) = cost_amt(i, j) + CDbl(rsCurrCostSum("cost"))
													cost_amt(i, 10) = cost_amt(i, 10) + CDbl(rsCurrCostSum("cost"))
													sum_amt(j) = sum_amt(j) + CDbl(rsCurrCostSum("cost"))
													sum_amt(10) = sum_amt(10) + CDbl(rsCurrCostSum("cost"))
													tot_amt(j) = tot_amt(j) + CDbl(rsCurrCostSum("cost"))
													tot_amt(10) = tot_amt(10) + CDbl(rsCurrCostSum("cost"))

													Exit For
												End If
											Next
											rsCurrCostSum.MoveNext()
										Loop
										rsCurrCostSum.Close()
										%>
										<tr>
							  				<td rowspan="<%=rec_cnt + 1%>" class="first">
											<%
											If jj = 2 Or jj = 3 Then
												Response.Write cost_tab(jj) & "<BR/>(���ݻ��)"
											Else
												Response.Write cost_tab(jj)
											End If
											%>
                  							</td>
											<td class="left"><%=detail_tab(1)%></td>

											<%
											For j = 1 To 10
												If j = 5 Or j = 10 Then
													Response.write "<td class='right'><strong>"&FormatNumber(cost_amt(1, j), 0)&"</strong></td>"
												Else
													Response.write "<td class='right'>" ' [["&jj&"]][[cost_amt(1,"&j&")="&cost_amt(1,j)&"]]

													If jj < 2 Then
														Response.Write FormatNumber(cost_amt(1, j), 0)
													Else
														If(j = 1 Or j = 2 Or j = 6 Or j = 7) And jj > 1 And cost_amt(1,j) <> 0 Then
														%>
			                  								<a href="#" onClick="pop_Window('/sales/profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(1)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')">
																<%=FormatNumber(cost_amt(1, j), 0)%>
															</a>
														<%
														Else
			                  								Response.Write FormatNumber(cost_amt(1, j), 0)
			                  							End If
			                  						End If
			                  						%>
		                  							</td>
												<%
												End If
											Next

											incr_amt = cost_amt(1, 10) - cost_amt(1, 5)

											If cost_amt(1, 5) = 0 And cost_amt(1, 10) = 0 Then
												incr_per = 0
											ElseIf cost_amt(1, 5) = 0 Then
												incr_per = 100
											Else
												incr_per = incr_amt / cost_amt(1, 5) * 100
											End If
											%>
											<td class="right"><%=FormatNumber(incr_amt, 0)%></td>
											<td class="right"><%=FormatNumber(incr_per, 2)%>%</td>
											<td class="right">&nbsp;</td>
										</tr>
										<%For i = 2 To rec_cnt%>
										<tr>
											<td class="left" style=" border-left:1px solid #e3e3e3;"><%=detail_tab(i)%></td>
											<%
											For j = 1 To 10
												If j = 5 Or j = 10 Then
													Response.Write "<td class='right'><strong>"&FormatNumber(cost_amt(i, j), 0)&"</strong></td>"
												Else
											%>
											<td class="right">
												<%If jj < 2	Then	'//2016-08-23 �˹ٺ� ����ȸ ��ũ �߰�
													If detail_tab(i) = "�˹ٺ�" Then
													%>
														<a href="#" onClick="pop_Window('/sales/profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(i)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')">
															<%=FormatNumber(cost_amt(i, j), 0)%>
														</a>
													<%
													Else
														Response.Write FormatNumber(cost_amt(i, j), 0)
													End IF
													%>
												<%Else 	%>
													<%
													If (j = 1 Or j = 2 Or j = 6 Or j = 7) And jj > 1 And cost_amt(i, j) <>  0 Then%>
														<a href="#" onClick="pop_Window('/sales/profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(i)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')">
															<%=FormatNumber(cost_amt(i, j), 0)%>
														</a>
													<%
													Else
													%>
														<%=FormatNumber(cost_amt(i, j), 0)%>
													<%
													End If	%>
												<%End If%>
											</td>
											<%
												End If
											Next

											incr_amt = cost_amt(i, 10) - cost_amt(i, 5)

											If cost_amt(i, 5) = 0 And cost_amt(i, 10) = 0 Then
													incr_per = 0
											ElseIf cost_amt(i, 5) = 0 Then
												incr_per = 100
											Else
												incr_per = incr_amt / cost_amt(i,5) * 100
											End If
											%>
											<td class="right"><%=FormatNumber(incr_amt, 0)%></td>
											<td class="right"><%=FormatNumber(incr_per, 2)%>%</td>
											<td class="right">&nbsp;</td>
										</tr>
										<%Next	%>

										<!--=== �Ұ� ===-->
										<tr>
											<td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">�Ұ�</td>
											<%
											For j = 1 To 10
												If j = 5 Or j = 10 Then
											%>
											<td class="right" bgcolor="#EEFFFF"><strong><%=FormatNumber(sum_amt(j), 0)%></strong></td>
											<%
												Else
											%>
											<td class="right" bgcolor="#EEFFFF"><%=FormatNumber(sum_amt(j), 0)%></td>
											<%
												End If
											Next

											incr_amt = sum_amt(10) - sum_amt(5)

											If sum_amt(5) = 0 And sum_amt(10) = 0 Then
												incr_per = 0
											ElseIf sum_amt(5) = 0 Then
												incr_per = 100
											Else
												incr_per = incr_amt / sum_amt(5) * 100
											End If
											%>
											<td class="right" bgcolor="#EEFFFF"><%=FormatNumber(incr_amt, 0)%></td>
											<td class="right" bgcolor="#EEFFFF"><%=FormatNumber(incr_per, 2)%>%</td>
											<td class="right" bgcolor="#EEFFFF">&nbsp;</td>
										</tr>
									<%
									End If
								Next
								Set rsCostDetail = Nothing
								Set rsPreCostSum = Nothing
								Set rsCurrCostSum = Nothing
								%>
								<!--=====	����հ�	=====-->
								<tr bgcolor="#FFFFCC">
									<td colspan="2" class="first" scope="col"><strong>����հ�</strong></td>
									<%
									For j = 1 To 10
										If j = 5 Then
											tot_amt(j) = CDbl(tot_amt(j)) + CDbl(beforeManageCost) + CDbl(beforePartCost) + CDbl(beforeKsysPartCost)
									%>
									<td class="right" alt="����հ�(����) > ��"><strong><%=FormatNumber(tot_amt(j), 0)%></strong></td>
									<%
										ElseIf j = 10 Then
											tot_amt(j) = CDbl(tot_amt(j)) + CDbl(currentManageCost) + CDbl(currentPartCost) + CDbl(currentKsysPartCost)
									%>
									<td class="right" alt="����հ�(���) > ��"><strong><%=FormatNumber(tot_amt(j), 0)%></strong></td>
									<%
										ElseIf j = 3 Then
									%>
									<td class="right" alt="��������(����)" style="color:blue;"><%=FormatNumber(beforeManageCost, 0)%></td>
									<%
										ElseIf j = 8 Then
									%>
									<td class="right" alt="��������(���)" style="color:blue;"><%=FormatNumber(currentManageCost, 0)%></td>
									<%
										ElseIf j = 4 Then
									%>
									<td class="right" alt="�ι������(����)" style="color:blue;"><%=FormatNumber(beforePartCost + beforeKsysPartCost, 0)%></td>
									<%
										ElseIf j = 9 Then
									%>
									<td class="right" alt="�ι������(���)" style="color:blue;"><%=FormatNumber(currentPartCost + currentKsysPartCost, 0)%></td>
									<%
										Else
									%>
									<td class="right"><%=FormatNumber(tot_amt(j), 0)%></td>
									<%
										End If
									Next

									incr_amt = tot_amt(10) - tot_amt(5)

									If tot_amt(5) = 0 And tot_amt(10) = 0 Then
										incr_per = 0
									ElseIf tot_amt(5) = 0 Then
										incr_per = 100
									Else
										incr_per = incr_amt / tot_amt(5) * 100
									End if
									%>
									<td scope="col" class="right" alt="����հ� > �ݾ�"><%=FormatNumber(incr_amt, 0)%></td>
									<td scope="col" class="right" alt="����հ� > ��"><%=FormatNumber(incr_per, 2)%>%</td>
									<td scope="col" class="right">&nbsp;</td>
								</tr>

								<!--=====	����	=====-->
								<tr bgcolor="#FFDFDF">
									<td colspan="2" bgcolor="#FFDFDF" class="first" scope="col"><strong>����</strong></td>
									<%
										Dim be_profit_loss, curr_profit_loss

										be_profit_loss = before_sales_amt - tot_amt(5)
										curr_profit_loss = curr_sales_amt - tot_amt(10)
										incr_amt = curr_profit_loss - be_profit_loss

										If be_profit_loss = 0 And curr_profit_loss = 0 Then
											incr_per = 0
										ElseIf be_profit_loss = 0 Then
											incr_per = 100
										Else
											incr_per = incr_amt / be_profit_loss * 100
										End If

										If be_profit_loss < 0 Then
											incr_per = incr_per * -1
										End If
									%>
									<td scope="col" colspan="5"><strong><%=FormatNumber(be_profit_loss, 0)%></strong></td>
									<td scope="col" colspan="5"><strong><%=FormatNumber(curr_profit_loss, 0)%></strong></td>
									<td scope="col" class="right"><%=FormatNumber(incr_amt, 0)%></td>
									<td scope="col" class="right"><%=FormatNumber(incr_per, 2)%>%</td>
									<td scope="col" class="right">&nbsp;</td>
								</tr>
							</tbody>
                		</table>
              			</DIV>
						</td>
           			</tr>
					</table>

					<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  	<tr>
				    	<td>
							<div class="btnCenter">
			            		<a href="/sales/saupbu_profit_loss_excel.asp?cost_year=<%=cost_year%>&cost_mm=<%=cost_mm%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">ȭ�� �����ٿ�ε�</a>

								<%
								If empProfitViewAll = "Y" Then
								%>
			            		<a href="/sales/cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">���ֺ�/������ �����ٿ�ε�</a>
								<%
								ElseIf empProfitViewSI = "Y" And (sales_saupbu = "SI1����" Or sales_saupbu = "SI2����") Then
								%>
								<a href="/sales/cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">���ֺ�/������ �����ٿ�ε�</a>
								<%
								ElseIf empProfitViewNI = "Y" And (sales_saupbu = "ICT����" Or sales_saupbu = "NI����") Then
								%>
								<a href="/sales/cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">���ֺ�/������ �����ٿ�ε�</a>
								<%
								ElseIf sales_saupbu = bonbu Then
								%>
								<a href="/sales/cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">���ֺ�/������ �����ٿ�ε�</a>
								<%
								End If
								%>

			            		<a href="/sales/saupbu_sales_detail_excel2.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">����� �����ٿ�ε�</a>
								<%If sales_grade = "0" And empProfitViewAll = "Y" Then	%>
			            			<a href="/sales/cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=��������" class="btnType04">�������� �����ٿ�ε�</a>
			          				<a href="/sales/cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=�ι������" class="btnType04">�ι������ �����ٿ�ε�</a>
									<a href="/sales/cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=�ι������(2)" class="btnType04">�ι������(2) �����ٿ�ε�</a>
								<%End If%>
							</div>
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