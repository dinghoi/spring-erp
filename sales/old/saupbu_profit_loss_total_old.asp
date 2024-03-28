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

Dim year_tab(15)	'�˻� �⵵
Dim sum_amt(20, 3, 13)
Dim saupbu_tab(20)

Dim rsSalesDept, rsCostStats, rsSaleStats, rsProfitStats, rsEtcStats
Dim title_line
Dim cost_saupbu

cost_year = f_Request("cost_year")	'��ȸ �⵵

If cost_year = "" Then
	'cost_year = Mid(CStr(Now()),1 , 4)
	cost_year = "2020"
	base_year = cost_year
	view_sw = "0"
End If

be_year = Int(cost_year) - 1

'�˻� ��ȸ �⵵
For i = 1 To 15
	year_tab(i) = Int(Mid(CStr(Now()), 1, 4)) - i + 1
Next

'For i = 1 To 15
'	year_tab(i) = Int(Mid(CStr(Now()), 1, 4)) - i
'Next

For i = 1 To 20
	saupbu_tab(i) = ""
Next

For i = 1 To 20
	For j = 1 To 3
		For k = 1 To 13
			sum_amt(i,j,k) = 0
		Next
	Next
Next

' 2019.02.22 ������ ��û '����κ� �����Ѱ�'���� �ش�⵵�� ����θ� �����ϸ��
' �������� ����
objBuilder.Append "SELECT saupbu "
objBuilder.Append "FROM sales_org "
objBuilder.Append "WHERE sales_year='" & cost_year & "' "

If team <> "ȸ���繫" And user_id <> "102592" Then
    ' ȸ���繫 �϶��� ��Ÿ����ΰ� ������ ����..
    ' INSERT INTO `nkp`.`sales_org` (`saupbu`, `sort_seq`, `sales_year`) VALUES ('��Ÿ�����', '7', '2019');
	'sql = "  SELECT saupbu                         " & chr(13) & _
	'			"    FROM sales_org                      " & chr(13) & _
	'			"   WHERE sales_year='" & cost_year & "' " & chr(13) & _
	'			"ORDER BY sort_seq                       "
	'objBuilder.Append "ORDER BY sort_seq ASC "
Else
	'sql = "  SELECT saupbu                         " & chr(13) & _
	'			"    FROM sales_org                      " & chr(13) & _
	'			"   WHERE sales_year='" & cost_year & "' " & chr(13) & _
	'			"     AND saupbu not in ( '��Ÿ�����' ) " & chr(13) & _
	'			"ORDER BY sort_seq
	objBuilder.Append "	AND saupbu NOT IN('��Ÿ�����') "
	'objBuilder.Append "ORDER BY sort_seq ASC "
End If

objBuilder.Append "ORDER BY sort_seq ASC "

Set rsSalesDept = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

i = 0
Do Until rsSalesDept.EOF
	i = i + 1
	saupbu_tab(i) = rsSalesDept("saupbu")

	rsSalesDept.MoveNext()
Loop

rsSalesDept.Close() : Set rsSalesDept = Nothing

'---------------------------------------------------------------------------------------------------------------
'// 2017-09-15 ȸ���繫 ���� ��Ÿ�����,ȸ�簣�ŷ� ��ȸ �����ϰ� ����
'---------------------------------------------------------------------------------------------------------------

If team="ȸ���繫" Or user_id = "102592" Then
	i = i + 1
	saupbu_tab(i) = "��Ÿ�����"
	i = i + 1
	saupbu_tab(i) = "ȸ�簣�ŷ�"
'	i = i + 1
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
	objBuilder.Append "	AND (cost_center = 'ȸ�簣�ŷ�') "
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
'sql = "select substring(sales_date,1,7) as sales_month,saupbu,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,4) = '"&cost_year&"' group by substring(sales_date,1,7), saupbu"

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

' ��� ����
'sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' group by saupbu"

objBuilder.Append "SELECT saupbu, SUM(cost_amt_01), SUM(cost_amt_02), "
objBuilder.Append "SUM(cost_amt_03), SUM(cost_amt_04), SUM(cost_amt_05), "
objBuilder.Append "SUM(cost_amt_06), SUM(cost_amt_07), SUM(cost_amt_08), "
objBuilder.Append "SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "
objBuilder.Append "SUM(cost_amt_12) "
objBuilder.Append "FROM saupbu_profit_loss "
objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
objBuilder.Append "GROUP BY saupbu "

Set rsProfitStats = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsProfitStats.EOF
	For i = 1 To 20
		If saupbu_tab(i) = rsProfitStats("saupbu") Then
			j = 2

			For k = 1 To 12
				sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsProfitStats(k))
			Next

			Exit For
		End If
	Next

	rsProfitStats.MoveNext()
Loop

rsProfitStats.Close() : Set rsProfitStats = Nothing

' ��� ���� (��Ÿ�����)
'sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' and (saupbu = '' or saupbu = '��Ÿ�����') group by saupbu"

objBuilder.Append "SELECT saupbu, SUM(cost_amt_01), SUM(cost_amt_02), "
objBuilder.Append "	SUM(cost_amt_03), SUM(cost_amt_04), SUM(cost_amt_05), "
objBuilder.Append "	SUM(cost_amt_06), SUM(cost_amt_07), SUM(cost_amt_08), "
objBuilder.Append "	SUM(cost_amt_09), SUM(cost_amt_10), SUM(cost_amt_11), "
objBuilder.Append "	SUM(cost_amt_12) "
objBuilder.Append "FROM saupbu_profit_loss "
objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
objBuilder.Append "	AND (saupbu = '' OR saupbu = '��Ÿ�����') "
objBuilder.Append "GROUP BY saupbu "

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
				sum_amt(i, j, k) = sum_amt(i, j, k) + CDbl(rsEtcStats(k))
			Next

			Exit For
		End If
	Next

	rsEtcStats.MoveNext()
Loop

rsEtcStats.Close() : Set rsEtcStats = Nothing
DBConn.Close() : Set DBConn = Nothing

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

title_line = "����κ� ���� �Ѱ� ��Ȳ"
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
				var c_year = parseInt($('#cost_year').val());

				if(c_year > 2020){
					$('#frm').attr('action', '/sales/saupbu_profit_loss_total.asp').submit();
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
				<form action="/sales/old/saupbu_profit_loss_total_old.asp" method="post" name="frm" id = "frm">
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
									'For i = 1 To 5
									For i = 1 To 15
									%>
										<option value="<%=year_tab(i)%>" <%If CInt(cost_year) = CInt(year_tab(i)) Then%>selected <%End If %>>&nbsp;<%=year_tab(i)%></option>
									<%Next %>
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
							  <th class="first" scope="col">�����</th>
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
								<%If(k < 13 And sum_amt(i, 2, k) > 0) And (saupbu_tab(i) <> "ȸ�簣�ŷ�") Then %>
									<a href="#" onClick="pop_Window('/sales/old/saupbu_profit_loss_report2_old.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>&sales_saupbu=<%=saupbu_tab(i)%>','saupbu_profit_loss_report_pop','scrollbars=yes,width=1230,height=650')">
										<%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%>
									</a>
								<%Else %>
								<%	If(k < 13 And sum_amt(i, 2, k) > 0) And (saupbu_tab(i) = "ȸ�簣�ŷ�") Then %>
								<a href="#" onClick="pop_Window('/sales/old/company_deal_detail_view_old.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>','company_deal_detail_view_pop','scrollbars=yes,width=1000,height=600')">
									<%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%>
								</a>
								<% 	Else %>
									<%=FormatNumber(sum_amt(i, 2, k)/1000, 0)%>
								<%	End If	%>
								<%End If	%>
							  </td>
								<%
								Next
								%>
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
								<td class="right"><%=FormatNumber(sum_amt(0,2,k)/1000, 0)%></td>
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
                    <a href="/sales/excel/saupbu_profit_loss_total_excel_old.asp?cost_year=<%=cost_year%>" class="btnType04">�����ٿ�ε�</a>
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

