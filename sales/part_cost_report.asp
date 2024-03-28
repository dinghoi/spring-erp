<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'on Error resume next
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
Dim cost_month, sales_saupbu, before_date
Dim condi_sql, mm, cost_year
Dim rsComCost, tot_part_cost
Dim from_date, end_date, to_date
Dim rsAsTot, tot_part_cnt
Dim title_line

cost_month = Request.Form("cost_month")
sales_saupbu = Request.Form("sales_saupbu")

If sales_saupbu = "" Then
	sales_saupbu = "��ü"
End If

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date),6,2)
	sales_saupbu = "��ü"
End If

from_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

mm = Mid(cost_month, 5, 2)
cost_year = Mid(cost_month, 1, 4)

'�ι������ ��ü ���
'sql = "SELECT SUM(cost_amt_"&mm&") AS tot_cost FROM company_cost WHERE cost_year ='"&cost_year&"' AND cost_center = '�ι������'"
'Set rs = DbConn.Execute(SQL)
objBuilder.Append "SELECT SUM(cost_amt_"& mm &") AS tot_cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_year ='"& cost_year &"' "
objBuilder.Append "AND cost_center = '�ι������' "

Set rsComCost = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsComCost("tot_cost")) Then
	tot_part_cost = 0
Else
	tot_part_cost = CLng(rsComCost("tot_cost"))
End If

rsComCost.Close() : Set rsComCost = Nothing

'If sales_saupbu = "��ü" Then
'	condi_sql = ""
'Else
'  	condi_sql = " AND trat.saupbu ='"& sales_saupbu &"' "
'End If

'A/S ��ü ī��Ʈ
'objBuilder.Append "SELECT COUNT(*) AS tot_cnt "
'objBuilder.Append "FROM as_acpt_end AS asat "
'objBuilder.Append "INNER JOIN emp_master_month AS emmt ON asat.mg_ce_id = emmt.emp_no "
'objBuilder.Append "	AND emmt.emp_month = '"&cost_month&"' "
'objBuilder.Append "INNER JOIN trade AS trat ON asat.company = trat.trade_name "
'objBuilder.Append "WHERE asat.as_type NOT IN ('����ó��', '��Ư��')"
'objBuilder.Append "	AND asat.as_process <> '���'"
'objBuilder.Append "	AND asat.reside = '0'"
'objBuilder.Append "	AND asat.reside_place = ''"
'objBuilder.Append "	AND (CAST(asat.visit_date AS DATE) >= '"&from_date&"' AND CAST(asat.visit_date AS DATE) <= '"&to_date&"') "
'objBuilder.Append "	AND emmt.cost_center = '�ι������' "
'objBuilder.Append condi_sql

objBuilder.Append "SELECT SUM(as_total) AS tot_cnt "
objBuilder.Append "FROM as_acpt_status "
objBuilder.Append "WHERE as_month = '"&cost_month&"' " & condi_sql

Set rsAsTot = DBconn.Execute(objBuilder.ToString())
objBuilder.Clear()

tot_part_cnt = f_toString(rsAsTot("tot_cnt"), 0)

rsAsTot.Close() : Set rsAsTot = Nothing

Dim rsPart, part_tot_cost, as_tot_cnt, set_tot_cost
'�ι������(���) - ��ġ���� ����
objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_detail = '��ġ����')) AS 'part_tot_cost', "
objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&cost_year&mm&"') AS 'as_tot_cnt', "
objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_detail = '��ġ����') AS 'set_tot_cost' "
objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '�ι������' "

Set rsPart = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

part_tot_cost = CDbl(f_toString(rsPart("part_tot_cost"), 0))	'�ι������(���)
as_tot_cnt = CInt(f_toString(rsPart("as_tot_cnt"), 0))	'AS �� �Ǽ�
set_tot_cost = CDbl(f_toString(rsPart("set_tot_cost"), 0))	'�� ��ġ���� ���

rsPart.Close() : Set rsPart = Nothing

'A/S ���ó�� ��Ȳ
objBuilder.Append "SELECT as_company, as_set, set_time, as_error, as_testing, as_collect, total_time, "
objBuilder.Append "	trade_name, saupbu, "
'objBuilder.Append	tot_part_cost&" / "&tot_part_cnt&" * as_total AS 'as_cost' "
objBuilder.Append	part_tot_cost&" / "&as_tot_cnt&" * (as_total - as_set) AS 'as_cost', "
objBuilder.Append set_tot_cost&" * set_time / (SELECT SUM(set_time) FROM as_acpt_status WHERE as_month = '"&cost_month&"' AND set_time > 0) AS 'as_set_cost' "
objBuilder.Append "FROM as_acpt_status AS aast "
objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
objBuilder.Append "	AND trdt.trade_id = '����' "
objBuilder.Append "WHERE as_month = '"&cost_month&"' "' & condi_sql

'Response.write objBuilder.ToString()
'Response.end

Set rsComCost = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "A/S ���ó�� ��Ȳ"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
		<!-- <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" /> -->
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
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("�߻������ �Է��ϼ���.");
					return false;
				}
				return true;
			}

			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<!--<h3 class="stit">����ó���� 5%, ���ݿܴ� 95% �������� ������ ��α����Դϴ�. </h3>-->
				<h3 class="stit">1. �ι������(���) = �ι������ �հ� - ��ġ���� ����</h3>
				<form action="/sales/part_cost_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>�߻����&nbsp;</strong>(��201401) :
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>

                                <!--<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>-->
								<img src="/image/but_ser.jpg" onclick="frmcheck();" style="cursor:pointer;" alt="�˻�">
                            </p>
						</dd>
					</dl>
				</fieldset>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
                    	<td>
      			<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
							<col width="10%" >
							<col width="10%" >
							<col width="8%" >
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">�ŷ�ó��</th>
								<!--<th scope="col">��ġ/����</th>-->
								<th scope="col">��ġ/����(�ð�)</th>
								<th scope="col">��ġ/�����</th>
								<th scope="col">���</th>
								<th scope="col">��������</th>
								<th scope="col">���ȸ��</th>
								<th scope="col">��������</th>
								<th scope="col">�ι������(���)</th>
								<th scope="col"></th>
							</tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="2%" >
						</colgroup>
						<tbody>
						<%
						Dim i, set_sum, set_time_sum, error_sum, testing_sum, collect_sum, total_time_sum, as_cost_sum, as_set_cost
						Dim as_set_sum

						set_sum = 0
						set_time_sum = 0
						error_sum = 0
						testing_sum = 0
						collect_sum = 0
						total_time_sum = 0
						as_cost_sum = 0
						as_set_sum = 0

						i = 0

						Do Until rsComCost.EOF
							i = i + 1

							'�׸� �� �� �հ�
							set_sum = set_sum + CLng(rsComCost("as_set"))
							set_time_sum = set_time_sum + CLng(rsComCost("set_time"))
							error_sum = error_sum + CLng(rsComCost("as_error"))
							testing_sum = testing_sum + CLng(rsComCost("as_testing"))
							collect_sum = collect_sum + CLng(rsComCost("as_collect"))
							total_time_sum = total_time_sum + CLng(rsComCost("total_time"))

							as_cost_sum = as_cost_sum + CDbl(rsComCost("as_cost"))
							as_set_sum = as_set_sum + CDbl(rsComCost("as_set_cost"))
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rsComCost("as_company")%></td>
								<!--<td class="right"><%=FormatNumber(rsComCost("as_set"), 0)%>&nbsp;</td>-->
								<td class="right"><%=FormatNumber(rsComCost("set_time"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("as_set_cost"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("as_error"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("as_testing"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("as_collect"), 0)%>&nbsp;</td>
								<td><%=rsComCost("saupbu")%></td>
								<td class="right"><%=FormatNumber(rsComCost("as_cost"), 0)%>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
							rsComCost.MoveNext()
						Loop
						rsComCost.Close() : Set rsComCost = Nothing
						DBConn.Close() : Set DBConn = Nothing

						Dim dist_part, dist_cost, part_cost

						If i > 0 Then
							'��ġ���� ���� = ��ġ���� �� �ð� / �� �ð� * 100
							'dist_part = FormatNumber(set_time_sum / total_time_sum * 100, 1)

							'��ġ���� ����  = �� �ι������ * ��ġ���� ����
							'dist_cost = FormatNumber(as_cost_sum * dist_part / 100, 0)
							dist_cost = FormatNumber(set_tot_cost, 0)

							'�ι������(���) = �ι������ �հ� - ��ġ���� ����
							'part_cost = FormatNumber(as_cost_sum - dist_cost, 0)
							part_cost = FormatNumber(tot_part_cost - dist_cost, 0)
						Else
							dist_part = 0
							dist_cost = 0
							part_cost = 0
						End If
						%>
							<tr>
								<td colspan="2" bgcolor="#FFE8E8" class="first">�Ѱ�</td>
								<!--<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(set_sum, 0)%>&nbsp;��</td>-->
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(set_time_sum, 0)%>&nbsp;�ð�</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(as_set_sum, 0)%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(error_sum, 0)%>&nbsp;��</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(testing_sum, 0)%>&nbsp;��</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(collect_sum, 0)%>&nbsp;��</td>
								<td colspan="2" bgcolor="#FFE8E8" class="right">
									<div style="font-weight:bold;">��ġ���� ���� : <%=dist_cost%></div>
									<div style="font-weight:bold;">�ι������(���) : <%=part_cost%></div>
								</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
							</tr>
						</tbody>
						</table>
                        </DIV>
						</td>
                    </tr>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="10%">
						<div class="btnCenter">
							<a href="/sales/part_cost_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">�����ٿ�ε�</a>
						</div>
                    </td>
				    <td width="90%">
                    </td>
			      </tr>
				  </table>
			</form>
				<br>
		</div>
	</div>
	</body>
</html>
