<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--include virtual="/include/db_create.asp" -->
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
Dim approve_no, emp_name, emp_grade, account, end_yn, curr_date, mg_saupbu
Dim rsTrade, title_line, company, u_type, end_date, rsSales

Dim sql, rs, rsOrg

approve_no = f_Request("approve_no")

emp_no = user_id
emp_name = user_name
emp_grade = user_grade
account = ""
end_yn = "N"
curr_date = Mid(CStr(Now()), 1, 10)

'������� ����
'objBuilder.Append "SELECT org_name FROM emp_master AS emtt "
'objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
'objBuilder.Append "WHERE emtt.emp_no = '"&emp_no&"' "

'Set rsOrg = DBConn.Execute(objBuilder.ToString())
'objBuilder.Clear()

'org_name = rsOrg("org_name")

'rsOrg.Close() : Set rsOrg = Nothing

'Sql="select * from tax_bill where approve_no = '"&approve_no&"'"
'Set rs = DbConn.Execute(Sql)

'Sql="select * from trade where trade_no = '"&rs("trade_no")&"'"
'Set rs_trade=DbConn.Execute(Sql)
'if rs_trade.eof or rs_trade.bof then
'	customer = rs("trade_name")
'  else
'	customer = rs_trade("trade_name")
'end If

'mg_saupbu = "����"

'������� ����
objBuilder.Append "SELECT saupbu FROM sales_org "
objBuilder.Append "WHERE sales_year = '"&Year(Now())&"' AND saupbu = '"&bonbu&"' "

Set rsSales = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsSales.EOF Or rsSales.BOF Then
	If bonbu = "�濵����" Then
		mg_saupbu = ""
	Else
		mg_saupbu = "����"
	End If
Else
	mg_saupbu = bonbu
End If

rsSales.Close() : Set rsSales = Nothing

Dim trade_name, owner_company, bill_date, cost_year, trade_no, tax_bill_memo, cost, cost_vat

objBuilder.Append "SELECT trade_name, owner_company, bill_date, trade_no, tax_bill_memo, cost, cost_vat "
objBuilder.Append "FROM tax_bill "
objBuilder.Append "WHERE approve_no = '"&approve_no&"' "

Set rsTrade = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

trade_name = rsTrade("trade_name")
owner_company = rsTrade("owner_company")
bill_date = rsTrade("bill_date")
trade_no = rsTrade("trade_no")
tax_bill_memo = rsTrade("tax_bill_memo")
cost = rsTrade("cost")
cost_vat = rsTrade("cost_vat")

Select Case owner_company
	Case "���̿��������" : owner_company = "���̿�"
	Case "�ڸ��Ƶ𿣾�" : owner_company = "���̽ý���"
	'Case Else
	'	owner_company = rs("owner_company")
End Select

cost_year = Mid(bill_date, 1, 4)

rsTrade.Close() : Set rsTrade = Nothing

title_line = "E���� ���� ���ݰ�꼭 ���κ�� �Է�"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>����ȸ��ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	    <script src="/java/jquery-1.9.1.js"></script>
	    <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
//				if(document.frm.bill_date.value <= document.frm.end_date.value) {
//					alert('�������ڰ� ������ �Ǿ� �ִ� �����Դϴ�');
//					frm.slip_date.focus();
//					return false;}
				if(document.frm.mg_saupbu.value =="����"){
					alert('��翵������θ� �����ϼ���');

					frm.mg_saupbu.focus();
					return false;
				}

				if(document.frm.company.value ==""){
					alert('���縦 �����ϼ���');

					frm.company.focus();
					return false;
				}

				if(document.frm.slip_gubun.value ==""){
					alert('��������� �����ϼ���');

					frm.slip_gubun.focus();
					return false;
				}

				{
					a=confirm('�Է��Ͻðڽ��ϱ�?');

					if(a==true){
						return true;
					}
				return false;
				}
			}
        </script>
	</head>
	<body>
		<div id="container">
			<h3 class="tit"><%=title_line%></h3>
			<form action="/cost/tax_esero_in_detail_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="13%" >
				      <col width="37%" >
				      <col width="13%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">��������</th>
				        <td class="left"><%=bill_date%></td>
				        <th>���޹޴�ȸ��</th>
				        <td class="left"><%=owner_company%></td>
			          </tr>
				      <tr>
				        <th class="first">�������</th>
				        <td class="left">
						<%
						'If cost_grade = "0" Or saupbu = "�濵������" Or team = "SM1��" Or team = "SM2��" Or team = "Repair��" Or user_name = "��ȣ��" Or user_id = "101756" Then
						'��ȣ��
						If cost_grade = "0" Or user_id = "100545" Then
						%>
                            <input name="org_name" type="text" readonly="true" value="<%=org_name%>" style="width:150px">
                            <a href="#" onClick="pop_Window('/org_search.asp?gubun=��꼭&org_company=<%=owner_company%>','org_search_pop','scrollbars=yes,width=600,height=400');" class="btnType03">��ȸ</a>
                        <%
						Else
						%>
                            <%=org_name%>
                            <input name="org_name" type="hidden" value="<%=org_name%>">
                        <%
						End If
						%>
							<input name="emp_company" type="hidden" value="<%=emp_company%>">
							<input name="bonbu" type="hidden" value="<%=bonbu%>">
							<input name="saupbu" type="hidden" value="<%=saupbu%>">
							<input name="team" type="hidden" value="<%=team%>">
							<input name="reside_place" type="hidden" value="<%=reside_place%>">
							<input name="reside_company" type="hidden" value="<%=reside_company%>">
                        </td>
				        <th>��翵�������</th>
				        <td class="left">
						<%
						Dim sql_org, rs_org
						'sql_org = "SELECT saupbu FROM sales_org WHERE sales_year='"&cost_year&"' ORDER BY sort_seq"
						objBuilder.Append "SELECT saupbu FROM sales_org "
						objBuilder.Append "WHERE sales_year='"&cost_year&"' "

						If user_id <> "100359" And SysAdminYn = "N" Then
							objBuilder.Append "AND saupbu <> '��Ÿ�����' "
						End If

						objBuilder.Append "ORDER BY sort_seq "

						Set rs_org = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()
						%>
                            <select name="mg_saupbu" id="mg_saupbu" style="width:150px">
                                <option value='����' <%If mg_saupbu = "����" Then %>selected<% End If %>>����</option>
                                <option value='' <%If mg_saupbu = "" Then %>selected<% End If %>>��翵���ξ���</option>
						<%
						Do Until rs_org.EOF
						%>
								<option value='<%=rs_org("saupbu")%>' <%If rs_org("saupbu") = mg_saupbu Then %>selected<% End If %>><%=rs_org("saupbu")%></option>
						<%
							rs_org.MoveNext()
						Loop
						rs_org.Close() : Set rs_org = Nothing

						DBConn.Close() : Set DBConn = Nothing
						%>
                            </select>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">������</th>
				        <td class="left"><%=Mid(trade_no, 1, 3)%>-<%=Mid (trade_no, 4, 2)%>-<%=Right(trade_no, 5)%>&nbsp;<%=trade_name%></td>
				        <th>�����</th>
				        <td class="left">
							<input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=emp_name%>" readonly="true">
							<input name="emp_grade" type="text" id="emp_grade" style="width:60px" value="<%=emp_grade%>" readonly="true">
							<a href="#" onClick="pop_Window('/insa/emp_search.asp?gubun=1','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">�����ȸ</a>
						</td>
			          </tr>
				      <tr>
				        <th class="first">���೻��</th>
				        <td class="left">
							<input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=tax_bill_memo%>">
						</td>
				        <th>�ݾ�</th>
				        <td class="left">
							<strong>���ް��� :</strong>&nbsp;<%=FormatNumber(cost, 0)%>&nbsp;&nbsp;&nbsp;
							<strong>�ΰ��� :</strong>&nbsp;<%=FormatNumber(cost_vat, 0)%>
						</td>
			          </tr>
				      <tr>
				        <th class="first">����</th>
				        <td class="left">
							<input name="company" type="text" value="<%=company%>" readonly="true" style="width:150px">
							<a href="#" onClick="pop_Window('/trade_search.asp?gubun=4','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
                        </td>
				        <th>�������</th>
				        <td class="left">
							<input type="text" name="slip_gubun" ID="slip_gubun" readonly="true" style="width:100px">
							<input name="account_view" type="text" readonly="true" style="width:150px">
							<a href="#" onClick="pop_Window('/tax_bill_account_search.asp','tax_bill_account_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
							<input name="account" type="hidden" id="account">
							<input name="account_item" type="hidden" id="account_item">
                        </td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align="center">
                    <% If end_yn = "N" Then	%>
                        <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" /></span>
                    <% End If %>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" />
				<input type="hidden" name="end_yn" value="<%=end_yn%>" />
				<input type="hidden" name="end_date" value="<%=end_date%>" />
				<input type="hidden" name="bill_date" value="<%=bill_date%>" />
				<input type="hidden" name="emp_no" value="<%=emp_no%>" />
				<input type="hidden" name="approve_no" value="<%=approve_no%>" />
			</form>
		</div>
	</body>
</html>