<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
slip_date = request("slip_date")
slip_seq = request("slip_seq")

'Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
SQL = "SELECT gect.slip_gubun, gect.customer, gect.customer_no, gect.emp_company, gect.bonbu, gect.saupbu, gect.team,  "
SQL = SQL & "	gect.org_name, gect.company, gect.account, gect.account_item, gect.price, gect.cost, gect.cost_vat, "
SQL = SQL & "	gect.slip_memo, gect.emp_name, gect.emp_grade, gect.reg_id, gect.mg_saupbu, gect.pl_yn, gect.emp_no, "
SQL = SQL & "	gect.slip_date, gect.slip_seq, gect.approve_no "
'SQL = SQL & "	emtt.mg_saupbu AS mgSaupbu, eomt.org_name AS orgName, eomt.org_company, eomt.org_bonbu, "
'SQL = SQL & "	eomt.org_saupbu, eomt.org_team "
SQL = SQL & "FROM general_cost AS gect "
'SQL = SQL & "INNER JOIN emp_master AS emtt ON gect.emp_no = emtt.emp_no "
'SQL = SQL & "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
SQL = SQL & "WHERE slip_date = '"&slip_date&"' AND slip_seq = '"&slip_seq&"' "

'Response.write Sql
Set rs = DbConn.Execute(SQL)

slip_gubun = rs("slip_gubun")
customer = rs("customer")
customer_no = rs("customer_no")
emp_company = rs("emp_company")

Select Case emp_company
	Case "���̿��������" : emp_company = "���̿�"
	Case "�ڸ��Ƶ𿣾�" : emp_company = "���̽ý���"
End Select

bonbu = rs("bonbu")
saupbu = rs("saupbu")
team = rs("team")
org_name = rs("org_name")'
company = rs("company")
account = rs("account")
account_item = rs("account_item")
price = rs("price")
cost = rs("cost")
cost_vat = rs("cost_vat")
slip_memo = rs("slip_memo")
emp_no = rs("emp_no")
emp_name = rs("emp_name")
emp_grade = rs("emp_grade")
reg_id = rs("reg_id")
mg_saupbu = rs("mg_saupbu")
pl_yn = rs("pl_yn")

if slip_gubun = "���" then
	account_view = account + "-" + account_item
  else
  	account_view = account_item
end if

title_line = "���� ���ݰ�꼭 ����"
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			function chkfrm() {
//				if(document.frm.slip_date.value <= document.frm.end_date.value) {
//					alert('�������ڰ� ������ �Ǿ� �ִ� �����Դϴ�');
//					frm.slip_date.focus();
//					return false;}
				if(document.frm.mg_saupbu.value =="����") {
					alert('��翵������θ� �����ϼ���');
					frm.mg_saupbu.focus();
					return false;}
				if(document.frm.company.value =="") {
					alert('���縦 �����ϼ���');
					frm.company.focus();
					return false;}
				if(document.frm.slip_gubun.value =="") {
					alert('��������� �����ϼ���');
					frm.slip_gubun.focus();
					return false;}
				if(document.frm.company.value =="����" || document.frm.company.value =="���̿�") {
					if(document.frm.mg_saupbu.value != "") {
						//if(document.frm.mg_saupbu.value != document.frm.saupbu.value) {
						if(document.frm.mg_saupbu.value != document.frm.bonbu.value) {
							alert('���簡 ������ ��� �����������ο� ��翵������ΰ� �����ؾ��մϴ�.');
							frm.org_name.focus();
							return false;}}}

				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function pl_view() {
			var d = document.frm.cost_grade.value;
				if (d == '0')
				{
					document.getElementById('pl_col').style.display = '';
				}
			}
			function delcheck()
				{
				a=confirm('���� �����Ͻðڽ��ϱ�?')
				if (a==true) {
					document.frm.action = "/tax_bill_in_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body onload="pl_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/cost/tax_bill_in_mod_save.asp" method="post" name="frm">
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
				        <td class="left"><%=rs("slip_date")%>&nbsp;
				          ������ : <%=end_date%>
                        </td>
				        <th>���޹޴�ȸ��</th>
				        <td class="left"><%=emp_company%></td>
			          </tr>
				      <tr>
				        <th class="first">�������</th>
				        <td class="left">
                        <input name="org_name" type="text" readonly="true" value="<%=org_name%>" style="width:150px">
                        <%=emp_company%><a href="#" onClick="pop_Window('/org_search.asp?gubun=<%="��꼭"%>&org_company=<%=emp_company%>','org_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
				        <input name="emp_company" type="hidden" value="<%=emp_company%>">
				        <input name="bonbu" type="hidden" value="<%=bonbu%>">
				        <input name="saupbu" type="hidden" value="<%=saupbu%>">
				        <input name="team" type="hidden" value="<%=team%>">
				        <input name="reside_place" type="hidden" value="<%=reside_place%>">
                        <input name="reside_company" type="hidden" value="<%=reside_company%>">
                        </td>
				        <th>��翵�������</th>
				        <td class="left"><%
                                cost_year = mid(rs("slip_date"),1,4)
								sql_org = "select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
                                rs_org.Open sql_org, Dbconn, 1
                            %>
                          <select name="mg_saupbu" id="mg_saupbu" style="width:150px">
                            <option value='����' <%If mg_saupbu = "����" then %>selected<% end if %>>����</option>
                            <option value='' <%If mg_saupbu = "" then %>selected<% end if %>>��翵���ξ���</option>
                            <%
                                do until rs_org.eof
                            %>
                            <option value='<%=rs_org("saupbu")%>' <%If rs_org("saupbu") = mg_saupbu  then %>selected<% end if %>><%=rs_org("saupbu")%></option>
                            <%
                                    rs_org.movenext()
                                loop
                                rs_org.Close()
                            %>
                        </select></td>
			          </tr>
				      <tr>
				        <th class="first">������</th>
				        <td class="left"><%=mid(rs("customer_no"),1,3)%>-<%=mid(rs("customer_no"),4,2)%>-<%=right(rs("customer_no"),5)%>&nbsp;<%=rs("customer")%></td>
				        <th>�����</th>
				        <td class="left"><input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=emp_name%>" readonly="true">
                          <input name="emp_grade" type="text" id="emp_grade" style="width:60px" value="<%=emp_grade%>" readonly="true">
                        <a href="#" onClick="pop_Window('/insa/emp_search.asp?gubun=<%="1"%>','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">�����ȸ</a></td>
			          </tr>
				      <tr>
				        <th class="first">���೻��</th>
				        <td class="left"><input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,150);" value="<%=rs("slip_memo")%>"></td>
				        <th>�ݾ�</th>
				        <td class="left"><strong>���ް��� :</strong>&nbsp;<%=formatnumber(rs("cost"),0)%>&nbsp;&nbsp;&nbsp;<strong>�ΰ��� :</strong>&nbsp;<%=formatnumber(rs("cost_vat"),0)%></td>
			          </tr>
				      <tr>
				        <th class="first">����</th>
				        <td class="left">
                  <input name="company" type="text" value="<%=rs("company")%>" readonly="true" style="width:150px">
			            <a href="#" onClick="pop_Window('/trade_search.asp?gubun=<%="4"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
                </td>
				        <th>�������</th>
				        <td class="left">
									<input type="text" name="slip_gubun" ID="slip_gubun" readonly="true" style="width:100px" value="<%=rs("slip_gubun")%>">
									<input name="account_view" type="text" readonly="true" style="width:150px" value="<%=account_view%>">
                  <a href="#" onClick="pop_Window('/tax_bill_account_search.asp','tax_bill_account_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
									<input name="account" type="hidden" id="account" value="<%=rs("account")%>">
									<input name="account_item" type="hidden" id="account_item" value="<%=rs("account_item")%>">
                </td>
			          </tr>
				      <tr id="pl_col" style="display:none">
				        <th class="first">��������</th>
				        <td colspan="3" class="left">
									<input type="radio" name="pl_yn" value="Y" <% if pl_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio2">��������
									<input type="radio" name="pl_yn" value="N" <% if pl_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio">���͹�����
								</td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%'	if end_yn = "N" then	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%'	end if	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
				<%
					if (user_id = reg_id or user_id = emp_no) Or user_id = "102592" then
						if end_yn <> "Y" then
				%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
        		<%
						end if
					end if
				%>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
				<input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
				<input type="hidden" name="slip_date" value="<%=rs("slip_date")%>" ID="Hidden1">
				<input type="hidden" name="slip_seq" value="<%=rs("slip_seq")%>" ID="Hidden1">
				<input type="hidden" name="approve_no" value="<%=rs("approve_no")%>" ID="Hidden1">
				<input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="cost_grade" value="<%=cost_grade%>" ID="Hidden1">
			</form>
		</div>
	</body>
</html>

