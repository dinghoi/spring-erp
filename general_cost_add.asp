<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
u_type = request("u_type")
slip_date = request("slip_date")
slip_seq = request("slip_seq")

slip_gubun = ""
account = ""
sign_no = ""
pay_method = ""
price = 0
vat_yn = "N"
pay_yn = "N"
company = "����"
customer = ""
emp_name = user_name
emp_no = user_id
emp_grade = user_grade
slip_memo = ""
end_yn = "N"
cancel_yn = "N"
curr_date = mid(cstr(now()),1,10)

title_line = "�Ϲݰ�� ���"
if u_type = "U" then

	Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)

	org_name = rs("org_name")
	account = rs("account") + "-" + rs("account_item")
	sign_no = rs("sign_no")
	pay_method = rs("pay_method")
	price = rs("price")
	company = rs("company")
	vat_yn = rs("vat_yn")
	pay_yn = rs("pay_yn")
	customer = rs("customer")
	emp_name = rs("emp_name")
	emp_no = rs("emp_no")
	emp_grade = rs("emp_grade")
	slip_memo = rs("slip_memo")
	end_yn = rs("end_yn")
	cancel_yn = rs("cancel_yn")
	reg_id = rs("reg_id")
	reg_date = rs("reg_date")
	reg_user = rs("reg_user")
	mod_id = rs("mod_id")
	mod_date = rs("mod_date")
	mod_user = rs("mod_user")
	rs.close()

	title_line = "�Ϲݰ�� ����"
end if
if end_yn = "Y" then
	end_view = "����"
  else
  	end_view = "����"
end if
if cancel_yn = "Y" then
	cancel_view = "���"
  else
  	cancel_view = "����"
end If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=slip_date%>" );
			});
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
				if(document.frm.slip_date.value <= document.frm.end_date.value) {
					alert('�߻����ڰ� ������ �Ǿ� �ִ� �����Դϴ�');
					frm.slip_date.focus();
					return false;}
				if(document.frm.slip_date.value > document.frm.curr_date.value) {
					alert('�߻����ڰ� �����Ϻ��� Ŭ���� �����ϴ�.');
					frm.slip_date.focus();
					return false;}
				if(document.frm.end_yn.value =="Y") {
					alert('�����Ǿ� ���� �� �� �����ϴ�');
					frm.end_yn.focus();
					return false;}
				if(document.frm.slip_date.value =="") {
					alert('�߻����ڸ� �Է��ϼ���');
					frm.slip_date.focus();
					return false;}
				if(document.frm.account.value =="") {
					alert('����׸� �����ϼ���');
					frm.account.focus();
					return false;}
				if(document.frm.pay_method.value =="") {
					alert('��뱸�� �����ϼ���');
					frm.pay_method.focus();
					return false;}
				if(document.frm.price.value =="") {
					alert('����� �Է��ϼ���');
					frm.price.focus();
					return false;}
				if(document.frm.sign_no.value =="") {
					alert('���ڰ����ȣ�� �Է��ϼ���');
					frm.sign_no.focus();
					return false;}
				if(document.frm.customer.value =="") {
					alert('�߻������� �Է��ϼ���');
					frm.customer.focus();
					return false;}
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.pay_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("���� ���θ� üũ�ϼ���");
					return false;
				}
				if(document.frm.slip_memo.value =="") {
					alert('��� �Է��ϼ���');
					frm.slip_memo.focus();
					return false;}

				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U')
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
			function delcheck()
				{
				a=confirm('���� �����Ͻðڽ��ϱ�?')
				if (a==true) {
					document.frm.action = "general_cost_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="general_cost_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="15%" >
				      <col width="37%" >
				      <col width="15%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">�߻�����</th>
				        <td class="left">
                        <input name="slip_date" type="text" id="datepicker" style="width:80px;text-align:center" value="<%=slip_date%>" readonly="true">
				          �������� : <%=end_date%>
				        <input name="curr_date" type="hidden" value="<%=curr_date%>">
				        <input name="slip_seq" type="hidden" value="<%=slip_seq%>">
                        </td>
				        <th>�Ҽ�</th>
				        <td class="left"><%=org_name%></td>
			          </tr>
				      <tr>
				        <th class="first">�����</th>
				        <td class="left"><%=emp_name%>&nbsp;(&nbsp;<%=emp_no%>&nbsp;)&nbsp;<%=emp_grade%></td>
				        <th>����׸�</th>
				        <td class="left">
                        <select name="account" id="account" style="width:200px">
		                	<option value="" <% if account = "" then %>selected<% end if %>>����</option>
				            <%
                                    Sql="select * from account_item where cost_yn = 'Y' order by account_name, account_item asc"
                                    rs_acc.Open Sql, Dbconn, 1
                                    do until rs_acc.eof
										account_item = rs_acc("account_name") + "-" + rs_acc("account_item")
								  %>
				            <option value='<%=account_item%>' <%If account_item = account then %>selected<% end if %>><%=account_item%></option>
				            <%
                                        rs_acc.movenext()
                                    loop
                                    rs_acc.close()
                                  %>
		                </select>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">��뱸��/�ݾ�</th>
				        <td class="left">
                        <select name="pay_method" id="pay_method" style="width:80px">
				          <option value='����' <%If pay_method = "����" then %>selected<% end if %>>����</option>
				        </select>
                        &nbsp;
					<% if u_type = "U" then	%>
                        <input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(price,0)%>" onKeyUp="plusComma(this);" >
					<%   else	%>
                        <input name="price" type="text" id="price" style="width:100px;text-align:right" onKeyUp="plusComma(this);" >
                    <% end if	%>
                        </td>
				        <th>���ȸ��</th>
				        <td class="left">
							<input name="company" type="text" value="<%=company%>" readonly="true" style="width:150px">
                            <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="4"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">��ȣ��(�����̸�)</th>
				        <td class="left"><input name="customer" type="text" id="customer" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=customer%>"></td>
				        <th>���ڰ���NO</th>
				        <td class="left"><input name="sign_no" type="text" id="sign_no" style="width:40px" onKeyUp="checkNum(this);" value="<%=sign_no%>" maxlength="5"> *����4�ڸ��� �Է� ����</td>
			          </tr>
				      <tr>
				        <th class="first">���꿩��</th>
				        <td class="left">
                        <input type="radio" name="pay_yn" value="N" <% if pay_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio3">
				          ������
				        <input type="radio" name="pay_yn" value="Y" <% if pay_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio4">
				            ����
                        </td>
				        <th>��볻��</th>
				        <td class="left"><input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=slip_memo%>"></td>
			          </tr>
    				  <tr id="cancel_col" style="display:none">
						<th class="first">��ҿ���</th>
						<td class="left"><%=cancel_view%></td>
                        <th>��������</th>
						<td class="left"><%=end_view%></td>
					</tr>
					<tr id="info_col" style="display:none">
						<th class="first">�������</th>
						<td class="left"><%=reg_user%>&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
                    	<th>��������</th>
						<td class="left"><%=mod_user%>&nbsp;<%=mod_id%>(<%=mod_date%>)</td>
					</tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align = "center">
				<%	if end_yn = "N" or end_yn = "C" then	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" /></span>
        		<%	end if	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();" /></span>
				<%
					if u_type = "U" and user_id = emp_no then
						if end_yn = "N" or end_yn = "C" then
				%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();" /></span>
        		<%
						end if
					end if
				%>
                </div>
                    <input type="hidden" name="u_type" value="<%=u_type%>" />
                    <input type="hidden" name="end_yn" value="<%=end_yn%>" />
                    <input type="hidden" name="end_date" value="<%=end_date%>" />
                    <input type="hidden" name="old_date" value="<%=slip_date%>" />
                    <input type="hidden" name="cancel_yn" value="<%=cancel_yn%>" />
                    <input type="hidden" name="mod_id" value="<%=mod_id%>" />
                    <input type="hidden" name="mod_user" value="<%=mod_user%>" />
                    <input type="hidden" name="mod_date" value="<%=mod_date%>" />
				</form>
		</div>
	</body>
</html>

