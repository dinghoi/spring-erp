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
pl_yn = "Y"
sign_no = ""
curr_date = mid(cstr(now()),1,10)

title_line = "���� �� �����μ� ��� ���"
if u_type = "U" then

	Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)

	emp_company = rs("emp_company")
	bonbu = rs("bonbu")
	saupbu = rs("saupbu")
	team = rs("team")
	org_name = rs("org_name")
	reside_place = rs("reside_place")
	account = rs("account") + "-" + rs("account_item")
	pay_method = rs("pay_method")
	price = rs("price")
	company = rs("company")
	vat_yn = rs("vat_yn")
	pay_yn = rs("pay_yn")
	customer = rs("customer")
	emp_name = rs("emp_name")
	emp_no = rs("emp_no")
	emp_grade = rs("emp_grade")
	sign_no = rs("sign_no")
	slip_memo = rs("slip_memo")
	end_yn = rs("end_yn")
	cancel_yn = rs("cancel_yn")
	reg_id = rs("reg_id")
	reg_date = rs("reg_date")
	reg_user = rs("reg_user")
	mod_id = rs("mod_id")
	mod_date = rs("mod_date")
	mod_user = rs("mod_user")
	pl_yn = rs("pl_yn")
	rs.close()

	title_line = "���� �� �����μ� ��� ����"
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				if(document.frm.price.value =="") {
					alert('����� �Է��ϼ���');
					frm.price.focus();
					return false;}
				if(document.frm.customer.value =="") {
					alert('�߻������� �Է��ϼ���');
					frm.customer.focus();
					return false;}
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
			var d = document.frm.cost_grade.value;
				if (c == 'U')
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
				if (d == '0')
				{
					document.getElementById('pl_col').style.display = '';
				}
			}

			function delcheck(){
				if(!confirm('���� �����Ͻðڽ��ϱ�?')){
					return false;
				}else{
					document.frm.action = "common_cost_del_ok.asp";
					document.frm.submit();
					return true;
				}
			}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="common_cost_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="15%" >
				      <col width="35%" >
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
				        <th>�������</th>
				        <td class="left"><% if cost_grade = "0" or saupbu = "�濵������" then	%>
                          <input name="org_name" type="text" value="<%=org_name%>" readonly="true" style="width:150px">
                          <a href="#" onClick="pop_Window('/insa/org_search.asp','org_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">������ȸ</a>
                          <%   else	%>
                          <%=org_name%>
                          <input name="org_name" type="hidden" value="<%=org_name%>">
                          <% end if	%>
                          <input name="bonbu" type="hidden" value="<%=bonbu%>">
                          <input name="saupbu" type="hidden" value="<%=saupbu%>">
                          <input name="team" type="hidden" value="<%=team%>">
                          <input name="reside_place" type="hidden" value="<%=reside_place%>">
                        <input name="reside_company" type="hidden" value="<%=reside_company%>"></td>
			          </tr>
				      <tr>
				        <th class="first">�����</th>
				        <td class="left"><%=emp_name%>&nbsp;(&nbsp;<%=emp_no%>&nbsp;)&nbsp;<%=emp_grade%></td>
				        <th>ȸ��</th>
				        <td class="left">
							<select name="emp_company" id="emp_company" style="width:120px">
                                <option value="" <% if emp_company = "" then %>selected<% end if %>>����</option>
                                <%
                                ' 2019.02.22 ������ ��û ȸ�縮��Ʈ�� ������ �ҽ� org_end_date�� null �� �ƴ� �������ڸ� �����ϸ� ����Ʈ�� ��Ÿ���� �ʴ´�.
                                'Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = 'ȸ��'  ORDER BY org_company ASC"
								sql = "SELECT org_company from emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') AND org_level = 'ȸ��' ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                                do until rs_org.eof
                                    %>
                                    <option value='<%=rs_org("org_company")%>' <%If emp_company = rs_org("org_company") then %>selected<% end if %>><%=rs_org("org_company")%></option>
                                    <%
                                    rs_org.movenext()
                                loop
                                rs_org.close()
                                %>
                            </select>
						</td>
			          </tr>
				      <tr>
				        <th class="first">����׸�</th>
				        <td class="left">
                            <select name="account" id="account" style="width:200px">
                            <option value="" <% if account = "" then %>selected<% end if %>>����</option>
                            <%
                                Sql="select * from account_item where cost_yn = 'Y' or cost_yn = 'C' order by account_name, account_item asc"
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
				        <th>��뱸��/�ݾ�</th>
				        <td class="left">
                            <select name="pay_method" id="pay_method" style="width:80px">
                                <option value='����' <%If pay_method = "����" then %>selected<% end if %>>����</option>
                            </select>
                            &nbsp;
                            <% if u_type = "U" then	%>
                            <input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(price,0)%>" onKeyUp="plusComma(this);" >
                            <% else	%>
                            <input name="price" type="text" id="price" style="width:100px;text-align:right" onKeyUp="plusComma(this);" >
                            <% end if %>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">����</th>
				        <td class="left">
						<%' if reside_company = "" or isnull(reside_company)	Then	%>
                            <input name="company" type="text" value="<%=company%>" readonly="true" style="width:150px">
                            <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="4"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
                        <%' else	%>
                        	<!--<input name="company" type="text" id="company" style="width:100px" value="<%'=reside_company%>" readonly="true" >-->
                        <%' end if	%>
                        </td>
				        <th>��ȣ��(�����̸�)</th>
				        <td class="left"><input name="customer" type="text" id="customer" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=customer%>"></td>
			          </tr>
				      <tr>
				        <th class="first">���꿩��</th>
				        <td class="left">
                        <input type="radio" name="pay_yn" value="N" <% if pay_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio3">������
				        <input type="radio" name="pay_yn" value="Y" <% if pay_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio4">����
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
                    <div align = "center">
                    <% if end_yn = "N" or end_yn = "C" then	%>
                        <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" /></span>
                    <% end if	%>
                        <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();" /></span>
                    <%' if u_type = "U" and user_id = emp_no And emp_no = "102592" then
						If (u_type = "U" And user_id = emp_no) Or (u_type = "U" And account_grade = "0") Then
                            if end_yn = "N" Or end_yn = "C" then
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
                    <input type="hidden" name="emp_name" value="<%=emp_name%>" />
                    <input type="hidden" name="emp_no" value="<%=emp_no%>" />
                    <input type="hidden" name="emp_grade" value="<%=emp_grade%>" />
                    <input type="hidden" name="sign_no" value="<%=sign_no%>" />
                    <input type="hidden" name="cancel_yn" value="<%=cancel_yn%>" />
                    <input type="hidden" name="mod_id" value="<%=mod_id%>" />
                    <input type="hidden" name="mod_user" value="<%=mod_user%>" />
                    <input type="hidden" name="mod_date" value="<%=mod_date%>" />
                    <input type="hidden" name="cost_grade" value="<%=cost_grade%>" />
				</form>
		</div>
	</body>
</html>

