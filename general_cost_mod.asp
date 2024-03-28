<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
slip_date = request("slip_date")
slip_seq = request("slip_seq")

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
confirm_yn = rs("confirm_yn")

title_line = "�Ϲݰ�� ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				if(document.frm.account.value =="") {
					alert('����׸� �����ϼ���');
					frm.account.focus();
					return false;}
				if(document.frm.price.value =="") {
					alert('����� �Է��ϼ���');
					frm.price.focus();
					return false;}
				if(document.frm.customer.value =="") {
					alert('��ȣ���� �Է��ϼ���');
					frm.customer.focus();
					return false;}
				if(document.frm.pay_yn.value =="N") {			
					k = 0;
					for (j=0;j<2;j++) {
						if (eval("document.frm.cancel_yn[" + j + "].checked")) {
							k = k + 1
						}
					}
					if (k==0) {
						alert ("��� ���θ� üũ�ϼ���");
						return false;
					}	
				}
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.confirm_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("Ȯ�� ���θ� üũ�ϼ���");
					return false;
				}	

				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="general_cost_mod_save.asp" method="post" name="frm">
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
				        <td class="left"><%=slip_date%></td>
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
				        <td class="left"><%=pay_method%>&nbsp;<input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(price,0)%>" onKeyUp="plusComma(this);" ></td>
				        <th>���ȸ��</th>
				        <td class="left"><%=company%></td>
			          </tr>
				      <tr>
				        <th class="first">��ȣ��(�����̸�)</th>
				        <td class="left"><input name="customer" type="text" id="customer" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=customer%>"></td>
				        <th>��볻��</th>
				        <td class="left"><%=slip_memo%></td>
			          </tr>
				      <tr>
				        <th class="first">��ҿ���</th>
				        <td class="left">
					<% if pay_yn = "Y" then	%>
                    	����Ǿ� ��� �Ұ�
                    <%   else	%>
                        <input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:30px" ID="Radio1">���
                        <input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:30px" ID="Radio2">����
					<% end if	%>
                        </td>
				        <th>Ȯ�ο���</th>
				        <td class="left">
                        <input type="radio" name="confirm_yn" value="Y" <% if confirm_yn = "Y" then %>checked<% end if %> style="width:30px" ID="Radio3">Ȯ��
  						<input type="radio" name="confirm_yn" value="N" <% if confirm_yn = "N" then %>checked<% end if %> style="width:30px" ID="Radio4">��Ȯ��
                        </td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
                    <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                    <input type="hidden" name="slip_date" value="<%=slip_date%>" ID="Hidden1">
                    <input type="hidden" name="slip_seq" value="<%=slip_seq%>" ID="Hidden1">
                    <input type="hidden" name="pay_yn" value="<%=pay_yn%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

