<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
emp_no = request("emp_no")
emp_name = request("emp_name")
person_no1 = request("emp_person1")
person_no2 = request("emp_person2")

emp_type = ""
emp_pay_type = ""
bank_code = ""
bank_name = ""
account_no = ""
account_holder = emp_name

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_no = rs_emp("emp_no")
		emp_name = rs_emp("emp_name")
		emp_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
		emp_team = rs_emp("emp_team")
		emp_org_code = rs_emp("emp_org_code")
		emp_org_name = rs_emp("emp_org_name")
   else
		emp_name = ""
		emp_company = ""
		emp_bonbu = ""
		emp_saupbu = ""
		emp_team = ""
		emp_org_code = ""
		emp_org_name = ""
end if
rs_emp.close()


title_line = " ���� ������� ��� "
if u_type = "U" then

	Sql="select * from pay_bank_account where emp_no = '"&emp_no&"'"
	Set rs=DbConn.Execute(Sql)

	emp_type = rs("emp_type")
    emp_pay_type = rs("emp_pay_type")
    bank_code = rs("bank_code")
	person_no1 = rs("person_no1")
	person_no2 = rs("person_no2")
    bank_name = rs("bank_name")
    account_no = rs("account_no")
	account_holder = rs("account_holder")
	
	rs.close()

	title_line = " ���� ������� ���� "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ�޿� �ý���</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=family_birthday%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.bank_name.value =="") {
					alert('������� �����ϼ���');
					frm.bank_name.focus();
					return false;}
				if(document.frm.account_no =="") {
					alert('���¹�ȣ�� �����ϼ���');
					frm.account_no.focus();
					return false;}
				if(document.frm.account_holder.value =="") {
					alert('�����ָ� �Է��ϼ���');
					frm.account_holder.focus();
					return false;}
				
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
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_bank_account_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">���</th>
                      <td class="left" bgcolor="#FFFFE6"><%=emp_no%></td>
					  <input name="emp_no" type="hidden" id="emp_no" size="14" value="<%=emp_no%>" readonly="true"></td>
                      <th style="background:#FFFFE6">����</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6"><%=emp_name%></td>
					  <input name="emp_name" type="hidden" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th style="background:#FFFFE6">�ֹε��<br>��ȣ</th>
                      <td colspan="5" class="left" bgcolor="#FFFFE6"><%=person_no1%> - <%=person_no2%></td>
					  <input name="person_no1" type="hidden" id="person_no1" size="14" value="<%=person_no1%>" readonly="true">
                      <input name="person_no2" type="hidden" id="person_no2" size="14" value="<%=person_no2%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>�����</th>
                      <td colspan="5" class="left">
					<%
					  Sql="select * from emp_etc_code where emp_etc_type = '50' order by emp_etc_code asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
					  <select name="bank_name" id="bank_name" style="width:130px">
                         <option value="" <% if bank_name = "" then %>selected<% end if %>>����</option>
                			  <% 
								do until rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If bank_name = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							  %>
            		  </select>                 
                      </td>
                    </tr>
                    <tr>
                      <th>���¹�ȣ</th>
                      <td colspan="5" class="left">
					  <input name="account_no" type="text" id="account_no" size="20" value="<%=account_no%>"></td>
                    </tr>
                    <tr>
                      <th>������</th>
                      <td colspan="5" class="left">
					  <input name="account_holder" type="text" id="account_holder" size="14" value="<%=account_holder%>"></td>
					  </td>
			    	</tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

