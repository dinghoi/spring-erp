<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
s_year = request("s_year")
s_emp_no = request("s_emp_no")
s_emp_name = request("s_emp_name")
s_id = request("s_id")
s_seq = request("s_seq")

'response.write(s_id)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = s_id + " �����׸� �Է� "
if u_type = "U" then

	Sql="select * from pay_yeartax_saving where s_year = '"&s_year&"' and s_emp_no = '"&s_emp_no&"' and s_id = '"&s_id&"' and s_seq = '"&s_seq&"'"
	Set rs=DbConn.Execute(Sql)

	s_emp_name = rs("s_emp_name")
    s_type = rs("s_type")
    s_bank_code = rs("s_bank_code")
    s_bank_name = rs("s_bank_name")
    s_account_no = rs("s_account_no")
    s_amt = rs("s_amt")

	rs.close()

	title_line = s_id + " �����׸� ����  "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=b_from_date%>" );
			});	
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=b_to_date%>" );
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
				if(document.frm.s_type.value =="") {
					alert('������ �Է��ϼ���');
					frm.s_type.focus();
					return false;}
				if(document.frm.s_bank_code.value =="") {
					alert('��������� �Է��ϼ���');
					frm.s_bank_code.focus();
					return false;}
				if(document.frm.s_account_no.value =="") {
					alert('����/���ǹ�ȣ�� �Է��ϼ���');
					frm.s_account_no.focus();
					return false;}
				if(document.frm.s_amt =="") {
					alert('�ݾ��� �����ϼ���');
					frm.s_amt.focus();
					return false;}
			
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			} 
			
			function num_chk(txtObj){
				ss_amt = parseInt(document.frm.s_amt.value.replace(/,/g,""));	
		
				ss_amt = String(ss_amt);
				num_len = ss_amt.length;
				sil_len = num_len;
				ss_amt = String(ss_amt);
				if (ss_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) ss_amt = ss_amt.substr(0,num_len -3) + "," + ss_amt.substr(num_len -3,3);
				if (sil_len > 6) ss_amt = ss_amt.substr(0,num_len -6) + "," + ss_amt.substr(num_len -6,3) + "," + ss_amt.substr(num_len -2,3);
				document.frm.s_amt.value = ss_amt;
			}		
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_saving_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="15%" >
						<col width="25%" >
						<col width="15%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">���</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="s_emp_no" type="text" id="s_emp_no" size="10" value="<%=s_emp_no%>" readonly="true">
                      <input type="hidden" name="s_year" value="<%=s_year%>" ID="s_year">
                      <input type="hidden" name="s_seq" value="<%=s_seq%>" ID="s_seq"></td>
                      <th style="background:#FFFFE6">����</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="s_emp_name" type="text" id="s_emp_name" size="10" value="<%=s_emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>����</th>
                      <td class="left">
            <% if s_id = "��������" then  %>              
                      <select name="s_type" id="s_type" value="<%=s_type%>" style="width:160px">
				          <option value="" <% if s_type = "" then %>selected<% end if %>>����</option>
				          <option value="���ο�������(2000������)" <%If s_type = "���ο�������(2000������)" then %>selected<% end if %>>���ο�������(2000������)</option>
				          <option value="��������(2001������)" <%If s_type = "��������(2001������)" then %>selected<% end if %>>��������(2001������)</option>
				          <option value="�������ݼҵ����" <%If s_type = "�������ݼҵ����" then %>selected<% end if %>>�������ݼҵ����</option>
                      </select>
            <% end if %>	
            <% if s_id = "���ø�������" then  %> 
                      <select name="s_type" id="s_type" value="<%=s_type%>" style="width:160px">
				          <option value="" <% if s_type = "" then %>selected<% end if %>>����</option>
				          <option value="û������" <%If s_type = "û������" then %>selected<% end if %>>û������</option>
				          <option value="����û����������" <%If s_type = "����û����������" then %>selected<% end if %>>����û����������</option>
				          <option value="�ٷ������ø�������" <%If s_type = "�ٷ������ø�������" then %>selected<% end if %>>�ٷ������ø�������</option>
                          <option value="������ø�������" <%If s_type = "������ø�������" then %>selected<% end if %>>������ø�������</option>
                      </select>                 
            <% end if %>
            <% if s_id = "����ֽ�������" then  %>  
                      <select name="s_type" id="s_type" value="<%=s_type%>" style="width:160px">
				          <option value="" <% if s_type = "" then %>selected<% end if %>>����</option>
				          <option value="2����" <%If s_type = "2����" then %>selected<% end if %>>2����</option>
				          <option value="3����" <%If s_type = "3����" then %>selected<% end if %>>3����</option>
				          <option value="4����" <%If s_type = "4����" then %>selected<% end if %>>4����</option>
                      </select>                 
            <% end if %>	            
                      </td>
                      <th>�������</th>
					  <td class="left">
                      <input name="s_bank_code" type="text" value="<%=s_bank_code%>" readonly="true" style="width:40px">
                      <input name="s_bank_name" type="text" value="<%=s_bank_name%>" readonly="true" style="width:150px">
					  <a href="#" class="btnType03" onClick="pop_Window('insa_bank_select.asp?gubun=<%="saving"%>&s_emp_no=<%=s_emp_no%>','stock_search_pop','scrollbars=yes,width=600,height=400')">ã��</a>
                      </td>
                    </tr>
                    <tr>
                      <th>����/���ǹ�ȣ</th>
					  <td class="left">
                      <input name="s_account_no" type="text" value="<%=s_account_no%>"  style="width:150px">
                      </td>
                      <th>�ݾ�</th>
					  <td class="left">
                      <input name="s_amt" type="text" id="s_amt" style="width:90px;text-align:right" value="<%=formatnumber(s_amt,0)%>" onKeyUp="num_chk(this);"></td>
                    <tr>
            <% if s_id = "��������" then  %>           
                      <td colspan="4" class="left">�� ���θ��Ƿ� ���Ե� �������ุ �������<br>
                �� ������ �ѱݾ��� �Է�<br>
                �� ���������� ����/���ǹ�ȣ�� ��Ȯ�� �Է�</td>
            <% end if %>  
            <% if s_id = "���ø�������" then  %>  
                      <td colspan="4" class="left">�� �ݵ�� �����ֿ��� �������� ��<br>
                �� ������ �ѱݾ��� �Է�<br>
                �� ����/�������/���¹�ȣ/�ݾ��� ��Ȯ�� �Է�</td>
            <% end if %>  
            <% if s_id = "����ֽ�������" then  %> 
                      <td colspan="4" class="left">�� ����/�������/���¹�ȣ/�ݾ��� ��Ȯ�� �Է�<br>
                �� ������ �ѱݾ��� �Է�<br>
                �� </td> 
            <% end if %>   
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
                <input type="hidden" name="s_id" value="<%=s_id%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

