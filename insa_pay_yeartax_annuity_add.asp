<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
a_year = request("a_year")
a_emp_no = request("a_emp_no")
a_emp_name = request("a_emp_name")
a_seq = request("a_seq")

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

title_line = " ���ݺ���� �����׸� �Է� "
if u_type = "U" then

	Sql="select * from pay_yeartax_annuity where a_year = '"&a_year&"' and a_emp_no = '"&a_emp_no&"' and a_seq = '"&a_seq&"'"
	Set rs=DbConn.Execute(Sql)

	a_emp_name = rs("a_emp_name")
    a_type = rs("a_type")
    a_bank_code = rs("a_bank_code")
    a_bank_name = rs("a_bank_name")
    a_account_no = rs("a_account_no")
    a_amt = rs("a_amt")

	rs.close()

	title_line = " ���ݺ���� �����׸� ����  "
	
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
				if(document.frm.a_type.value =="") {
					alert('������ �Է��ϼ���');
					frm.a_type.focus();
					return false;}
				if(document.frm.a_bank_code.value =="") {
					alert('��������� �Է��ϼ���');
					frm.a_bank_code.focus();
					return false;}
				if(document.frm.a_account_no.value =="") {
					alert('����/���ǹ�ȣ�� �Է��ϼ���');
					frm.a_account_no.focus();
					return false;}
				if(document.frm.a_amt =="") {
					alert('�ݾ��� �����ϼ���');
					frm.a_amt.focus();
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
				aa_amt = parseInt(document.frm.a_amt.value.replace(/,/g,""));	
		
				aa_amt = String(aa_amt);
				num_len = aa_amt.length;
				sil_len = num_len;
				aa_amt = String(aa_amt);
				if (aa_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) aa_amt = aa_amt.substr(0,num_len -3) + "," + aa_amt.substr(num_len -3,3);
				if (sil_len > 6) aa_amt = aa_amt.substr(0,num_len -6) + "," + aa_amt.substr(num_len -6,3) + "," + aa_amt.substr(num_len -2,3);
				document.frm.a_amt.value = aa_amt;
			}		
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_annuity_save.asp" method="post" name="frm">
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
					  <input name="a_emp_no" type="text" id="a_emp_no" size="10" value="<%=a_emp_no%>" readonly="true">
                      <input type="hidden" name="a_year" value="<%=a_year%>" ID="b_year">
                      <input type="hidden" name="a_seq" value="<%=a_seq%>" ID="b_seq"></td>
                      <th style="background:#FFFFE6">����</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="a_emp_name" type="text" id="a_emp_name" size="10" value="<%=a_emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>����</th>
                      <td class="left">
                      <select name="a_type" id="a_type" value="<%=a_type%>" style="width:160px">
				          <option value="" <% if a_type = "" then %>selected<% end if %>>����</option>
				          <option value="Ȯ���ÿ�����������(DC)" <%If a_type = "Ȯ���ÿ�����������(DC)" then %>selected<% end if %>>Ȯ���ÿ�����������(DC)</option>
				          <option value="�������������ݰ���(IRP)" <%If a_type = "�������������ݰ���(IRP)" then %>selected<% end if %>>�������������ݰ���(IRP)</option>
				          <option value="��Ÿ" <%If a_type = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                      </select>
                      </td>
                      <th>�������</th>
					  <td class="left">
                      <input name="a_bank_code" type="text" value="<%=a_bank_code%>" readonly="true" style="width:40px">
                      <input name="a_bank_name" type="text" value="<%=a_bank_name%>" readonly="true" style="width:150px">
					  <a href="#" class="btnType03" onClick="pop_Window('insa_bank_select.asp?gubun=<%="yeara"%>&b_emp_no=<%=b_emp_no%>','stock_search_pop','scrollbars=yes,width=600,height=400')">ã��</a>
                      </td>
                    </tr>
                    <tr>
                      <th>����/���ǹ�ȣ</th>
					  <td class="left">
                      <input name="a_account_no" type="text" value="<%=a_account_no%>"  style="width:150px">
                      </td>
                      <th>�ݾ�</th>
					  <td class="left">
                      <input name="a_amt" type="text" id="a_amt" style="width:90px;text-align:right" value="<%=formatnumber(a_amt,0)%>" onKeyUp="num_chk(this);"></td>
                    <tr>
                      <td colspan="4" class="left">�� �������/���¹�ȣ or ���ǹ�ȣ�� ��Ȯ�ϰ� �Է�<br>
                �� ���������� Ȯ���⿩����������(DC)���� �������������ݰ���(IRP)�� ������ �߰��� ������ �ݾ׿� ���Ͽ� �Է�</td>
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

