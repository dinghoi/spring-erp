<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
qual_empno = request("qual_empno")
qual_seq = request("qual_seq")
emp_name = request("emp_name")

qual_type = ""
qual_grade = ""
qual_pass_date = ""
qual_org = ""
qual_no = ""
qual_passport = ""
qual_pay_id = "N"

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " �ڰݻ��� ��� "

if u_type = "U" then

	Sql="select * from emp_qual where qual_empno = '"&qual_empno&"' and qual_seq = '"&qual_seq&"'"
	Set rs=DbConn.Execute(Sql)

	qual_empno = rs("qual_empno")
    qual_seq = rs("qual_seq")
	
	qual_type = rs("qual_type")
    qual_grade = rs("qual_grade")
    qual_pass_date = rs("qual_pass_date")
    qual_org = rs("qual_org")
    qual_no = rs("qual_no")
	qual_passport = rs("qual_passport")
	qual_pay_id = rs("qual_pay_id")

	rs.close()

	title_line = " �ڰݻ��� ���� "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=qual_pass_date%>" );
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
				if(document.frm.qual_type.value =="") {
					alert('�ڰ������� �Է��ϼ���');
					frm.qual_type.focus();
					return false;}
				if(document.frm.qual_org =="") {
					alert('�߱ޱ���� �����ϼ���');
					frm.qual_org.focus();
					return false;}
				if(document.frm.qual_no.value =="") {
					alert('�ڰݵ�Ϲ�ȣ�� �Է��ϼ���');
					frm.qual_no.focus();
					return false;}
				if(document.frm.qual_pass_date.value =="") {
					alert('�հݳ���ϸ� �Է��ϼ���');
					frm.qual_pass_date.focus();
					return false;}
				if(document.frm.curr_date.value < document.frm.qual_pass_date.value) {
						alert('�հݳ������ �������ں��� �����ϴ�');
						frm.qual_pass_date.focus();
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
				<form action="insa_qual_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="*" >
						<col width="11%" >
						<col width="22%" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">���</th>
                      <td class="left" bgcolor="#FFFFE6"><%=qual_empno%>
					  <input name="qual_empno" type="hidden" id="qual_empno" size="14" value="<%=qual_empno%>" readonly="true">
                      <input name="qual_seq" type="hidden" id="qual_seq" size="14" value="<%=qual_seq%>" readonly="true"></td>
                      <th style="background:#FFFFE6">����</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6"><%=emp_name%></td>
					  <input name="emp_name" type="hidden" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>�ڰ�����</th>
                      <td class="left">
                    <%
					  Sql="select * from emp_etc_code where emp_etc_type = '30' order by emp_etc_name asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
					  <select name="qual_type" id="qual_type" style="width:140px">
                         <option value="" <% if qual_type = "" then %>selected<% end if %>>����</option>
                			  <% 
								do until rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If qual_type = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							  %>
            		  </select>                             
                      </td>
                      <th>���</th>
                      <td colspan="3" class="left">
                      <select name="qual_grade" id="qual_grade" value="<%=qual_grade%>" style="width:90px">
			            	        <option value="" <% if qual_grade = "" then %>selected<% end if %>>����</option>
                                    <option value='1��' <%If qual_grade = "1��" then %>selected<% end if %>>1��</option>
                                    <option value='2��' <%If qual_grade = "2��" then %>selected<% end if %>>2��</option>
				                    <option value='3��' <%If qual_grade = "3��" then %>selected<% end if %>>3��</option>
                                    <option value='�ʱ�' <%If qual_grade = "�ʱ�" then %>selected<% end if %>>�ʱ�</option>
                                    <option value='�߱�' <%If qual_grade = "�߱�" then %>selected<% end if %>>�߱�</option>
                                    <option value='���' <%If qual_grade = "���" then %>selected<% end if %>>���</option>
                                    <option value='Ư��' <%If qual_grade = "Ư��" then %>selected<% end if %>>Ư��</option>
                      </select>
                      </td>
                    </tr>
                    <tr>  
                      <th>�߱ޱ��</th>
                      <td class="left">
                      <input name="qual_org" type="text" id="qual_org" style="width:140px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=qual_org%>"></td>
                      <th>�ڰ�<br>��Ϲ�ȣ</th>
                      <td colspan="3" class="left">
					  <input name="qual_no" type="text" id="qual_no" style="width:150px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=qual_no%>">&nbsp;</td>
                    </tr>
                    <tr>
                      <th>�հݳ����</th>
                      <td colspan="5" class="left">
					  <input name="qual_pass_date" type="text" value="<%=qual_pass_date%>" style="width:80px;text-align:center" id="datepicker">&nbsp;</td>
                    </tr>
                    <tr>  
                      <th>��¼�øNo</th>
                      <td class="left">
                      <input name="qual_passport" type="text" id="qual_passport" style="width:140px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=qual_passport%>"></td>
                      <th>�ڰݼ���</th>
                      <td colspan="3" class="left">
					  <input type="radio" name="qual_pay_id" value="Y" <% if qual_pay_id = "Y" then %>checked<% end if %> style="width:40px" id="Radio1">����
                      <input type="radio" name="qual_pay_id" value="N" <% if qual_pay_id = "N" then %>checked<% end if %> style="width:40px" id="Radio2">����</td>
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
                <input type="hidden" name="curr_date" value="<%=curr_date%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

