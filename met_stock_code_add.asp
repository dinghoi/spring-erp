<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
org_code = request("org_code")

stock_level=Request("stock_level")
stock_code=Request("stock_code")
stock_name=Request("stock_name")

code_last = ""

stock_name = ""
stock_company = ""
stock_bonbu = ""
stock_saupbu = ""
stock_team = ""
stock_open_date = ""
stock_end_date = ""
stock_manager_code = ""
stock_manager_name = ""

' response.write(reg_date)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = stock_level + " â�� ��� "

if u_type = "U" then

	Sql="select * from met_stock_code where stock_code = '"&stock_code&"'"
	Set rs=DbConn.Execute(Sql)

	stock_code = rs("stock_code")
	stock_level = rs("stock_level")
	stock_name = rs("stock_name")
    stock_company = rs("stock_company")
    stock_bonbu = rs("stock_bonbu")
    stock_saupbu = rs("stock_saupbu")
    stock_team = rs("stock_team")
    stock_open_date = rs("stock_open_date")
    stock_end_date = rs("stock_end_date")
	stock_manager_code = rs("stock_manager_code")
    stock_manager_name = rs("stock_manager_name")
	
	stock_go_man = rs("stock_go_man")
    stock_go_name = rs("stock_go_name")
	stock_in_man = rs("stock_in_man")
    stock_in_name = rs("stock_in_name")
	if stock_end_date = "1900-01-01" then
	      stock_end_date = ""
	end if
	
	rs.close()
    
	title_line = stock_level + " â�� ���� "
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��ǰ������� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "6 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=stock_open_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=stock_end_date%>" );
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
				if(document.frm.stock_level.value =="") {
					alert('â�������� �����ϼ���');
					frm.stock_level.focus();
					return false;}
				if(document.frm.stock_code.value =="") {
					alert('â����� �����ϼ���');
					frm.stock_code.focus();
					return false;}
				if(document.frm.stock_go_man.value =="") {
					alert('������ڸ� �����ϼ���');
					frm.stock_go_man.focus();
					return false;}
				if(document.frm.stock_in_man.value =="") {
					alert('�԰����ڸ� �����ϼ���');
					frm.stock_in_man.focus();
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
				<form action="met_stock_code_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
						       <col width="11%" >
						       <col width="30%" >
						       <col width="15%" >
						       <col width="15%" >
						       <col width="15%" >
						       <col width="*" >
						</colgroup>
						<tbody>
                            <tr>
                                <th class="first" style="background:#FFFFE6">â������</th>
                                <td colspan="5" class="left" bgcolor="#FFFFE6">
                                <input type="radio" name="stock_level" value="����" <% if stock_level = "����" then %>checked<% end if %>>���� 
              		            <input type="radio" name="stock_level" value="����" <% if stock_level = "����" then %>checked<% end if %>>���� 
                                <input type="radio" name="stock_level" value="�����" <% if stock_level = "�����" then %>checked<% end if %>>����� 
                                <input type="radio" name="stock_level" value="��" <% if stock_level = "��" then %>checked<% end if %>>�� 
                                </td>
                            </tr>
							<tr>
                                <th>â���</th>
                                <td colspan="2" class="left">
                                <input name="stock_code" type="text" id="stock_code" style="width:40px" readonly="true" value="<%=stock_code%>">
                                &nbsp;��&nbsp;
                                <input name="stock_name" type="text" id="stock_name" style="width:100px" readonly="true" value="<%=stock_name%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="stock"%>&stock_level=<%=stock_level%>','orgselect','scrollbars=yes,width=800,height=400')">����</a>
                                </td> 
                                <th class="first">â����</th>
                                <td colspan="2" class="left">
                                <input name="stock_manager_code" type="text" id="stock_manager_code" readonly="true" style="width:40px" value="<%=stock_manager_code%>">
                                &nbsp;��&nbsp;
                                <input name="stock_manager_name" type="text" id="stock_manager_name" readonly="true" style="width:60px" value="<%=stock_manager_name%>"></td>   
                            </tr>
							<tr>
                                <th>�Ҽ�����</th>
                                <td colspan="5" class="left">
                                <input name="stock_company" type="text" id="stock_company" style="width:100px" readonly="true" value="<%=stock_company%>">
              					<input name="stock_bonbu" type="text" id="stock_bonbu" style="width:100px" readonly="true" value="<%=stock_bonbu%>">
              					<input name="stock_saupbu" type="text" id="stock_saupbu" style="width:100px" readonly="true" value="<%=stock_saupbu%>">
              					<input name="stock_team" type="text" id="stock_team" style="width:100px" readonly="true" value="<%=stock_team%>">
                                </td>    
                            </tr>
                             <tr>
                                <th>������</th>
                                <td colspan="2" class="left">
                                <input name="stock_open_date" type="text" size="10" readonly="true" id="datepicker" style="width:70px;" value="<%=stock_open_date%>" >
              					</td>
                                <th>������</th>
                                <td colspan="2" class="left">
                                <input name="stock_end_date" type="text" size="10" readonly="true" id="datepicker1" style="width:70px;" value="<%=stock_end_date%>" >
              					</td>
                             </tr>
                             <tr>
                                <th>�����</th>
                                <td colspan="2" class="left">
                                <input name="stock_go_man" type="text" id="stock_go_man" style="width:40px" readonly="true" value="<%=stock_go_man%>">
                                <input name="stock_go_name" type="text" id="stock_go_name" style="width:60px" readonly="true" value="<%=stock_go_name%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_emp_select.asp?gubun=<%="st_emp1"%>&view_condi=<%=view_condi%>','orgempselect','scrollbars=yes,width=600,height=400')">ã��</a>
                                </td>
                                <th>�԰���</th>
                                <td colspan="2" class="left">
                                <input name="stock_in_man" type="text" id="stock_in_man" style="width:40px" readonly="true" value="<%=stock_in_man%>">
                                <input name="stock_in_name" type="text" id="stock_in_name" style="width:60px" readonly="true" value="<%=stock_in_name%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_emp_select.asp?gubun=<%="st_emp2"%>&view_condi=<%=view_condi%>','orgempselect','scrollbars=yes,width=600,height=400')">ã��</a>
                                </td>
                             </tr>
                             <tr>
                                <th class="first">�Է�����</th>
                                <td colspan="2" class="left"><%=reg_date%>(<%=reg_user%>)</td>
                                <th>��������</th>
                                <td colspan="2" class="left"><%=mod_date%>(<%=mod_user%>)</td>
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
				</form>
		</div>				
	</body>
</html>

