<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

win_sw = "close"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If Request.Form("in_empno")  <> "" Then 
   Sql = "SELECT * FROM emp_master where emp_no = '"&in_empno&"'"
   Set rs_emp = DbConn.Execute(SQL)
   in_name = rs_emp("emp_name")
   rs_emp.close()
End If

sql = "select * from emp_language where lang_empno = '" + in_empno + "' ORDER BY lang_empno,lang_seq ASC"
Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " ���дɷ� ���� "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
			function goAction () {
			   window.close () ;
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.in_empno.value == "") {
					alert ("����� �Է��Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_individual_language.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>������ �˻���</dt>
                        <dd>
                            <p>
							<strong>��� : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=in_empno%>" readonly="true" style="width:100px; text-align:left">
								</label>
                            <strong>���� : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=in_name%>" readonly="true" style="width:150px; text-align:left">
								</label>
                                
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="5%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th colspan="3">���б���</th>
                                <th colspan="2">��������</th>
                                <th colspan="2">����</th>
                                <th colspan="2">�޼�</th>
                                <th colspan="3">�����</th>
                                <th>����</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						do until rs.eof
						%>
							<tr>
                              <td colspan="3" ><%=rs("lang_id")%>&nbsp;</td>
                              <td colspan="2" ><%=rs("lang_id_type")%>&nbsp;</td>
                              <td colspan="2" ><%=rs("lang_point")%>&nbsp;</td>
                              <td colspan="2" ><%=rs("lang_grade")%>&nbsp;</td>
                              <td colspan="3" ><%=rs("lang_get_date")%>&nbsp;</td>
							  <td><a href="#" onClick="pop_Window('insa_language_add.asp?lang_empno=<%=rs("lang_empno")%>&lang_seq=<%=rs("lang_seq")%>&emp_name=<%=in_name%>&u_type=<%="U"%>','insa_language_add_pop','scrollbars=yes,width=750,height=300')">����</a></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_language_add.asp?lang_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_language_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">���л��� ���</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="lang_empno" value="<%=in_empno%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

