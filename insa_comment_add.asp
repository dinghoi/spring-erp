<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
cmt_empno = request("cmt_empno")
cmt_date = request("cmt_date")
emp_name = request("emp_name")

cmt_emp_name = ""
cmt_company = ""
cmt_bonbu = ""
cmt_saupbu = ""
cmt_team = ""
cmt_org_name = ""
cmt_org_code = ""
cmt_comment = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

If cmt_empno  <> "" Then 
   Sql = "SELECT * FROM emp_master where emp_no = '"&cmt_empno&"'"
   Set rs_emp = DbConn.Execute(SQL)
   emp_company = rs_emp("emp_company")
   emp_bonbu = rs_emp("emp_bonbu")
   emp_saupbu = rs_emp("emp_saupbu")
   emp_team = rs_emp("emp_team")
   emp_org_code = rs_emp("emp_org_code")
   emp_org_name = rs_emp("emp_org_name")
   rs_emp.close()
End If


title_line = " 인사특이사항 등록 "
if u_type = "U" then

	Sql="select * from emp_comment where cmt_empno = '"&cmt_empno&"' and cmt_date = '"&cmt_date&"'"
	Set rs=DbConn.Execute(Sql)

	If rs.BOF or rs.EOF Then
		cmt_comment = ""
    	cmt_emp_name = ""
	Else
		cmt_comment = rs("cmt_comment")
    	cmt_emp_name = rs("cmt_emp_name")
	End If

	rs.close()

	title_line = " 인사특이사항 변경 "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=cmt_date%>" );
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
				if(document.frm.cmt_date.value =="") {
					alert('발생일을 입력하세요');
					frm.cmt_date.focus();
					return false;}
				if(document.frm.cmt_comment =="") {
					alert('특이사항을 입력하세요');
					frm.cmt_comment.focus();
					return false;}
				
				{
				a=confirm('입력하시겠습니까?')
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
				<form action="insa_comment_add_save.asp" method="post" name="frm">
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
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="cmt_empno" type="text" id="cmt_empno" size="9" value="<%=cmt_empno%>" readonly="true"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>발생일</th>
                      <td colspan="5" class="left">
                   <% if u_type = "U" then %>
					  <input name="cmt_date" type="text" value="<%=cmt_date%>" style="width:80px;text-align:center" readonly="true">
                   <%     else  %>   
                      <input name="cmt_date" type="text" value="<%=cmt_date%>" style="width:80px;text-align:center" id="datepicker">
                   <% end if %>
					  </td>
                    </tr>
                    <tr>
					  <th class="first">특이사항</th>
					  <td colspan="5" class="left">
                      <textarea name="cmt_comment" rows="2" id="textarea"><%=cmt_comment%></textarea></td>
                   </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="emp_company" value="<%=emp_company%>" ID="Hidden1">
                <input type="hidden" name="emp_bonbu" value="<%=emp_bonbu%>" ID="Hidden1">
                <input type="hidden" name="emp_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="emp_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="emp_org_code" value="<%=emp_org_code%>" ID="Hidden1">
                <input type="hidden" name="emp_org_name" value="<%=emp_org_name%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

