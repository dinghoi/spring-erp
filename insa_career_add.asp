<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
career_empno = request("career_empno")
career_seq = request("career_seq")
emp_name = request("emp_name")

career_join_date = ""
career_end_date = ""
career_office = ""
career_dept = ""
career_position = ""
career_task = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 경력사항 등록 "
if u_type = "U" then

	Sql="select * from emp_career where career_empno = '"&career_empno&"' and career_seq = '"&career_seq&"'"
	Set rs=DbConn.Execute(Sql)

    career_empno = rs("career_empno")
    career_seq = rs("career_seq")
	
	career_join_date = rs("career_join_date")
    career_end_date = rs("career_end_date")
    career_office = rs("career_office")
    career_dept = rs("career_dept")
    career_position = rs("career_position")
    career_task = rs("career_task")
	
	rs.close()

	title_line = " 경력사항 변경 "
	
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
												$( "#datepicker" ).datepicker("setDate", "<%=career_join_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=career_end_date%>" );
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
				if(document.frm.career_join_date.value =="") {
					alert('재직기간을 입력하세요');
					frm.career_join_date.focus();
					return false;}
				if(document.frm.career_end_date.value =="") {
					alert('재직기간을 입력하세요');
					frm.career_end_date.focus();
					return false;}
				if(document.frm.career_office =="") {
					alert('회사명을 선택하세요');
					frm.career_office.focus();
					return false;}
				if(document.frm.career_dept.value =="") {
					alert('부서명을 입력하세요');
					frm.career_dept.focus();
					return false;}
				if(document.frm.career_position.value =="") {
					alert('직위를 입력하세요');
					frm.career_position.focus();
					return false;}
				if(document.frm.career_task.value =="") {
					alert('담당업무를 입력하세요');
					frm.career_task.focus();
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
				<form action="insa_career_add_save.asp" method="post" name="frm">
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
					  <input name="career_empno" type="text" id="career_empno" size="14" value="<%=career_empno%>" readonly="true">
                      <input type="hidden" name="career_seq" value="<%=career_seq%>" ID="Hidden1"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>재직기간</th>
                      <td colspan="5" class="left">
                      <input name="career_join_date" type="text" value="<%=career_join_date%>" style="width:80px;text-align:center" id="datepicker">
                      &nbsp;-&nbsp;
                      <input name="career_end_date" type="text" value="<%=career_end_date%>" style="width:80px;text-align:center" id="datepicker1">
                      </td>
                    </tr>
                    <tr>  
                      <th>회사명</th>
                      <td class="left">
                      <input name="career_office" type="text" id="career_office" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=career_office%>"></td>
                      <th>부서</th>
                      <td colspan="3" class="left">
					  <input name="career_dept" type="text" id="career_dept" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=career_dept%>">&nbsp;</td>
                    </tr>
                    <tr>
                      <th>직위/직책</th>
                      <td class="left">
					  <input name="career_position" type="text" id="career_position" style="width:130px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=career_position%>">&nbsp;</td>
                      <th>담당업무</th>
                      <td colspan="3" class="left">
					  <input name="career_task" type="text" id="career_task" style="width:250px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=career_task%>">&nbsp;</td>
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
				</form>
		</div>				
	</body>
</html>

