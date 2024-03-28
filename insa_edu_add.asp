<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
edu_empno = request("edu_empno")
edu_seq = request("edu_seq")
emp_name = request("emp_name")

edu_name = ""
edu_office = ""
edu_finish_no = ""
edu_start_date = ""
edu_end_date = ""
edu_pay = 0
edu_comment = ""
edu_reg_date = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 교육사항 등록 "
if u_type = "U" then

	Sql="select * from emp_edu where edu_empno = '"&edu_empno&"' and edu_seq = '"&edu_seq&"'"
	Set rs=DbConn.Execute(Sql)

	edu_empno = rs("edu_empno")
    edu_seq = rs("edu_seq")
	edu_name = rs("edu_name")
    edu_office = rs("edu_office")
    edu_finish_no = rs("edu_finish_no")
    edu_start_date = rs("edu_start_date")
    edu_end_date = rs("edu_end_date")
    edu_pay = rs("edu_pay")
    edu_comment = rs("edu_comment")
    edu_reg_date = rs("edu_reg_date")
	
	rs.close()

	title_line = " 교육사항 변경 "
	
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
												$( "#datepicker" ).datepicker("setDate", "<%=edu_start_date%>" );
			});	 
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=edu_end_date%>" );
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
				if(document.frm.edu_name.value =="") {
					alert('교육과정을 입력하세요');
					frm.edu_name.focus();
					return false;}
				if(document.frm.edu_finish_no =="") {
					alert('교육수료증No.을 입력하세요');
					frm.edu_finish_no.focus();
					return false;}
				if(document.frm.edu_office.value =="") {
					alert('교육기관을 입력하세요');
					frm.edu_office.focus();
					return false;}
				if(document.frm.edu_start_date.value =="") {
					alert('교육기간을 입력하세요');
					frm.edu_start_date.focus();
					return false;}
				if(document.frm.edu_end_date.value =="") {
					alert('교육기간을 입력하세요');
					frm.edu_end_date.focus();
					return false;}
				if(document.frm.edu_end_date.value < document.frm.edu_start_date.value) {
						alert('교육시작일이 교육마지막일자보다 빠릅니다');
						frm.edu_end_date.focus();
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
				<form action="insa_edu_add_save.asp" method="post" name="frm">
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
					  <input name="edu_empno" type="text" id="edu_empno" size="14" value="<%=edu_empno%>" readonly="true">
                      <input type="hidden" name="edu_seq" value="<%=edu_seq%>" ID="Hidden1"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>  
                      <th>교육과정명</th>
                      <td class="left">
                      <input name="edu_name" type="text" id="edu_name" style="width:140px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=edu_name%>">&nbsp;</td>
                      <th colspan="2">교육수료증N0.</th>
                      <td colspan="2" class="left">
                      <input name="edu_finish_no" type="text" id="edu_finish_no" style="width:130px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=edu_finish_no%>"></td>
                    </tr>
                    <tr>  
                      <th>교육기관</th>
                      <td class="left">
                      <input name="edu_office" type="text" id="edu_office" style="width:140px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=edu_office%>"></td>
                    </tr>
                    <tr>
                      <th>교육기간</th>
                      <td colspan="5" class="left">
					  <input name="edu_start_date" type="text" value="<%=edu_start_date%>" style="width:80px;text-align:center" id="datepicker">&nbsp;
                      &nbsp;-&nbsp;
                      <input name="edu_end_date" type="text" value="<%=edu_end_date%>" style="width:80px;text-align:center" id="datepicker1">&nbsp;</td>
                    </tr>
                    <tr>
                      <th>교육<br>주요내용</th>
                      <td class="left" colspan="5"><textarea name="edu_comment"><%=edu_comment%></textarea></td>
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

