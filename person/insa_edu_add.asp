<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim u_type,edu_empno, edu_seq, emp_name, edu_name, edu_office
Dim edu_finish_no, edu_start_date, edu_pay, edu_comment, edu_reg_date
Dim curr_date, title_line, rsEdu, edu_end_date

u_type = Request.QueryString("u_type")
edu_empno = Request.QueryString("edu_empno")
edu_seq = Request.QueryString("edu_seq")
emp_name = Request.QueryString("emp_name")

edu_name = ""
edu_office = ""
edu_finish_no = ""
edu_start_date = ""
edu_end_date = ""
edu_pay = 0
edu_comment = ""
edu_reg_date = ""

curr_date = Mid(CStr(Now()), 1, 10)
title_line = "교육사항 등록"

If u_type = "U" Then
	objBuilder.Append "SELECT edu_empno, edu_seq, edu_name, edu_office, "
	objBuilder.Append "	edu_finish_no, edu_start_date, edu_end_date, edu_pay, "
	objBuilder.Append "	edu_comment, edu_reg_date "
	objBuilder.Append "FROM emp_edu "
	objBuilder.Append "WHERE edu_empno = '"&edu_empno&"' AND edu_seq = '"&edu_seq&"';"

	Set rsEdu = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	edu_empno = rsEdu("edu_empno")
    edu_seq = rsEdu("edu_seq")
	edu_name = rsEdu("edu_name")
    edu_office = rsEdu("edu_office")
    edu_finish_no = rsEdu("edu_finish_no")
    edu_start_date = rsEdu("edu_start_date")
    edu_end_date = rsEdu("edu_end_date")
    edu_pay = rsEdu("edu_pay")
    edu_comment = rsEdu("edu_comment")
    edu_reg_date = rsEdu("edu_reg_date")

	rsEdu.Close() : Set rsEdu = Nothing
	DBConn.Close() : Set DBConn = Nothing

	title_line = "교육사항 변경"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무관리</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			//교육 시작일
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=edu_start_date%>" );
			});

			//교육 종료일
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=edu_end_date%>" );
			});

			function goAction(){
			   window.close();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.edu_name.value ==""){
					alert('교육과정을 입력해주세요.');
					frm.edu_name.focus();
					return false;
				}

				if(document.frm.edu_finish_no ==""){
					alert('교육수료증No.을 입력해주세요.');
					frm.edu_finish_no.focus();
					return false;
				}

				if(document.frm.edu_office.value ==""){
					alert('교육기관을 입력해주세요.');
					frm.edu_office.focus();
					return false;
				}

				if(document.frm.edu_start_date.value ==""){
					alert('교육기간을 입력해주세요.');
					frm.edu_start_date.focus();
					return false;
				}

				if(document.frm.edu_end_date.value ==""){
					alert('교육기간을 입력해주세요.');
					frm.edu_end_date.focus();
					return false;
				}

				if(document.frm.edu_end_date.value < document.frm.edu_start_date.value){
					alert('교육시작일이 교육마지막일자보다 빠릅니다');
					frm.edu_end_date.focus();
					return false;
				}

				var result = confirm('등록 하시겠습니까?');

				if(result){
					return true;
				}
				return false;
			}
        </script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body>
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/person/insa_edu_add_save.asp" method="post" name="frm">
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
						<input type="text" name="edu_empno" id="edu_empno" size="14" value="<%=edu_empno%>" class="no-input" readonly/>
						<input type="hidden" name="edu_seq" value="<%=edu_seq%>"/></td>
					<th style="background:#FFFFE6">성명</th>
					<td colspan="3" class="left" bgcolor="#FFFFE6">
						<input type="text" name="emp_name" id="emp_name" size="14" value="<%=emp_name%>" class="no-input" readonly/>
					</td>
				</tr>
				<tr>
					<th>교육과정명<span style="color:red;">*</span></th>
					<td class="left">
						<input type="text" name="edu_name" id="edu_name" style="width:140px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=edu_name%>"/>&nbsp;
					</td>
					<th colspan="2">교육수료증N0.<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="edu_finish_no" id="edu_finish_no" style="width:130px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=edu_finish_no%>"/>
					</td>
				</tr>
				<tr>
					<th>교육기관<span style="color:red;">*</span></th>
					<td class="left">
						<input type="text" name="edu_office" id="edu_office" style="width:140px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=edu_office%>"/>
					</td>
				</tr>
				<tr>
					<th>교육기간<span style="color:red;">*</span></th>
					<td colspan="5" class="left">
						<input type="text" name="edu_start_date" value="<%=edu_start_date%>" style="width:80px;text-align:center" id="datepicker"/>&nbsp;
						&nbsp;-&nbsp;
						<input type="text" name="edu_end_date" value="<%=edu_end_date%>" style="width:80px;text-align:center" id="datepicker1"/>&nbsp;
					</td>
				</tr>
				<tr>
					<th>교육<br>주요내용</th>
					<td class="left" colspan="5"><textarea name="edu_comment"><%=edu_comment%></textarea></td>
				</tr>
				</tbody>
			  </table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="<%If u_type = "U" Then%>수정<%Else%>등록<%End If%>" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"/></span>
			</div>
			<input type="hidden" name="u_type" value="<%=u_type%>"/>
			</form>
		</div>
	</body>
</html>