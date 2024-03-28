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
Dim u_type,career_empno, career_seq, emp_name, title_line
Dim rsCareer
Dim career_join_date, career_end_date
Dim career_office, career_dept, career_position, career_task

u_type = Request.QueryString("u_type")
career_empno = Request.QueryString("career_empno")
career_seq = Request.QueryString("career_seq")
emp_name = Request.QueryString("emp_name")

career_join_date = ""
career_end_date = ""
career_office = ""
career_dept = ""
career_position = ""
career_task = ""

title_line = " 경력사항 등록 "

If u_type = "U" Then
	objBuilder.Append "SELECT career_empno, career_seq, career_join_date, career_end_date,  "
	objBuilder.Append "	career_office, career_dept, career_position, career_task "
	objBuilder.Append "FROM emp_career "
	objBuilder.Append "WHERE career_empno = '"&career_empno&"' and career_seq = '"&career_seq&"';"

	Set rsCareer = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

    career_empno = rsCareer("career_empno")
    career_seq = rsCareer("career_seq")

	career_join_date = rsCareer("career_join_date")
    career_end_date = rsCareer("career_end_date")
    career_office = rsCareer("career_office")
    career_dept = rsCareer("career_dept")
    career_position = rsCareer("career_position")
    career_task = rsCareer("career_task")

	rsCareer.Close() : Set rsCareer = Nothing

	title_line = " 경력사항 변경 "
End If
DBConn.Close() : Set DBConn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
			//재직(입사)일자
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=career_join_date%>" );
			});

			//재직(퇴사)일자
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=career_end_date%>" );
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
				if(document.frm.career_join_date.value == ""){
					alert('재직기간을 입력하세요');
					frm.career_join_date.focus();
					return false;
				}

				if(document.frm.career_end_date.value == ""){
					alert('재직기간을 입력하세요');
					frm.career_end_date.focus();
					return false;
				}

				if(document.frm.career_office == ""){
					alert('회사명을 선택하세요');
					frm.career_office.focus();
					return false;
				}

				if(document.frm.career_dept.value == ""){
					alert('부서명을 입력하세요');
					frm.career_dept.focus();
					return false;
				}

				if(document.frm.career_position.value == ""){
					alert('직위/직책를 입력하세요');
					frm.career_position.focus();
					return false;
				}

				if(document.frm.career_task.value == ""){
					alert('담당업무를 입력하세요');
					frm.career_task.focus();
					return false;
				}

				var result = confirm('등록하시겠습니까?');

				if(result == true){
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
				<form action="/insa/insa_career_add_save.asp" method="post" name="frm">
					<input type="hidden" name="career_seq" value="<%=career_seq%>"/>
					<input type="hidden" name="u_type" value="<%=u_type%>" />
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
							<input type="text" name="career_empno" id="career_empno" size="14" value="<%=career_empno%>" class="no-input" readonly/>
						</td>
						<th style="background:#FFFFE6">성명</th>
						<td colspan="3" class="left" bgcolor="#FFFFE6">
							<input type="text" name="emp_name" id="emp_name" size="14" value="<%=emp_name%>" class="no-input" readonly/>
						</td>
                    </tr>
                 	<tr>
						<th>재직기간<span style="color:red;">*</span></th>
						<td colspan="5" class="left">
							<input type="text" name="career_join_date" value="<%=career_join_date%>" style="width:80px;text-align:center;" id="datepicker"/>
							&nbsp;-&nbsp;
							<input type="text" name="career_end_date" value="<%=career_end_date%>" style="width:80px;text-align:center;" id="datepicker1"/>
						</td>
                    </tr>
                    <tr>
						<th>회사명<span style="color:red;">*</span></th>
						<td class="left">
							<input type="text" name="career_office" id="career_office" style="width:130px; ime-mode:active;" onKeyUp="checklength(this,30);" value="<%=career_office%>"/>
						</td>
						<th>부서<span style="color:red;">*</span></th>
						<td colspan="3" class="left">
							<input type="text" name="career_dept" id="career_dept" style="width:130px; ime-mode:active;" onKeyUp="checklength(this,30);" value="<%=career_dept%>"/>&nbsp;
						</td>
                    </tr>
                    <tr>
						<th>직위/직책<span style="color:red;">*</span></th>
						<td class="left">
							<input type="text" name="career_position" id="career_position" style="width:130px; ime-mode:active;" onKeyUp="checklength(this,20);" value="<%=career_position%>"/>&nbsp;
						</td>
						<th>담당업무<span style="color:red;">*</span></th>
						<td colspan="3" class="left">
							<input type="text" name="career_task" id="career_task" style="width:250px; ime-mode:active;" onKeyUp="checklength(this,50);" value="<%=career_task%>"/>&nbsp;
						</td>
					</tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();"/></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"/></span>
                </div>
			</form>
		</div>
	</body>
</html>

