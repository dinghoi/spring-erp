<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
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
Dim title_line

title_line = "경력사항 등록"
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
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "" );
			});
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "" );
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
				if(document.frm.c_join_date.value ==""){
					alert('재직기간을 입력하세요');
					frm.c_join_date.focus();
					return false;
				}

				if(document.frm.c_end_date.value ==""){
					alert('재직기간을 입력하세요');
					frm.c_end_date.focus();
					return false;
				}

				if(document.frm.c_office ==""){
					alert('회사명을 선택하세요');
					frm.c_office.focus();
					return false;
				}

				if(document.frm.c_dept.value ==""){
					alert('부서명을 입력하세요');
					frm.c_dept.focus();
					return false;
				}

				if(document.frm.c_position.value ==""){
					alert('직위/직책를 입력하세요');
					frm.c_position.focus();
					return false;
				}

				if(document.frm.c_task.value ==""){
					alert('담당업무를 입력하세요');
					frm.c_task.focus();
					return false;
				}

				var result = "confirm('등록 하시겠습니까?')";
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
			<form action="/member/member_career_proc.asp" method="post" name="frm">
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
					<th style="background:#FFFFE6">성명</th>
					<td colspan="5" class="left" bgcolor="#FFFFE6">
						<input name="m_name" type="text" id="m_name" value="<%=m_name%>" size="14" class="no-input" readonly/>
					</td>
				</tr>
				<tr>
					<th>재직기간<span style="color:red;">*</span></th>
					<td colspan="5" class="left">
						<input type="text" name="c_join_date" id="datepicker" style="width:80px;text-align:center" />
						&nbsp;-&nbsp;
						<input type="text" name="c_end_date" id="datepicker1" style="width:80px;text-align:center" />
					</td>
				</tr>
				<tr>
					<th>회사명<span style="color:red;">*</span></th>
					<td class="left">
						<input type="text" name="c_office" id="c_office" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);"/>
					</td>
					<th>부서<span style="color:red;">*</span></th>
					<td colspan="3" class="left">
						<input type="text" name="c_dept" id="c_dept" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);"/>&nbsp;
					</td>
				</tr>
				<tr>
					<th>직위/직책<span style="color:red;">*</span></th>
					<td class="left">
						<input type="text" name="c_position" id="c_position" style="width:130px; ime-mode:active" onKeyUp="checklength(this,20);"/>&nbsp;
					</td>
					<th>담당업무<span style="color:red;">*</span></th>
					<td colspan="3" class="left">
						<input type="text" name="c_task" id="c_task" style="width:250px; ime-mode:active" onKeyUp="checklength(this,50);"/>&nbsp;
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