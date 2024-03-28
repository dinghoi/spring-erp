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
Dim u_type, lang_empno, lang_seq, emp_name, lang_id, lang_id_type
Dim lang_point, lang_grade, pang_get_date, curr_date, title_line, lang_get_date
Dim rsLng, rs_etc, rsEtc

u_type = Request.QueryString("u_type")
lang_empno = Request.QueryString("lang_empno")
lang_seq = Request.QueryString("lang_seq")
emp_name = Request.QueryString("emp_name")

lang_id = ""
lang_id_type = ""
lang_point = ""
lang_grade = ""
lang_get_date = ""

curr_date = Mid(CStr(Now()), 1, 10)

title_line = "어학능력 등록"

If u_type = "U" Then
	objBuilder.Append "SELECT lang_id, lang_id_type, lang_point, lang_grade, lang_get_date "
	objBuilder.Append "FROM emp_language "
	objBuilder.Append "WHERE lang_empno = '"&lang_empno&"' AND lang_seq = '"&lang_seq&"';"

	Set rsLng = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	lang_id = rsLng("lang_id")
    lang_id_type = rsLng("lang_id_type")
    lang_point = rsLng("lang_point")
    lang_grade = rsLng("lang_grade")
    lang_get_date = rsLng("lang_get_date")

	rsLng.Close() : Set rsLng = Nothing

	title_line = "어학능력 변경"
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
			//취득일
			$(function() {
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=lang_get_date%>" );
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
				if(document.frm.lang_id.value == ""){
					alert('어학구분을 선택해주세요.');
					frm.lang_id.focus();
					return false;
				}

				if(document.frm.lang_id_type == ""){
					alert('어학종류을 선택해주세요.');
					frm.lang_id_type.focus();
					return false;
				}

				if(document.frm.lang_grade.value == ""){
					alert('급수를 입력해주세요.');
					frm.lang_grade.focus();
					return false;
				}

				if(document.frm.lang_point.value == ""){
					alert('점수를 입력해주세요.');
					frm.lang_point.focus();
					return false;
				}

				if(document.frm.lang_get_date.value == ""){
					alert('취득일을 입력해주세요.');
					frm.lang_get_date.focus();
					return false;
				}

				if(document.frm.lang_get_date.value > document.frm.curr_date.value){
					alert('취득일이 현재일보다 빠릅니다.');
					frm.lang_get_date.focus();
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
			<form action="/person/insa_language_add_save.asp" method="post" name="frm">
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
						<input type="text" name="lang_empno" id="lang_empno" size="14" value="<%=lang_empno%>" class="no-input" readonly/>
						<input type="hidden" name="lang_seq" value="<%=lang_seq%>"/>
					</td>
					<th style="background:#FFFFE6">성명</th>
					<td colspan="3" class="left" bgcolor="#FFFFE6">
						<input type="text" name="emp_name" id="emp_name" size="14" value="<%=emp_name%>" class="no-input" readonly/>
					</td>
				</tr>
				<tr>
					<th>어학구분<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
					<%
					objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
					objBuilder.Append "WHERE emp_etc_type = '08' ORDER BY emp_etc_code ASC;"

					Set rs_etc = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()
					%>
						<select name="lang_id" id="lang_id" style="width:90px">
							<option value="" <%If lang_id = "" Then %>selected<%End If %>>선택</option>
					<%
					Do until rs_etc.EOF
					%>
							<option value='<%=rs_etc("emp_etc_name")%>' <%If lang_id = rs_etc("emp_etc_name") Then %>selected<%End If %>><%=rs_etc("emp_etc_name")%></option>
					<%
						rs_etc.MoveNext()
					Loop
					rs_etc.Close() : Set rs_etc = Nothing
					%>
				  </select>
				  </td>
				</tr>
				<tr>
					<th>어학종류<span style="color:red;">*</span></th>
					<td class="left">
					<%
					objBuilder.Append "SELECT emp_etc_name  FROM emp_etc_code "
					objBuilder.Append "WHERE emp_etc_type = '09' ORDER BY emp_etc_code ASC;"

					Set rsEtc = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()
					%>
						<select name="lang_id_type" id="lang_id_type" style="width:90px">
							<option value="" <%If lang_id_type = "" Then %>selected<%End If %>>선택</option>
					<%
					Do Until rsEtc.EOF
					%>
								<option value='<%=rsEtc("emp_etc_name")%>' <%If lang_id_type = rsEtc("emp_etc_name") Then %>selected<%End If %>><%=rsEtc("emp_etc_name")%></option>
					<%
						rsEtc.MoveNext()
					Loop
					rsEtc.Close() : Set rsEtc = Nothing
					DBConn.Close : Set DBConn = Nothing
					%>
						</select>
					</td>
					<th>급수<span style="color:red;">*</span></th>
					<td class="left">
						<select name="lang_grade" id="lang_grade" value="<%=lang_grade%>" style="width:100px">
							<option value="" <%If lang_grade = "" Then %>selected<%End If %>>선택</option>
							<option value='급수었음' <%If lang_grade = "급수었음" Then %>selected<%End If %>>급수없음</option>
							<option value='3급' <%If lang_grade = "3급" Then %>selected<%End If %>>3급</option>
							<option value='2급' <%If lang_grade = "2급" Then %>selected<%End If %>>2급</option>
							<option value='1급' <%If lang_grade = "1급" Then %>selected<%End If %>>1급</option>
						</select>
					</td>
				  <th>점수<span style="color:red;">*</span></th>
				  <td class="left">
					<input type="text" name="lang_point" id="lang_point" style="width:80px; ime-mode:active" onKeyUp="checklength(this,4);" value="<%=lang_point%>"/>
				  </td>
				</tr>
				<tr>
					<th>취득일<span style="color:red;">*</span></th>
					<td colspan="5" class="left">
						<input type="text" name="lang_get_date" value="<%=lang_get_date%>" style="width:80px;text-align:center" id="datepicker"/>&nbsp;
					</td>
				</tr>
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
			<input type="hidden" name="curr_date" value="<%=curr_date%>"/>
			</form>
		</div>
	</body>
</html>