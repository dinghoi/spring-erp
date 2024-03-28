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
Dim rsCrr, title_line

If m_seq = "" Or m_name = "" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('회원기본가입 등록 후 이용 가능합니다.');"
	Response.Write "	location.href='/member/member_add.asp';"
	Response.Write "</script>"

	Response.End
End If


'If f_toString(Request.Form("in_empno"), "") <> "" Then
'   objBuilder.Append "SELECT emp_name FROM emp_master WHERE emp_no = '"&in_empno&"';"

'   in_name = rs_emp("emp_name")
'   rs_emp.Close() : Set rs_emp = Nothing
'End If

objBuilder.Append "SELECT c_join_date, c_end_date, c_office, c_dept, c_position, "
objBuilder.Append "	c_task, c_seq "
objBuilder.Append "FROM member_career "
objBuilder.Append "WHERE m_seq = '"&m_seq&"' "
objBuilder.Append "ORDER BY m_seq, c_seq ASC "

Set rsCrr = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "경력 사항"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>회원관리</title>
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

			//경력 등록 팝업
			function careerAddPopup(){
				var url = '/member/member_career_add.asp';
				var pop_name = '학력사항 등록';
				var features = 'scrollbars=yes,width=750,height=300';

				console.log(url);

				pop_Window(url, pop_name, features);
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
                            <strong>성명 : </strong>
                            <label>
								<input type="text" name="m_name" id="m_name" value="<%=m_name%>" readonly="true" style="width:150px; text-align:left"/>
							</label>
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
						</colgroup>
						<thead>
                            <tr>
                            <th colspan="3">재직기간</th>
                            <th colspan="2">회사명</th>
                            <th colspan="2">부서</th>
                            <th colspan="1">직위</th>
                            <th colspan="3">담당업무</th>
                            <th>순번</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsCrr.EOF Or rsCrr.BOF Then
							Response.Write "<tr><td colspan='12' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsCrr.EOF
						%>
							<tr>
                              <td colspan="3"><%=rsCrr("c_join_date")%>∼<%=rsCrr("c_end_date")%>&nbsp;</td>
                              <td colspan="2"><%=rsCrr("c_office")%>&nbsp;</td>
                              <td colspan="2"><%=rsCrr("c_dept")%>&nbsp;</td>
                              <td colspan="1"><%=rsCrr("c_position")%>&nbsp;</td>
                              <td colspan="3"><%=rsCrr("c_task")%>&nbsp;</td>
                              <td class="right"><%=rsCrr("c_seq")%>&nbsp;</td>
							</tr>
						<%
								rsCrr.MoveNext()
							Loop
						End If
						rsCrr.Close() : Set rsCrr = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="careerAddPopup();" class="btnType04">경력등록</a>
					</div>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>

