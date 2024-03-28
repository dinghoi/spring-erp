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
Dim rs_emp, rsLng, title_line

in_name = user_name
in_empno = user_id

If f_toString(Request.Form("in_empno"), "")  <> "" Then
   'Sql = "SELECT * FROM emp_master where emp_no = '"&in_empno&"'"
   'Set rs_emp = DbConn.Execute(SQL)
   objBuilder.Append "SELECT emp_name FROM emp_master "
   objBuilder.Append "WHERE emp_no='"&in_empno&"';"

   Set rs_emp = DBConn.Execute(objBuilder.ToString())
   objBuilder.Clear()

   in_name = rs_emp("emp_name")
   rs_emp.Close() : Set rs_emp = Nothing
End If

objBuilder.Append "SELECT lang_id, lang_id_type, lang_point, lang_grade, lang_get_date, "
objBuilder.Append "	lang_empno, lang_seq "
objBuilder.Append "FROM emp_language "
objBuilder.Append "WHERE lang_empno='"&in_empno&"' ORDER BY lang_empno, lang_seq ASC"

Set rsLng = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "어학능력 사항"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무관리</title>
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
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<!--<dt>◈조건 검색◈</dt>-->
                        <dd>
                            <p>
							<strong>사번 : </strong>
							<label>
								<input type="text" name="in_empno" id="in_empno" value="<%=in_empno%>" style="width:80px;" class="no-input" readonly/>
							</label>
                            <strong>성명 : </strong>
							<label>
								<input type="text" name="in_name" id="in_name" value="<%=in_name%>" style="width:80px;" class="no-input" readonly/>
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
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th colspan="3">어학구분</th>
                                <th colspan="2">어학종류</th>
                                <th colspan="2">점수</th>
                                <th colspan="2">급수</th>
                                <th colspan="3">취득일</th>
                                <th>수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsLng.EOF Or rsLng.BOF Then
							Response.Write "<tr><td colspan='13' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsLng.EOF
							%>
								<tr>
									<td colspan="3"><%=rsLng("lang_id")%>&nbsp;</td>
									<td colspan="2"><%=rsLng("lang_id_type")%>&nbsp;</td>
									<td colspan="2"><%=rsLng("lang_point")%>&nbsp;</td>
									<td colspan="2"><%=rsLng("lang_grade")%>&nbsp;</td>
									<td colspan="3"><%=rsLng("lang_get_date")%>&nbsp;</td>
									<td>
										<a href="#" onClick="pop_Window('/person/insa_language_add.asp?lang_empno=<%=rsLng("lang_empno")%>&lang_seq=<%=rsLng("lang_seq")%>&emp_name=<%=in_name%>&u_type=U','어학능력사항 변경','scrollbars=yes,width=750,height=300')">수정</a>
									</td>
								</tr>
							<%
								rsLng.MoveNext()
							Loop
						End If
						rsLng.Close() : Set rsLng = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="pop_Window('/person/insa_language_add.asp?lang_empno=<%=in_empno%>&emp_name=<%=in_name%>','어학사항 등록','scrollbars=yes,width=750,height=300')" class="btnType04">어학사항 등록</a>
					</div>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>