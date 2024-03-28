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
Dim strSql, rs_emp, rsEdu, title_line

in_name = user_name
in_empno = user_id

If f_toString(Request.Form("in_empno"), "") <> "" Then
   strSql = "SELECT emp_name FROM emp_master WHERE emp_no = '"&in_empno&"';"

   Set rs_emp = DBConn.Execute(strSql)

   in_name = rs_emp("emp_name")

   rs_emp.Close() : Set rs_emp = Nothing
End If

objBuilder.Append "SELECT edu_name, edu_office, edu_finish_no, edu_start_date, edu_end_date, edu_comment, "
objBuilder.Append "	edu_empno, edu_seq "
objBuilder.Append "FROM emp_edu "
objBuilder.Append "WHERE edu_empno = '"&in_empno&"' "
objBuilder.Append "ORDER BY edu_empno, edu_seq ASC;"

Set rsEdu = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "교육 사항"
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
                        <dd>
                            <p>
							<strong>사번 : </strong>
							<label>
							<input name="in_empno" type="text" id="in_empno" value="<%=in_empno%>" style="width:80px;" class="no-input" readonly/>
							</label>
                            <strong>성명 : </strong>
							<label>
								<input name="in_name" type="text" id="in_name" value="<%=in_name%>" style="width:80px;" class="no-input" readonly/>
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
                              <th colspan="3">교육&nbsp;과정명</th>
                              <th colspan="2">교육기관</th>
                              <th colspan="2">교육&nbsp;수료증No.</th>
                              <th colspan="2">교육&nbsp;기간</th>
                              <th colspan="3">교육&nbsp;주요&nbsp;내용</th>
                              <th>수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsEdu.EOF Or rsEdu.BOF Then
							Response.Write "<tr><td colspan='12' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsEdu.EOF
						%>
							<tr>
                              <td colspan="3"><%=rsEdu("edu_name")%>&nbsp;</td>
                              <td colspan="2"><%=rsEdu("edu_office")%>&nbsp;</td>
                              <td colspan="2"><%=rsEdu("edu_finish_no")%>&nbsp;</td>
                              <td colspan="2"><%=rsEdu("edu_start_date")%>∼<%=rsEdu("edu_end_date")%>&nbsp;</td>
                              <td colspan="3"><%=rsEdu("edu_comment")%>&nbsp;</td>
							  <td>
								<a href="#" onClick="pop_Window('/person/insa_edu_add.asp?edu_empno=<%=rsEdu("edu_empno")%>&edu_seq=<%=rsEdu("edu_seq")%>&emp_name=<%=in_name%>&u_type=U','교육사항 변경','scrollbars=yes,width=750,height=320')">수정</a>
							  </td>
							</tr>
						<%
								rsEdu.MoveNext()
							Loop
						End If
						rsEdu.Close() : Set rsEdu = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="pop_Window('/person/insa_edu_add.asp?edu_empno=<%=in_empno%>&emp_name=<%=in_name%>','교육사항 등록','scrollbars=yes,width=750,height=320')" class="btnType04">교육 등록</a>
					</div>
                    </td>
			      </tr>
				</table>
                <input type="hidden" name="edu_empno" value="<%=in_empno%>"/>
			</form>
		</div>
	</div>
	</body>
</html>