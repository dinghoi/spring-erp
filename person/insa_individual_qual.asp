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
Dim rs_emp, title_line, rsQual

title_line = "자격 사항"

in_name = user_name
in_empno = user_id

If f_toString(Request.Form("in_empno"), "")  <> "" Then
	objBuilder.Append "SELECT emp_name FROM emp_master WHERE emp_no= '"&in_empno&"';"

	Set rs_emp = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not rs_emp.EOF Then
		in_name = rs_emp("emp_name")
	Else
		Response.Write "<script type='text/javascript'>"
		Response.Write "	alert('등록된 직원이 아닙니다.');"
		Response.Write"</script>"
		Response.End
	End If
	rs_emp.Close() : Set rs_emp = Nothing
End If

objBuilder.Append "SELECT qual_type,qual_grade, qual_pass_date, qual_org, qual_no, "
objBuilder.Append "	qual_passport, qual_seq, qual_empno "
objBuilder.Append "FROM emp_qual "
objBuilder.Append "WHERE qual_empno = '"&in_empno&"' "
objBuilder.Append "ORDER BY qual_empno, qual_seq ASC;"

Set rsQual = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
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
								<input name="in_empno" type="text" id="in_empno" value="<%=in_empno%>" style="width:80px;" class="no-input" readonly="true"/>
							</label>
                            <strong>성명 : </strong>
							<label>
								<input name="in_name" type="text" id="in_name" value="<%=in_name%>" style="width:80px;" class="no-input" readonly="true"/>
							</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="8%" >
							<col width="9%" >
							<col width="15%" >
							<col width="*" >
							<col width="15%" >
                            <col width="5%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                            <th>자격증 종목</th>
                            <th>등급</th>
                            <th>합격년월일</th>
                            <th>발급 기관명</th>
                            <th>자격 등록번호</th>
                            <th>경력수첩No.</th>
                            <th>순번</th>
                            <th>수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsQual.EOF Or rsQual.BOF Then
							Response.Write "<tr><td colspan='8' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsQual.EOF
						%>
							<tr>
								<td><%=rsQual("qual_type")%>&nbsp;</td>
								<td><%=rsQual("qual_grade")%>&nbsp;</td>
								<td><%=rsQual("qual_pass_date")%>&nbsp;</td>
								<td><%=rsQual("qual_org")%>&nbsp;</td>
								<td><%=rsQual("qual_no")%>&nbsp;</td>
								<td><%=rsQual("qual_passport")%>&nbsp;</td>
								<td class="right"><%=rsQual("qual_seq")%>&nbsp;</td>
								<td>
									<a href="#" onClick="pop_Window('/person/insa_individual_qual_add.asp?qual_empno=<%=rsQual("qual_empno")%>&qual_seq=<%=rsQual("qual_seq")%>&emp_name=<%=in_name%>&u_type=U','자격사항 변경','scrollbars=yes,width=750,height=300')">수정</a>
								</td>
							</tr>
						<%
								rsQual.MoveNext()
							Loop
						End If
						rsQual.Close() : Set rsQual = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="pop_Window('/person/insa_individual_qual_add.asp?qual_empno=<%=in_empno%>&emp_name=<%=in_name%>','자격사항 등록','scrollbars=yes,width=750,height=300')" class="btnType04">자격사항 등록</a>
					</div>
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="qual_empno" value="<%=in_empno%>"/>
			</form>
		</div>
	</div>
	</body>
</html>