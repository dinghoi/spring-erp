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
Dim title_line, rsEdu, arrEdu

emp_no = f_Request("emp_no")

title_line = "교육 사항"

objBuilder.Append "CALL USP_INSA_EDU_INFO('"&emp_no&"')"

Call Rs_Open(rsEdu, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rsEdu.EOF Then
	arrEdu = rsEdu.getRows()
End If

Call Rs_Close(rsEdu)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>인사 관리 시스템</title>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

	<style type="text/css">
		.no-input{
			color:gray;
			background-color:#E0E0E0;
			border:1px solid #999999;
		}
	</style>
</head>
<body oncontextmenu="return false" ondragstart="return false">
	<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<fieldset class="srch">
				<legend>조회영역</legend>
				<dl>
					<dd>
						<p>
						<strong>사번 : </strong>
							<label>
							<input name="in_empno" type="text" id="in_empno" value="<%=emp_no%>" style="width:80px;" class="no-input" readonly/>
							</label>
						<strong>성명 : </strong>
							<label>
							<input name="in_name" type="text" id="in_name" value="<%Call EmpInfo_Name(emp_no)%>" style="width:80px;" class="no-input" readonly/>
							</label>
						</p>
					</dd>
				</dl>
			</fieldset>
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableList">
					<colgroup>
						<col width="14%" >
						<col width="14%" >
						<col width="10%" >
						<col width="14%" >
						<col width="14%" >
						<col width="8%" >
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">교육&nbsp;과정명</th>
							<th scope="col">교육기관</th>
							<th scope="col">교육수료증</th>
							<th scope="col">교육기간</th>
							<th colspan="2" scope="col">교육&nbsp;&nbsp;주요&nbsp;내용</th>
						</tr>
					</thead>
					<tbody>
					<%
					DBConn.Close() : Set DBConn = Nothing
					Dim i

					If IsArray(arrEdu) Then
						Dim edu_name, edu_office, edu_finish_no, edu_start_date, edu_end_date, edu_comment

						For i = LBound(arrEdu) To UBound(arrEdu, 2)
							edu_name = arrEdu(0, i)
							edu_office = arrEdu(1, i)
							edu_finish_no = arrEdu(2, i)
							edu_start_date = arrEdu(3, i)
							edu_end_date = arrEdu(4, i)
							edu_comment = arrEdu(5, i)
					%>
						<tr>
							<td><%=edu_name%>&nbsp;</td>
							<td><%=edu_office%>&nbsp;</td>
							<td><%=edu_finish_no%>&nbsp;</td>
							<td><%=edu_start_date%>&nbsp;∼&nbsp;<%=edu_end_date%>&nbsp;</td>
							<td colspan="2" class="left"><%=edu_comment%>&nbsp;</td>
						</tr>
					<%
						Next
					Else
					%>
						<tr>
							<td class="first" colspan="5" style="height:30px;">조회된 내역이 없습니다</td>
						</tr>
					<%
					End If
					%>
					</tbody>
				</table>
			</div>
		</div>
	</div>
	<br>
	<div align="right">
		<a href="#" class="btnType04" onclick="close_win();" >닫기</a>&nbsp;&nbsp;
	</div>
	<br>
</body>
</html>