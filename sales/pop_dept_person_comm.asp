<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
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
Dim rs_emp
Dim dept, dt
Dim title_line

dept = Request("dept")	'사업부명
dt = Request("dt")	'검색 일자

title_line = dept & " 인력 리스트"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<script src="/java/jquery-1.9.1.js"></script>
		<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
	</head>
<body>
	<div style="margin:0px 10px 0px 10px;">
		<div id="container">
		<h3 class="stit">* <%=title_line%></h3>
			<table cellpadding="0" cellspacing="0" summary="" class="tableList">
			<colgroup>
				<col width="56%" >
				<col width="22%" >
				<col width="22%" >
			</colgroup>
			<thead>
				<tr>
					<th class="first" scope="col">이름</th>
					<th scope="col">구분</th>
					<th scope="col">손익 제외</th>
				</tr>
			</thead>
			<tbody>
			<%
			Dim sortCompany, sortGrade
			Dim totCnt : totCnt = 0

			sortCompany = "'케이원정보통신', '케이네트웍스', '코리아디엔씨', '에스유에이치'"
			sortGrade = "'사장', '부사장', '이사', '전무이사', '상무이사', '고문', '부장', '차장', '과장', '대리', '대리1급', '대리2급', '사원'"

			objBuilder.Append "SELECT emmt.emp_name, emmt.emp_job, emmt.emp_type, "
			objBuilder.Append "	IF(emmt.cost_except=2,'Y','N') AS cost_except "
			objBuilder.Append "FROM pay_month_give AS pmgt "
			objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
			objBuilder.Append "WHERE pmgt.pmg_id = '1' "
			objBuilder.Append "	AND emmt.cost_except IN ('0','1') /*손익적용*/ "
			objBuilder.Append "	AND pmgt.pmg_yymm = '"&dt&"' "
			objBuilder.Append "	AND emmt.emp_month = '"&dt&"' "
			objBuilder.Append "	AND pmgt.mg_saupbu = '"&dept&"' "
			objBuilder.Append "	AND emmt.emp_saupbu = '"&dept&"' "
			objBuilder.Append "ORDER BY FIELD(emmt.emp_job, "&sortGrade&") ASC "

			Set rs_emp = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If Not(rs_emp.BOF Or rs_emp.EOF) Then
				Do Until rs_emp.EOF
					totCnt = totCnt + 1
			%>
				<tr>
					<td><%=rs_emp("emp_name")%>&nbsp;<%=rs_emp("emp_job")%></td>
					<td><%=rs_emp("emp_type")%></td>
					<td><%=rs_emp("cost_except")%></td>
				</tr>
			<%
					rs_emp.MoveNext()
				Loop
			%>
				<tr bgcolor="#FFE8E8">
					<td class="first" colspan="2">총인원</td>
					<td class="center"><%=FormatNumber(totCnt, 0)%>&nbsp;명</td>
				</tr>
			<%
			Else
			%>
				<tr>
					<td colspan="5">해당 데이터가 없습니다.</td>
				</tr>
			<%
			End If

			rs_emp.Close()
			Set rs_emp = Nothing

			DBConn.Close()
			Set DBConn = Nothing
			%>
			</tbody>
			</table>
		</div>
	</div>
</body>
</html>
