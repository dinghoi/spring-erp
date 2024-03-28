<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
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
'==================================================
Dim title_line, savefilename, i, rsLog
Dim from_date, to_date

from_date = f_Request("from_date")
to_date = f_Request("to_date")

title_line = "시스템 로그 정보"
savefilename = title_line & ".xls"

'엑셀 다운로드 설정
Call ViewExcelType(savefilename)

objBuilder.Append "SELECT memt.emp_no, memt.user_name, memt.user_grade, memt.position, "
objBuilder.Append "	memt.emp_company, memt.bonbu, memt.saupbu, memt.team, memt.org_name, "
objBuilder.Append "	logt.remote_ip, logt.menu_name, logt.menu_title, logt.excel_yn, logt.reg_date "
objBuilder.Append "FROM emp_sys_log AS logt "
objBuilder.Append "INNER JOIN memb AS memt ON logt.emp_no = memt.emp_no "
objBuilder.Append "WHERE logt.reg_date BETWEEN '"&from_date&"' AND '"&to_date&"' "
objBuilder.Append "ORDER BY logt.reg_date DESC "

'Response.write objBuilder.ToString()


Set rsLog = DBConn.Execute(objBuilder.ToString)
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">사번</th>
								<th scope="col">성명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">회사</th>
								<th scope="col">본부</th>
								<th scope="col">사업부</th>
								<th scope="col">팀</th>
								<th scope="col">조직명</th>
								<th scope="col">접근IP</th>
								<th scope="col">접근메뉴</th>
								<th scope="col">메뉴명</th>
								<th scope="col">엑셀다운로드</th>
								<th scope="col">접근시간</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0

						Do Until rsLog.EOF
							i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rsLog("emp_no")%></td>
								<td><%=rsLog("user_name")%></td>
								<td><%=rsLog("user_grade")%></td>
								<td><%=rsLog("position")%></td>
								<td><%=rsLog("emp_company")%></td>
								<td><%=rsLog("bonbu")%></td>
							  	<td><%=rsLog("saupbu")%></td>
							  	<td><%=rsLog("team")%></td>
							  	<td><%=rsLog("org_name")%></td>
								<td><%=rsLog("remote_ip")%></td>
								<td><%=rsLog("menu_name")%></td>
								<td><%=rsLog("menu_title")%></td>
								<td><%=rsLog("excel_yn")%></td>
								<td><%=rsLog("reg_date")%></td>
							</tr>
						<%
							rsLog.MoveNext()
						Loop
						rsLog.Close() : Set rsLog = Nothing
						%>
						</tbody>
					</table>
				</div>
			</div>
		</div>
	</body>
</html>
<%
DBConn.Close() : Set DBConn = Nothing
%>