<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim srchEmpMonth, savefilename, pre_date, rsEmp, title_line
Dim pre_yyyy, pre_mm

srchEmpMonth = Request.QueryString("srchEmpMonth")

title_line = srchEmpMonth & " 월 추가 인원"
savefilename = title_line & ".xls"

pre_yyyy = Left(srchEmpMonth, 4)
pre_mm = Right(srchempMonth, 2) - 1

If pre_mm = 0 Then
	pre_mm = 12
End If

'이전 년월
If pre_mm < 10 Then
	pre_date = CStr(pre_yyyy&"0"&pre_mm)
Else
	pre_date = CStr(pre_yyyy & pre_mm)
End If

'엑셀 지정
Call ViewExcelType(savefilename)

objBuilder.Append "SELECT emmt.emp_first_date, emmt.emp_no, emmt.emp_name, emp_job, "
objBuilder.Append "	emmt.emp_reside_company, emp_reside_place, emmt.cost_center, emmt.cost_group, "
objBuilder.Append "	eomm.org_name, eomm.org_company, eomm.org_bonbu, eomm.org_saupbu, eomm.org_team "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&srchEmpMonth&"' "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomm ON emmt.emp_org_code = eomm.org_code "
objBuilder.Append "	AND eomm.org_month = '"&srchEmpMonth&"'	"
objBuilder.Append "WHERE pmgt.pmg_id = '1' "
objBuilder.Append "	AND pmgt.pmg_yymm = '"&srchEmpMonth&"' "
objBuilder.Append "	AND emmt.emp_no NOT IN ("
objBuilder.Append "		SELECT emmt.emp_no "
objBuilder.Append "		FROM pay_month_give AS pmgt "
objBuilder.Append "		INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "			AND emmt.emp_month = '"&pre_date&"'	"
objBuilder.Append "		WHERE pmgt.pmg_id = '1' "
objBuilder.Append "			AND pmgt.pmg_yymm = '"&pre_date&"') "
objBuilder.Append "ORDER BY emmt.emp_first_date ASC	"

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title><%=title_line%></title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">입사일자</th>
								<th scope="col">사번</th>
								<th scope="col">성명</th>
								<th scope="col">직급</th>
								<th scope="col">조직명</th>
								<th scope="col">회사</th>
								<th scope="col">본부</th>
								<th scope="col">사업부</th>
								<th scope="col">팀</th>
								<th scope="col">상주처</th>
								<th scope="col">상주회사</th>
								<th scope="col">비용구분</th>
								<th scope="col">비용그룹</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsEmp.EOF
						%>
							<tr>
								<td class="first"><%=rsEmp("emp_first_date")%></td>
								<td><%=rsEmp("emp_no")%></td>
								<td><%=rsEmp("emp_name")%></td>
								<td><%=rsEmp("emp_job")%></td>
								<td><%=rsEmp("org_name")%></td>
								<td><%=rsEmp("org_company")%></td>
								<td><%=rsEmp("org_bonbu")%></td>
								<td><%=rsEmp("org_saupbu")%></td>
								<td><%=rsEmp("org_team")%></td>
								<td><%=rsEmp("emp_reside_place")%></td>
								<td><%=rsEmp("emp_reside_company")%></td>
								<td><%=rsEmp("cost_center")%></td>
								<td><%=rsEmp("cost_group")%></td>
							</tr>
						<%
							rsEmp.MoveNext()
						Loop
						rsEmp.Close() : Set rsEmp = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>