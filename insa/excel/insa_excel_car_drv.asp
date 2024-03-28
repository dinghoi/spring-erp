<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### 작업 내역
'===================================================
' 허정호_20210723 :
'	- 신규 페이지 작성 및 코드 정리

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
Dim view_condi, from_date, to_date, curr_date, title_line
Dim savefilename, rsTran

view_condi = Request.QueryString("view_condi")
from_date = Request.QueryString("from_date")
to_date = Request.QueryString("to_date")

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

title_line = CStr(from_date) & "~ " & CStr(to_date) & " 차량 운행현황"

savefilename = title_line & ".xls"

Call ViewExcelType(savefilename)

objBuilder.Append "SELECT trct.car_owner, trct.car_name, trct.mg_ce_id, trct.car_no, "
objBuilder.Append "	trct.start_km, trct.end_km, trct.far, trct.run_date, "
objBuilder.Append "	IF(car_owner = '대중교통', trct.transit, trct.oil_kind) AS 'tran_type', "
objBuilder.Append "	trct.start_company, trct.start_point, trct.end_company, "
objBuilder.Append "	trct.end_point, trct.run_memo, trct.fare, trct.oil_price,  "
objBuilder.Append "	trct.parking, trct.toll, "
objBuilder.Append "	emtt.emp_name "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON trct.mg_ce_id = emtt.emp_no "
objBuilder.Append "WHERE run_date >= '"&from_date&"' AND run_date <= '"&to_date&"' "

If view_condi <> "" Then
	objBuilder.Append "	AND car_no = '"&view_condi&"' "
End If

objBuilder.Append "ORDER BY trct.car_no, trct.run_date, trct.run_seq ASC "

Set rsTran = Server.CreateObject("ADODB.RecordSet")
rsTran.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
	<tr bgcolor="#EFEFEF" class="style11">
	<td colspan="17" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
	</tr>
	<tr bgcolor="#EFEFEF" class="style11">
		<td><div align="center" class="style1">차량번호</div></td>
		<td><div align="center" class="style1">차종</div></td>

		<td><div align="center" class="style1">운행일자</div></td>
		<td><div align="center" class="style1">운행자</div></td>
		<td><div align="center" class="style1">구분</div></td>
		<td><div align="center" class="style1">유종/대중교통</div></td>
		<td><div align="center" class="style1">출발업체명</div></td>
		<td><div align="center" class="style1">출발지</div></td>
		<td><div align="center" class="style1">출발KM</div></td>
		<td><div align="center" class="style1">도착업체명</div></td>
		<td><div align="center" class="style1">도착지</div></td>
		<td><div align="center" class="style1">도착KM</div></td>
		<td><div align="center" class="style1">운행목적</div></td>
		<td><div align="center" class="style1">대중교통경비</div></td>
		<td><div align="center" class="style1">주유금액</div></td>
		<td><div align="center" class="style1">주차비</div></td>
		<td><div align="center" class="style1">통행료</div></td>
	</tr>
	<%
	Dim car_name, mg_ce_id, emp_name, drv_owner_emp_name
	Dim start_view, end_view, run_km

	Do Until rsTran.EOF
		car_name = rsTran("car_name")
		mg_ce_id = rsTran("mg_ce_id")
		emp_name = rsTran("emp_name")

		If emp_name = "" Or IsNull(emp_name) Then
			drv_owner_emp_name = rsTran("mg_ce_id")
		End If

		If rsTran("start_km") = "" Or IsNull(rsTran("start_km")) Then
			start_view = 0
		Else
			start_view = rsTran("start_km")
		End If

		If rsTran("end_km") = "" Or IsNull(rsTran("end_km")) Then
			end_view = 0
		Else
			end_view = rsTran("end_km")
		End If

		run_km = rsTran("far")
	%>
	<tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rsTran("car_no")%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsTran("run_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=drv_owner_emp_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsTran("car_owner")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsTran("tran_type")%></td>
    <td width="115"><div align="center" class="style1"><%=rsTran("start_company")%></div></td>
    <td width="200"><div align="left" class="style1"><%=rsTran("start_point")%></div></td>
    <td width="115"><div align="right" class="style1"><%=FormatNumber(start_view,0)%></div></td>
    <td width="115"><div align="center" class="style1"><%=rsTran("end_company")%></div></td>
    <td width="200"><div align="left" class="style1"><%=rsTran("end_point")%></div></td>
    <td width="115"><div align="right" class="style1"><%=FormatNumber(end_view,0)%></div></td>
    <td width="200"><div align="left" class="style1"><%=rsTran("run_memo")%></div></td>
    <td width="115"><div align="right" class="style1"><%=FormatNumber(rsTran("fare"),0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=FormatNumber(rsTran("oil_price"),0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=FormatNumber(rsTran("parking"),0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=FormatNumber(rsTran("toll"),0)%></div></td>
  </tr>
<%
		rsTran.MoveNext()
	Loop
	rsTran.Close() : Set rsTran = Nothing
%>
</table>
</body>
</html>
<!--#include virtual="/common/inc_footer.asp" -->