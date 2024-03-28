<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim owner_view, field_check, field_view, curr_date
Dim owner_gubun, savefilename
Dim owner_sql, rsCar, sqlWhereStr
Dim arrCar

owner_view = Request("owner_view")
field_check = Request("field_check")
field_view = Request("field_view")

curr_date = datevalue(mid(cstr(now()),1,10))

Select Case owner_view
	Case "C"
		owner_gubun = "회사 "
	Case "P"
		owner_gubun = "개인 "
	Case Else
		owner_gubun = "전체"
End Select

savefilename = owner_gubun & " 차량 현황 " & CStr(curr_date) & ".xls"

Call ViewExcelType(savefilename)

objBuilder.Append "CALL USP_INSA_CAR_INFO_SELECT('"&owner_view&"', '"&field_check&"', '"&field_view&"');"
Set rsCar = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsCar.EOF Then
	arrCar = rsCar.getRows()
End If
rsCar.Close() : Set rsCar = Nothing
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=owner_gubun%> &nbsp;차량 현황&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">차량번호</div></td>
    <td><div align="center" class="style1">차종</div></td>
    <td><div align="center" class="style1">연식</div></td>
    <td><div align="center" class="style1">유류종류</div></td>
	<td><div align="center" class="style1">소유</div></td>
    <td><div align="center" class="style1">차량소유회사</div></td>
    <td><div align="center" class="style1">사용부서</div></td>
    <td><div align="center" class="style1">용도</div></td>
    <td><div align="center" class="style1">운행자</div></td>
    <td><div align="center" class="style1">차량등록일</div></td>
    <td><div align="center" class="style1">운행Km</div></td>
    <td><div align="center" class="style1">보험기간</div></td>
    <td><div align="center" class="style1">보험회사</div></td>
    <td><div align="center" class="style1">보험료</div></td>
    <td><div align="center" class="style1">최종점검일</div></td>
    <td><div align="center" class="style1">차량상태</div></td>
    <td><div align="center" class="style1">차량정보</div></td>
    <%' 아래부분은 일단 막아놓구... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">입고 세부내역 </div> %>
    <%'</div></td> %>
  </tr>
    <%
	Dim i, car_no, car_name, car_year, oil_kind
	Dim car_owner, car_company, car_use_dept, car_use
	Dim owner_emp_name, owner_emp_no, car_reg_date, last_km, insurance_date
	Dim insurance_company, insurance_amt, last_check_date, car_status, car_comment

	If IsArray(arrCar) Then
		For i = LBound(arrCar) To UBound(arrCar, 2)
			car_no = arrCar(0, i)
			car_name = arrCar(1, i)
			car_year = arrCar(2, i)
			oil_kind = arrCar(3, i)
			car_owner = arrCar(4, i)
			car_company = arrCar(5, i)
			car_use_dept = arrCar(6, i)
			car_use = arrCar(7, i)
			owner_emp_name = arrCar(8, i)
			owner_emp_no = arrCar(9, i)
			car_reg_date = arrCar(10, i)
			last_km = arrCar(11, i)
			insurance_date = arrCar(12, i)
			insurance_company = arrCar(13, i)
			insurance_amt = arrCar(14, i)
			last_check_date = arrCar(15, i)
			car_status = arrCar(16, i)
			car_comment = arrCar(17, i)
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=car_no%></div></td>
    <td width="145"><div align="center" class="style1"><%=car_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_year%></div></td>
    <td width="115"><div align="center" class="style1"><%=oil_kind%></div></td>
	<td width="115"><div align="center" class="style1"><%=car_owner%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_company%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_use_dept%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_use%></div></td>
    <td width="145"><div align="center" class="style1"><%=owner_emp_name%>(<%=owner_emp_no%>)&nbsp;</div></td>
    <td width="145"><div align="center" class="style1"><%=car_reg_date%></div></td>
    <td width="145"><div align="center" class="style1"><%=FormatNumber(last_km, 0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=insurance_date%></div></td>
    <td width="145"><div align="center" class="style1"><%=insurance_company%></div></td>
    <td width="145"><div align="center" class="style1"><%=FormatNumber(insurance_amt, 0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=last_check_date%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_status%></div></td>
    <td width="200"><div align="center" class="style1"><%=car_comment%></div></td>
  </tr>
	<%
		Next
	End If
	%>
</table>
</body>
</html>
<!--#include virtual="/common/inc_footer.asp" -->