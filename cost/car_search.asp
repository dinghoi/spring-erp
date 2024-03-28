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
Dim rsCar, arrCar, title_line

objBuilder.Append "SELECT car_owner, car_no, car_name, oil_kind, last_km, "
objBuilder.Append "(SELECT MAX(end_km) FROM transit_cost WHERE car_no = cait.car_no) AS 'max_km' "
objBuilder.Append "FROM car_info AS cait "
objBuilder.Append "WHERE owner_emp_no = '"&user_id&"' "
objBuilder.Append "ORDER BY car_owner dESC, car_no "

Set rsCar = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsCar.EOF Then
	arrCar = rsCar.getRows()
End If
rsCar.Close() : Set rsCar = Nothing
DBConn.Close() : Set DBConn = Nothing

title_line = "차량 검색"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>차량 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>

		<script type="text/javascript">
			function car_list(car_owner,car_no,car_name,oil_kind,last_km){
				opener.document.frm.car_owner.value = car_owner;
				opener.document.frm.car_no.value = car_no;
				opener.document.frm.car_name.value = car_name;
				opener.document.frm.oil_kind.value = oil_kind;
				opener.document.frm.last_km.value = last_km;
				opener.document.frm.start_km.value = last_km;
				opener.document.frm.end_km.value = last_km;
				window.close();
			}
		</script>
	</head>
	<body>
		<div id="container">
			<h3 class="tit"><%=title_line%></h3>
			<fieldset class="srch">
				<legend>조회영역</legend>
				<dl>
					<dd>
						<p>
						<strong>운행자 정보 : </strong><%=user_name%>(<%=user_id%>)
						</p>
					</dd>
				</dl>
			</fieldset>
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableList">
					<colgroup>
						<col width="*" >
						<col width="20%" >
						<col width="20%" >
						<col width="20%" >
						<col width="20%" >
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">차량번호</th>
							<th scope="col">소유</th>
							<th scope="col">차종</th>
							<th scope="col">유종</th>
							<th scope="col">최종KM</th>
						</tr>
					</thead>
					<tbody>
					<%
					Dim i, car_owner, car_no, car_name, oil_kind, last_km, max_km

					If IsArray(arrCar) Then
						For i=LBound(arrCar) To UBound(arrCar, 2)
							car_owner = arrCar(0, i)
							car_no = arrCar(1, i)
							car_name = arrCar(2, i)
							oil_kind = arrCar(3, i)
							last_km = arrCar(4, i)
							max_km = arrCar(5, i)
					%>
						<tr>
							<td class="first">
							<a href="#" onClick="car_list('<%=car_owner%>','<%=car_no%>','<%=car_name%>','<%=oil_kind%>','<%=last_km%>');"><%=car_no%><%'=rs("car_no")%></a>
							</td>
							<td><%=car_owner%></td>
							<td><%=car_name%></td>
							<td><%=oil_kind%></td>
							<td><%=last_km%></td>
						</tr>
					<%
						Next
					Else
					%>
						<tr>
							<td class="first" colspan="5">조회된 내역이 없습니다</td>
						</tr>
					<%
					End If
					%>
					</tbody>
				</table>
			</div>
		</div>
	</body>
</html>