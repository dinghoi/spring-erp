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
Dim car_no, car_name, car_year, car_reg_date
Dim title_line, be_pg, pgsize, page, start_page
Dim stpage, rsCount, total_record, rsTran, total_page
Dim str_param

car_no = f_Request("car_no")
car_name = f_Request("car_name")
car_year = f_Request("car_year")
car_reg_date = f_Request("car_reg_date")
page = f_Request("page")

title_line = " 차량 운행 현황 "
be_pg = "/insa/insa_car_drv_view.asp"
pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
str_param = "&car_no="&car_no&"&car_name="&car_name&"&car_year="&car_year&"&car_reg_date="&car_reg_date

'car_drv 테이블에 저장된 데이터가 없고, 사용되는 페이지 없음, transit_cost로 대체 사용[허정호_20210721]
'Sql = "SELECT count(*) FROM car_drv where drv_car_no = '"&car_no&"'"
objBuilder.Append "SELECT COUNT(*) FROM transit_cost "
objBuilder.Append "WHERE car_no = '"&car_no&"' "

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

'sql = "select * from car_drv where drv_car_no = '" + car_no + "' ORDER BY drv_car_no,drv_date,drv_seq DESC limit "& stpage & "," &pgsize
objBuilder.Append "SELECT run_date, user_name, car_owner, "
objBuilder.Append "	IF(car_owner = '대중교통', transit, oil_kind) AS 'tran_type', "
objBuilder.Append "	start_company, start_point, start_km, "
objBuilder.Append "	end_company, end_point, end_km, run_memo, "
objBuilder.Append "	fare, oil_price, parking, toll "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE car_no = '"&car_no&"' "
objBuilder.Append "ORDER BY car_no, run_date, run_seq "
objBuilder.Append "LIMIT "& stpage & "," &pgsize

Set rsTran = Server.CreateObject("ADODB.RecordSet")
rsTran.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_drv_view.asp?car_no=<%=car_no%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>차량번호 : </strong>
								<label>
        						<input name="in_carno" type="text" id="in_carno" value="<%=car_no%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            <strong>차종/연식/취득일 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=car_name%>" style="width:100px; text-align:left" readonly="true">
                                -
                                <input name="in_year" type="text" id="in_year" value="<%=car_year%>" style="width:70px; text-align:left" readonly="true">
                                 -
                                <input name="car_reg_date" type="text" id="car_reg_date" value="<%=car_reg_date%>" style="width:70px; text-align:left" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="10%" >
							<col width="10%" >
							<col width="5%" >
							<col width="10%" >
							<col width="10%" >
							<col width="5%" >
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">운행일자</th>
								<th rowspan="2" scope="col">운행자</th>
								<th rowspan="2" scope="col">구분</th>
								<th rowspan="2" scope="col">유종<br>/<br>대중<br>교통</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">출 발</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">도 착</th>
								<th rowspan="2" scope="col">운행목적</th>
								<th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">경 비 </th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">업체명</th>
								<th scope="col">출발지</th>
								<th scope="col">출발KM</th>
								<th scope="col">업체명</th>
								<th scope="col">도착지</th>
								<th scope="col">도착KM</th>
								<th scope="col">대중교통</th>
								<th scope="col">주유금액</th>
								<th scope="col">주차비</th>
								<th scope="col">통행료</th>
							</tr>
						</thead>
						<tbody>
						<%
							Do Until rsTran.EOF or rsTran.BOF
						%>
							<tr>
								<td class="first"><%=rsTran("run_date")%></td>
								<td><%=rsTran("user_name")%></td>
								<td><%=rsTran("car_owner")%></td>
								<td><%=rsTran("tran_type")%></td>
								<td><%=rsTran("start_company")%></td>
								<td><%=rsTran("start_point")%></td>
								<td class="right"><%=FormatNumber(rsTran("start_km"), 0)%></td>
								<td><%=rsTran("end_company")%></td>
								<td><%=rsTran("end_point")%></td>
								<td class="right"><%=FormatNumber(rsTran("end_km"), 0)%></td>
								<td><%=rsTran("run_memo")%></td>
								<td class="right"><%=FormatNumber(rsTran("fare"), 0)%></td>
								<td class="right"><%=FormatNumber(rsTran("oil_price"), 0)%></td>
								<td class="right"><%=FormatNumber(rsTran("parking"), 0)%></td>
								<td class="right"><%=FormatNumber(rsTran("toll"), 0)%></td>
							</tr>
							<%
								rsTran.MoveNext()
							Loop
							rsTran.Close() : Set rsTran = Nothing
							%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <%
					'page navigator[허정호_20210720]
					Call Page_Navi(page, be_pg, str_param, total_page)
					%>
                    </td>
				    <td width="20%">
					<div align="right">
						<a href="#" class="btnType04" onclick="javascript:toclose();" >닫기</a>&nbsp;&nbsp;
					</div>
                    </td>
			      </tr>
			  </table>
         </div>
	</form>
	  </div>
	</body>
</html>