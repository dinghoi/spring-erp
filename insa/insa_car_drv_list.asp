<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### 작업 내역
'===================================================
' 허정호_20210722 :
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
Dim be_pg, from_date, to_date, page, view_condi
Dim curr_dd, start_page, pgsize, stpage
Dim rsCount, total_record, whereSql, orderSql, rsTran
Dim title_line, total_page, str_param

be_pg = "/insa/insa_car_drv_list.asp"

from_date = f_Request("from_date")
to_date = f_Request("to_date")
page = f_Request("page")
view_condi = f_Request("view_condi")

curr_dd = CStr(DatePart("d", Now()))

If from_date = "" Or IsNull(from_date) Then
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
End If

If to_date = "" Or IsNull(to_date) Then
	to_date = Mid(CStr(Now()), 1, 10)
End If

'If view_condi = "" Then
'	view_condi = "전체"
'	curr_dd = CStr(DatePart("d", Now()))
'	to_date = Mid(CStr(Now()), 1, 10)
'	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
'End If

pgsize = 10 ' 화면 한 페이지
str_param = "&from_date="&from_date&"&to_date="&to_date

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

whereSql = "WHERE run_date >= '"&from_date&"' AND run_date <= '"&to_date&"' "
orderSql = "ORDER BY car_no,run_date,run_seq ASC "

If view_condi <> "" Then
	whereSql = whereSql & "	AND car_no='"&view_condi&"' "
End If

objBuilder.Append "SELECT COUNT(*) FROM transit_cost "
objBuilder.Append whereSql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

objBuilder.Append "SELECT trct.mg_ce_id, trct.start_km, trct.end_km, trct.far, trct.car_no, "
objBuilder.Append "	trct.run_date, trct.car_owner, trct.transit, trct.oil_kind, "
objBuilder.Append "	IF(trct.car_owner = '대중교통', trct.transit, trct.oil_kind) AS 'tran_type', "
objBuilder.Append "	trct.start_company, trct.start_point, trct.end_company, trct.end_point, "
objBuilder.Append "	trct.run_memo, trct.fare, trct.oil_price, trct.parking, trct.toll, "
objBuilder.Append "	emtt.emp_name "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON trct.mg_ce_id = emtt.emp_no "
objBuilder.Append whereSql & orderSql
objBuilder.Append "LIMIT "&stpage&","&pgsize

Set rsTran = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = view_condi & " - 차량 운행현황 "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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

		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
			}

			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd");
				$( "#datepicker" ).datepicker("setDate", "<%=from_date%>");
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd");
				$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>");
			});

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit();
				}
			}

			//엑셀 다운로드[허정호_20210721]
			function carDrvExcel(c_view, f_date, t_date){
				var url = '/insa/excel/insa_excel_car_drv.asp';
				var param;

				param = '?view_condi='+c_view+'&from_date='+f_date+'&to_date='+t_date;

				location.href = url + param;
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="<%=be_pg%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>

								<strong>차량번호 : </strong>
                                <label>
									<input name="view_condi" type="text" value="<%=view_condi%>" style="width:100px; text-align:left" >
                                </label>

                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
                            <col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="8%" >
							<col width="9%" >
							<col width="5%" >
							<col width="8%" >
							<col width="*" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="4%" >
                		</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">차량번호</th>
                                <th rowspan="2" scope="col">운행일자</th>
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
						Dim drv_owner_emp_name, start_view, end_view, run_km

						Do Until rsTran.EOF
							drv_owner_emp_name = rsTran("emp_name")

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
							<tr>
								<td class="first"><%=rsTran("car_no")%></td>
                                <td><%=rsTran("run_date")%></td>
								<td><%=drv_owner_emp_name%></td>
								<td><%=rsTran("car_owner")%></td>
								<td><%=rsTran("tran_type")%></td>
								<td><%=rsTran("start_company")%>&nbsp;</td>
								<td class="left"><%=rsTran("start_point")%></td>
								<td class="right"><%=FormatNumber(start_view, 0)%></td>
								<td><%=rsTran("end_company")%>&nbsp;</td>
								<td class="left"><%=rsTran("end_point")%></td>
								<td class="right"><%=FormatNumber(end_view, 0)%></td>
								<td><%=rsTran("run_memo")%></td>
								<td class="right"><%=FormatNumber(rsTran("fare"), 0)%></td>
								<td class="right"><%=FormatNumber(rsTran("oil_price"), 0)%></td>
								<td class="right"><%=FormatNumber(rsTran("parking"), 0)%></td>
								<td class="right"><%=FormatNumber(rsTran("toll"), 0)%></td>
							</tr>
						<%
							rsTran.MoveNext()
						Loop
						rsTran.close() : Set rsTran = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
						<a href="#" class="btnType04" onclick="carDrvExcel('<%=view_condi%>', '<%=from_date%>', '<%=to_date%>');">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
					<%
					'page navigator[허정호_20210720]
					Call Page_Navi(page, be_pg, str_param, total_page)

					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>

