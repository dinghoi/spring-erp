<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/query/person/query_person_cost.asp" -->
<%
	On Error Resume Next
	
	cost_month=Request("cost_month")
	be_yy = int(mid(cost_month,1,4))
	be_mm = int(mid(cost_month,5) - 1)
	
	if be_mm = 0 then
		be_month = cstr(be_yy - 1) + "12"
	else
  	be_month = cstr(be_yy) + right("0" + cstr(be_mm),2)
	end if
	
	month_view = cstr(mid(cost_month,1,4)) + "년 " + cstr(mid(cost_month,5,2)) + "월"
	be_month_view = cstr(mid(be_month,1,4)) + "년 " + cstr(mid(be_month,5,2)) + "월"

' 전월추가
'sql = "select * from person_cost where cost_month = '"&be_month&"' and emp_no = '"&user_id&"'"
'sql = "call COST_PERSON_01('" & mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01" & "','"&user_id&"',@ret)"       

'set rs=dbconn.execute(sql)
  arParams = Array(mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _ 
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _ 
                    mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01", _
                    user_id)
	Set cmd = server.CreateObject("ADODB.Command")
	cmd.CommandText = query_person_cost
'Response.write	query_person_cost&"<br><br>"

  for i = 0 to 8
     'Response.write	arParams(i)&"<br>"
  Next
      
	Set cmd.ActiveConnection = dbconn
	Set rs = cmd.execute(,arParams,1)

	if rs.eof or rs.bof then
		be_general_cnt = 0
		be_general_cost = 0
		be_overtime_cnt = 0 
		be_overtime_cost = 0 
		be_somopum_cost = 0 
		be_fare_cnt = 0		 
		be_fare_cost = 0		 
		be_oil_cash_cost = 0	 
		be_repair_cost = 0
		be_parking_cost = 0 
		be_toll_cost = 0
		be_card_cost = 0
		be_card_cost_vat = 0	 
		be_return_cash = 0
		be_tot_km = 0
		be_tot_cost = 0
		be_card_price = 0
		be_juyoo_card_price = 0
		be_cash_tot_cost = 0
	  be_general_cost_01 =0  '차량유지비' 
	  be_general_cost_02 =0  '여비교통비' 
	  be_general_cost_03 =0  '복리후생비' 
	  be_general_cost_04 =0  '접대비'     
	  be_general_cost_05 =0  '회의비'     
	  be_general_cost_06 =0  '사무용품비' 
	  be_general_cost_07 =0  '소모품비'   
	  be_general_cost_08 =0  '운반비'     
	  be_general_cost_09 =0  '통신비'     
	  be_general_cost_10 =0  '국내출장비' 
	  be_general_cost_11 =0  '수선비'     
	  be_general_cost_12 =0  '지급수수료' 
  else
		be_general_cnt = rs("general_cnt")	 
		be_general_cost = rs("general_cost")	 
		be_overtime_cnt = rs("overtime_cnt")	 
		be_overtime_cost = rs("overtime_cost")	 
		gas_km = cdbl(rs("gas_km"))  
	  gas_unit = cdbl(rs("gas_unit"))  
	  gas_cost = cdbl(rs("gas_cost"))  
	  gasol_km = cdbl(rs("gasol_km"))  
	  gasol_unit = cdbl(rs("gasol_unit"))  
	  gasol_cost = cdbl(rs("gasol_cost"))  
	  diesel_km = cdbl(rs("diesel_km"))  
	  diesel_unit = cdbl(rs("diesel_unit"))  
	  diesel_cost = cdbl(rs("diesel_cost"))  
	  be_somopum_cost = cdbl(rs("somopum_cost"))	 
		be_fare_cost = rs("fare_cost")	 		 
		be_oil_cash_cost = rs("oil_cash_cost")	 
		be_repair_cost = rs("repair_cost")	 
		be_parking_cost = rs("parking_cost")	 
		be_toll_cost = rs("toll_cost")	 
		be_card_cost = rs("card_cost")	 
		be_card_cost_vat = rs("card_cost_vat")	 
		be_return_cash = rs("return_cash")	 
		be_tot_km = gas_km + diesel_km + gasol_km
		be_tot_cost = gas_cost + diesel_cost + gasol_cost
		be_card_price = be_card_cost + be_card_cost_vat
		be_juyoo_card_price = rs("juyoo_card_cost") + rs("juyoo_card_cost_vat")
		'be_cash_tot_cost = be_general_cost + gas_cost + diesel_cost + gasol_cost + be_somopum_cost + be_fare_cost + be_oil_cash_cost + be_toll_cost + be_parking_cost
		be_company_yn = cdbl(rs("company_yn"))
  
  	if be_company_yn > 0 then
    	be_cash_tot_cost =  be_fare_cost + be_oil_cash_cost + be_toll_cost + be_parking_cost
  	else
    	be_cash_tot_cost = be_general_cost + gas_cost + diesel_cost + gasol_cost + be_somopum_cost + be_fare_cost + be_oil_cash_cost + be_toll_cost + be_parking_cost
  	end if

	  be_general_cost_01 = cdbl(rs("general_cost_01"))
	  be_general_cost_02 = cdbl(rs("general_cost_02"))
	  be_general_cost_03 = cdbl(rs("general_cost_03"))
	  be_general_cost_04 = cdbl(rs("general_cost_04"))
	  be_general_cost_05 = cdbl(rs("general_cost_05"))
	  be_general_cost_06 = cdbl(rs("general_cost_06"))
	  be_general_cost_07 = cdbl(rs("general_cost_07"))
	  be_general_cost_08 = cdbl(rs("general_cost_08"))
	  be_general_cost_09 = cdbl(rs("general_cost_09"))
	  be_general_cost_10 = cdbl(rs("general_cost_10"))
	  be_general_cost_11 = cdbl(rs("general_cost_11"))
	  be_general_cost_12 = cdbl(rs("general_cost_12")) 
	end if
	rs.close()

	'sql = "select * from person_cost where cost_month ='"&cost_month&"' and emp_no ='"&user_id&"'"
	'sql = "call COST_PERSON_01('" & mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01" & "','"&user_id&"',@ret)" 

	'set rs=dbconn.execute(sql)
  arParams = Array(mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                   mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _
                   mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _                   
                   mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _                   
                   mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _                   
                   mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _                  
                   mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", _                   
                   mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01", user_id)
   Set cmd = server.CreateObject("ADODB.Command")
   cmd.CommandText = query_person_cost
'Response.write	query_person_cost&"<br><br>"

  for i = 0 to 8
     'Response.write	arParams(i)&"<br>"
  Next
  
   Set cmd.ActiveConnection = dbconn
   Set rs = cmd.execute(,arParams,1)

	' 차량 정보
	sql = "select * from car_info where owner_emp_no ='"&user_id&"'"
	'Response.write sql
	set rs_car=dbconn.execute(sql)
	if rs_car.eof then
		car_info = "차량없음"
		car_owner = ""
	else  	
		car_info = rs_car("car_owner") + "차량 , 차종 : " + rs_car("car_name") + " , 유종 : " + rs_car("oil_kind")
		car_owner = rs_car("car_owner")
	end if	

	general_cnt = rs("general_cnt")	 
	general_cost = rs("general_cost")	 
	overtime_cnt = rs("overtime_cnt")	 
	overtime_cost = rs("overtime_cost")	 
	gas_km = cdbl(rs("gas_km"))  
	gas_unit = cdbl(rs("gas_unit"))
	gas_cost = cdbl(rs("gas_cost"))  
	gasol_km = cdbl(rs("gasol_km"))  
	gasol_unit = cdbl(rs("gasol_unit"))  
	gasol_cost = cdbl(rs("gasol_cost"))  
	diesel_km = cdbl(rs("diesel_km")  ) 
	diesel_unit = cdbl(rs("diesel_unit"))  
	diesel_cost = cdbl(rs("diesel_cost"))  
	somopum_cost = cdbl(rs("somopum_cost")  )  
	fare_cnt = rs("fare_cnt")	 		 
	fare_cost = rs("fare_cost")	 		 
	oil_cash_cost = rs("oil_cash_cost")	 
	repair_cost = rs("repair_cost")	 
	parking_cost = rs("parking_cost")	 
	toll_cost = rs("toll_cost")	 
	card_cost = rs("card_cost")	 
	card_cost_vat = rs("card_cost_vat")	 
	juyoo_card_cost = rs("juyoo_card_cost")	 
	juyoo_card_cost_vat = rs("juyoo_card_cost_vat")	 
	return_cash = rs("return_cash")	 
	tot_km = gas_km + diesel_km + gasol_km
	tot_cost = gas_cost + diesel_cost + gasol_cost
	card_price = card_cost + card_cost_vat
	juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat
	'cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
  company_yn = cdbl(rs("company_yn"))
  
  if company_yn > 0 then
  	cash_tot_cost =   fare_cost + oil_cash_cost + toll_cost + parking_cost
  else
    cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
  end if
	variation_memo = rs("variation_memo")	

	general_cost_01 = cdbl(rs("general_cost_01"))
	general_cost_02 = cdbl(rs("general_cost_02"))
	general_cost_03 = cdbl(rs("general_cost_03"))
	general_cost_04 = cdbl(rs("general_cost_04"))
	general_cost_05 = cdbl(rs("general_cost_05"))
	general_cost_06 = cdbl(rs("general_cost_06"))
	general_cost_07 = cdbl(rs("general_cost_07"))
	general_cost_08 = cdbl(rs("general_cost_08"))
	general_cost_09 = cdbl(rs("general_cost_09"))
	general_cost_10 = cdbl(rs("general_cost_10"))
	general_cost_11 = cdbl(rs("general_cost_11"))
	general_cost_12 = cdbl(rs("general_cost_12"))

	title_line = "개인별 비용 정산 전표"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
		  		 window.close () ;
			}
			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = false; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 13; //오른쯕 여백 설정
                factory.printing.bottomMargin = 15; //바닦 여백 설정
        //		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
        //		factory.printing.printer = ""; //프린터 할 프린터 이름
        //		factory.printing.paperSize = "A4"; //용지선택
        //		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
        //		factory.printing.collate = true; //순서대로 출력하기
        //		factory.printing.copies = "1"; //인쇄할 매수
        //		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
        //		factory.printing.Printer(true); //출력하기
                factory.printing.Preview(); //윈도우를 통해서 출력
                factory.printing.Print(false); //윈도우를 통해서 출력
            }
        </script>
	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="wrap">			
			<div id="container">
				<h3 class="btit"><%=title_line%></h3>
				<form action="person_cost_report.asp" method="post" name="frm">
				<div class="gView">
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td height="50px" width="30%">&nbsp;</td>
				    <td height="50px" width="30%">&nbsp;</td>
				    <td height="50px" width="*"><table cellspacing="0" cellpadding="0" class="tablePrt">
				      <tr>
				        <td rowspan="2" style=" border-left:1px solid #000000;"><strong>결<br><br>재</strong></td>
				        <td class="center" width="23%"><strong>담 당</strong></td>
				        <td class="center" width="23%"><strong>팀 장</strong></td>
				        <td class="center" width="23%"><strong>사업부장</strong></td>
				        <td class="center" width="23%" style=" border-right:1px solid #000000;"><strong>본부장</strong></td>
			          </tr>
				      <tr>
				        <td height="60px" style=" border-left:1px solid #000000;">&nbsp;</td>
				        <td>&nbsp;</td>
				        <td>&nbsp;</td>
				        <td style=" border-right:1px solid #000000;">&nbsp;</td>
			          </tr>
				      </table>
                    </td>
			      </tr>
				  </table>
					<br>
                    <h3 class="stit">* 사원명 : <%=user_name%>&nbsp;<%=user_grade%>&nbsp;(<%=user_id%>),&nbsp;&nbsp;조직명 : <%=emp_company%>&nbsp;<%=bonbu%>&nbsp;<%=saupbu%>&nbsp;<%=team%>&nbsp;<%=reside_place%></h3>
                    <br>
                    <h3 class="stit">* 차  량 : <%=car_info%></h3>
                    <br>
					<table cellpadding="0" cellspacing="0" class="tablePrt">
						<colgroup>
							<col width="*" >
							<col width="4%" >
							<col width="6%" >
							<col width="4%" >
							<col width="6%" >
							<col width="4%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
						</colgroup>
						<thead>
							<tr bgcolor="#666666">
								<th rowspan="3" class="first" scope="col">년월</th>
								<th colspan="2" style=" border-bottom:1px solid #000000;" scope="col">야특근</th>
								<th colspan="11" style=" border-bottom:1px solid #000000;" scope="col">현금 사용</th>
								<th rowspan="3" scope="col">주유카드</th>
								<th rowspan="3" scope="col">법인카드</th>
								<th rowspan="3" scope="col">정산금액</th>
							</tr>
							<tr>
							  <th colspan="2" style=" border-bottom:1px solid #000000;border-left:1px solid #000000;" scope="col">신청금액</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #000000;">일반비용</th>
							  <th colspan="2" style=" border-bottom:1px solid #000000;" scope="col">대중교통비</th>
							  <th colspan="3" style=" border-bottom:1px solid #000000;" scope="col"><%=car_owner%> 차량 주행비용</th>
							  <th style=" border-bottom:1px solid #000000;" scope="col">회사차량</th>
							  <th colspan="2" style=" border-bottom:1px solid #000000;" scope="col">차량 유지비</th>
							  <th rowspan="2" scope="col"><p>현금사용</p><p>소계</p></th>
						  </tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #000000;">건수</th>
							  <th scope="col">금액</th>
							  <th scope="col">건수</th>
							  <th scope="col">금액</th>
							  <th scope="col">건수</th>
							  <th scope="col">금액</th>
							  <th scope="col">주행(KM)</th>
							  <th scope="col">유류비</th>
							  <th scope="col">소모품</th>
							  <th scope="col">주유비</th>
							  <th scope="col">주차비</th>
							  <th scope="col">통행료</th>
						  </tr>
						</thead>
						<tbody>
							<tr>
								<td height="25" class="first"><%=be_month_view%></td>
	    	      	<td class="right"><%=formatnumber(be_overtime_cnt,0)%></td>
								<td class="right"><%=formatnumber(be_overtime_cost,0)%></td>
								<td class="right"><%=formatnumber(be_general_cnt,0)%></td>
								<td class="right"><%=formatnumber(be_general_cost,0)%></td>
								<td class="right"><%=formatnumber(be_fare_cnt,0)%></td>
								<td class="right"><%=formatnumber(be_fare_cost,0)%></td>
								<td class="right"><%=formatnumber(be_tot_km,0)%></td>
								<td class="right"><%=formatnumber(be_tot_cost,0)%></td>
								<td class="right"><%=formatnumber(be_somopum_cost,0)%></td>
								<td class="right"><%=formatnumber(be_oil_cash_cost,0)%></td>
								<td class="right"><%=formatnumber(be_parking_cost,0)%></td>
								<td class="right"><%=formatnumber(be_toll_cost,0)%></td>
								<td class="right"><%=formatnumber(be_cash_tot_cost,0)%></td>
								<td class="right"><%=formatnumber(be_juyoo_card_price,0)%></td>
								<td class="right"><%=formatnumber(be_card_cost,0)%></td>
								<td class="right"><%=formatnumber(be_return_cash,0)%></td>
							</tr>
							<tr>
								<td class="first" height="25"><%=month_view%></td>
								<td class="right"><%=formatnumber(overtime_cnt,0)%></td>
								<td class="right"><%=formatnumber(overtime_cost,0)%></td>
								<td class="right"><%=formatnumber(general_cnt,0)%></td>
								<td class="right"><%=formatnumber(general_cost,0)%></td>
								<td class="right"><%=formatnumber(fare_cnt,0)%></td>
								<td class="right"><%=formatnumber(fare_cost,0)%></td>
								<td class="right"><%=formatnumber(tot_km,0)%></td>
								<td class="right"><%=formatnumber(tot_cost,0)%></td>
								<td class="right"><%=formatnumber(somopum_cost,0)%></td>
								<td class="right"><%=formatnumber(oil_cash_cost,0)%></td>
								<td class="right"><%=formatnumber(parking_cost,0)%></td>
								<td class="right"><%=formatnumber(toll_cost,0)%></td>
								<td class="right"><%=formatnumber(cash_tot_cost,0)%></td>
								<td class="right"><%=formatnumber(juyoo_card_price,0)%></td>
								<td class="right"><%=formatnumber(card_price,0)%></td>
								<td class="right"><%=formatnumber(return_cash,0)%></td>
							</tr>
					<%
					overtime_cal = overtime_cost - be_overtime_cost
					if be_overtime_cost = 0 then
						overtime_per = 100
					end if
					if overtime_cost = 0 then
						overtime_per = -100
					end if
					if overtime_cost = 0 and be_overtime_cost = 0 then
						overtime_per = 0
					end if
					if overtime_cost <> 0 and be_overtime_cost <> 0 then
						overtime_per = overtime_cal / be_overtime_cost * 100
					end if

					general_cal = general_cost - be_general_cost
					if be_general_cost = 0 then
						general_per = 100
					end if
					if general_cost = 0 then
						general_per = -100
					end if
					if general_cost = 0 and be_general_cost = 0 then
						general_per = 0
					end if
					if general_cost <> 0 and be_general_cost <> 0 then
						general_per = general_cal / be_general_cost * 100
					end if

					fare_cal = fare_cost - be_fare_cost
					if be_fare_cost = 0 then
						fare_per = 100
					end if
					if fare_cost = 0 then
						fare_per = -100
					end if
					if fare_cost = 0 and be_fare_cost = 0 then
						fare_per = 0
					end if
					if fare_cost <> 0 and be_fare_cost <> 0 then
						fare_per = fare_cal / be_fare_cost * 100
					end if

					tot_km_cal = tot_km - be_tot_km
					if be_tot_km = 0 then
						tot_km_per = 100
					end if
					if tot_km = 0 then
						tot_km_per = -100
					end if
					if tot_km = 0 and be_tot_km = 0 then
						tot_km_per = 0
					end if
					if tot_km <> 0 and be_tot_km <> 0 then
						tot_km_per = tot_km_cal / be_tot_km * 100
					end if

					tot_cost_cal = tot_cost - be_tot_cost
					if be_tot_cost = 0 then
						tot_cost_per = 100
					end if
					if tot_cost = 0 then
						tot_cost_per = -100
					end if
					if tot_cost = 0 and be_tot_cost = 0 then
						tot_cost_per = 0
					end if
					if tot_cost <> 0 and be_tot_cost <> 0 then
						tot_cost_per = tot_cost_cal / be_tot_cost * 100
					end if

					somopum_cost_cal = somopum_cost - be_somopum_cost
					if be_somopum_cost = 0 then
						somopum_per = 100
					end if
					if somopum_cost = 0 then
						somopum_per = -100
					end if
					if somopum_cost = 0 and be_somopum_cost = 0 then
						somopum_per = 0
					end if
					if somopum_cost <> 0 and be_somopum_cost <> 0 then
						somopum_per = somopum_cost_cal / be_somopum_cost * 100
					end if

					oil_cash_cost_cal = oil_cash_cost - be_oil_cash_cost
					if be_oil_cash_cost = 0 then
						oil_cash_per = 100
					end if
					if oil_cash_cost = 0 then
						oil_cash_per = -100
					end if
					if oil_cash_cost = 0 and be_oil_cash_cost = 0 then
						oil_cash_per = 0
					end if
					if oil_cash_cost <> 0 and be_oil_cash_cost <> 0 then
						oil_cash_per = oil_cash_cost_cal / be_oil_cash_cost * 100
					end if

					parking_cost_cal = parking_cost - be_parking_cost
					if be_parking_cost = 0 then
						parking_per = 100
					end if
					if parking_cost = 0 then
						parking_per = -100
					end if
					if parking_cost = 0 and be_parking_cost = 0 then
						parking_per = 0
					end if
					if parking_cost <> 0 and be_parking_cost <> 0 then
						parking_per = parking_cost_cal / be_parking_cost * 100
					end if

					cash_tot_cost_cal = cash_tot_cost - be_cash_tot_cost
					if be_cash_tot_cost = 0 then
						cash_tot_per = 100
					end if
					if cash_tot_cost = 0 then
						cash_tot_per = -100
					end if
					if cash_tot_cost = 0 and be_cash_tot_cost = 0 then
						cash_tot_per = 0
					end if
					if cash_tot_cost <> 0 and be_cash_tot_cost <> 0 then
						cash_tot_per = cash_tot_cost_cal / be_cash_tot_cost * 100
					end if

					juyoo_card_price_cal = juyoo_card_price - be_juyoo_card_price
					if be_juyoo_card_price = 0 then
						juyoo_card_per = 100
					end if
					if juyoo_card_price = 0 then
						juyoo_card_per = -100
					end if
					if juyoo_card_price = 0 and be_juyoo_card_price = 0 then
						juyoo_card_per = 0
					end if
					if juyoo_card_price <> 0 and be_juyoo_card_price <> 0 then
						juyoo_card_per = juyoo_card_price_cal / be_juyoo_card_price * 100
					end if

					card_cost_cal = card_cost - be_card_cost
					if be_card_cost = 0 then
						card_per = 100
					end if
					if card_cost = 0 then
						card_per = -100
					end if
					if card_cost = 0 and be_card_cost = 0 then
						card_per = 0
					end if
					if card_cost <> 0 and be_card_cost <> 0 then
						card_per = card_cost_cal / be_card_cost * 100
					end if

					return_cash_cal = return_cash - be_return_cash
					if be_return_cash = 0 then
						return_per = 100
					end if
					if return_cash = 0 then
						return_per = -100
					end if
					if return_cash = 0 and be_return_cash = 0 then
						return_per = 0
					end if
					if return_cash <> 0 and be_return_cash <> 0 then
						return_per = return_cash_cal / be_return_cash * 100
					end if

					%>
							<tr>
								<td height="25" class="first">증감(%)</td>
				      		  	<td class="right">&nbsp;</td>
								<td class="right"><%=formatnumber(overtime_per,2)%>%</td>
								<td class="right">&nbsp;</td>
								<td class="right"><%=formatnumber(general_per,2)%>%</td>
								<td class="right">&nbsp;</td>
								<td class="right"><%=formatnumber(fare_per,2)%>%</td>
								<td class="right"><%=formatnumber(tot_km_per,2)%>%</td>
								<td class="right"><%=formatnumber(tot_cost_per,2)%>%</td>
								<td class="right"><%=formatnumber(somopum_per,2)%>%</td>
								<td class="right"><%=formatnumber(oil_cash_per,2)%>%</td>
								<td class="right"><%=formatnumber(parking_per,2)%>%</td>
								<td class="right"><%=formatnumber(toll_per,2)%>%</td>
								<td class="right"><%=formatnumber(cash_tot_per,2)%>%</td>
								<td class="right"><%=formatnumber(juyoo_card_per,2)%>%</td>
								<td class="right"><%=formatnumber(card_per,2)%>%</td>
								<td class="right"><%=formatnumber(return_per,2)%>%</td>
						  </tr>
						   <tr>
								<td height="25" class="first"><strong>증감사유</strong></td>
				      		  	<td colspan="16" class="left"><%=variation_memo%>&nbsp;</td>
						  </tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    </td>
			      </tr>
				  </table>
				<br>
				<h3 class="stit">1) 전표 출력후 뒷면에 비용 영수증을 붙여서 해당 결재란에 결재를 득한 후 경리부로 발송하시면 됩니다.</h3>
                <br>
                <h3 class="stit">2) 만약 본인의 정산금액이 상이하면 근거 자료를 경리부로 발송해 주시길 바랍니다.</h3> 
			
        <table width="100%" border="0" cellpadding="0" cellspacing="0"  class="tablePrt" style=" margin-top: 20px; margin-bottom: 5px; ">
          <colgroup>
              <col width="*" >
              <col width="9%" >
              <col width="9%" >
              <col width="9%" >
              <col width="9%" >
              <col width="9%" >
              <col width="8%" >
              <col width="8%" >
              <col width="9%" >
              <col width="8%" >
              <col width="9%" >
            </colgroup>
            <thead>
              <tr bgcolor="#666666">
                <th rowspan="2" class="first" scope="col">년월</th>
                <th colspan="10" style=" border-bottom:1px solid #000000; " scope="col">일반경비 계정별 분류</th>
              </tr>
              <tr>
                <th style=" border-left:1px solid #000000;" scope="col">차량유지비</th>
                <th scope="col">여비교통비</th>
                <th scope="col">복리후생비</th>                
                <th scope="col">접대비</th>
                <th scope="col">회의비</th>
                <th scope="col">사무용품비</th>
                <th scope="col">소모품비</th>
                <th scope="col">국내출장비</th>
                <th scope="col">수선비</th>
                <th scope="col">지급수수료</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td height="25" class="first"><%=be_month_view%></td>
                <td class="right"><%=formatnumber(be_general_cost_01,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_02,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_03,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_04,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_05,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_06,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_07,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_10,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_11,0)%></td>
                <td class="right"><%=formatnumber(be_general_cost_12,0)%></td>
              </tr>
              <tr>
                <td class="first" height="25"><%=month_view%></td>
                <td class="right"><%=formatnumber(general_cost_01,0)%></td>
                <td class="right"><%=formatnumber(general_cost_02,0)%></td>
                <td class="right"><%=formatnumber(general_cost_03,0)%></td>
                <td class="right"><%=formatnumber(general_cost_04,0)%></td>
                <td class="right"><%=formatnumber(general_cost_05,0)%></td>
                <td class="right"><%=formatnumber(general_cost_06,0)%></td>
                <td class="right"><%=formatnumber(general_cost_07,0)%></td>
                <td class="right"><%=formatnumber(general_cost_10,0)%></td>
                <td class="right"><%=formatnumber(general_cost_11,0)%></td>
                <td class="right"><%=formatnumber(general_cost_12,0)%></td> 
              </tr>
          </table>			
			
			
			
			</form>
				<br>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				<br>
		</div>				
	</div>        				
	</body>
</html>

