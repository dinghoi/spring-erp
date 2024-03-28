<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
'on Error resume next

Dim from_date
Dim to_date
Dim win_sw
	 
cost_month=Request.form("cost_month")
reside_company=Request.form("reside_company")
if cost_month = "" then
	cost_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
end If
from_date = mid(cost_month,1,4) + "-" + mid(cost_month,6,2) + "-31"

if user_id = "100031" or user_id = "100178" then
	reside_company = "한진"
	reside_company_view = "한진"
end if
if user_id = "100029" then
	reside_company = "한화생명"
	reside_company_view = "한화생명"
end if

sql = "select * from emp_master_month where (emp_month = '"&cost_month&"') and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&from_date&"') and (emp_reside_company like '%"&reside_company&"%') order by emp_reside_place, emp_name"
rs_emp.Open sql, Dbconn, 1
	
title_line = "상주처별 개인별 비용사용 및 정산현황"
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
			function getPageCode(){
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("발생년월을 입력하세요.");
					return false;
				}	
				return true;
			}

			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<h3 class="stit">바로 옆 메뉴에 있는 비용마감처리후에 정확한 금액을 조회할 수 있습니다. 만약 마감 후 금액 조정이 필요하시면 마감을 취소하고 수정후 재 마감을 하셔야 정확한 금액을 조회할수 있습니다. </h3>
				<form action="reside_person_cal_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>발생년월&nbsp;</strong>(예201401) : 
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>
                                <label>
								<strong>상주처 &nbsp;:</strong>
						<% 
							if cost_grade = "0" then	
								sql_org="select org_cost_group from emp_org_mst where (org_level = '상주처') group by org_cost_group order by org_cost_group asc"
	                            rs_org.Open sql_org, Dbconn, 1
                        %>
                                <select name="reside_company" id="reside_company" style="width:150px">
                                    <option value="" <%If reside_company = "" then %>selected<% end if %>>선택</option>
						<% 
								do until rs_org.eof
						%>
          							<option value='<%=rs_org("org_cost_group")%>' <%If rs_org("org_cost_group") = reside_company  then %>selected<% end if %>><%=rs_org("org_cost_group")%></option>
          				<%
									rs_org.movenext()  
								loop 
								rs_org.Close()
						%>
                                </select>
						<%	  else	%>
								<%=reside_company_view%>
						<% 
							end if 
						%>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
					<table cellpadding="0" cellspacing="0">
					<tr>
                    	<td>
      					<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="7%" >
							<col width="3%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="3" class="first" scope="col">팀</th>
								<th rowspan="3" scope="col">사용자</th>
								<th rowspan="3" scope="col">차량</th>
								<th style=" border-bottom:1px solid #e3e3e3;" scope="col">야특근</th>
								<th colspan="9" style=" border-bottom:1px solid #e3e3e3;" scope="col">현금 사용</th>
								<th rowspan="3" scope="col">주유카드</th>
								<th rowspan="3" scope="col">정산금액</th>
								<th rowspan="3" scope="col"><p>현금</p><p>수리비</p></th>
								<th rowspan="3" scope="col">법인카드<br>(VAT별도)</th>
								<th rowspan="3" scope="col">사용 계</th>
							</tr>
							<tr>
							  <th style=" border-bottom:1px solid #e3e3e3;border-left:1px solid #e3e3e3;" scope="col">신청금액</th>
							  <th scope="col" style=" border-bottom:1px solid #e3e3e3;">일반비용</th>
							  <th style=" border-bottom:1px solid #e3e3e3;" scope="col">대중교통</th>
							  <th colspan="3" style=" border-bottom:1px solid #e3e3e3;" scope="col">개인 차량 주행비용</th>
							  <th style=" border-bottom:1px solid #e3e3e3;" scope="col">회사차량</th>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">차량 유지비</th>
							  <th rowspan="2" scope="col"><p>현금사용</p><p>소계</p></th>
						  </tr>
							<tr>
							  <th style="border-left:1px solid #e3e3e3;" scope="col">금액</th>
							  <th scope="col">금액</th>
							  <th scope="col">금액</th>
							  <th scope="col">주행(KM)</th>
							  <th scope="col">유류비</th>
							  <th scope="col">소모품</th>
							  <th scope="col">주유비</th>
							  <th scope="col">주차비</th>
							  <th scope="col">통행료</th>
						  </tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="*" >
							<col width="7%" >
							<col width="3%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<tbody>
						<%
						emp_cnt = 0
						sum_emp = 0
						tot_emp = 0
						sum_general_cnt = 0 
						sum_general_cost = 0 
						sum_overtime_cnt = 0	 
						sum_overtime_cost = 0
						sum_fare_cost = 0	 
						sum_tot_km = 0
						sum_tot_cost = 0
						sum_somopum_cost = 0
						sum_oil_cash_cost = 0
						sum_parking_cost = 0
						sum_toll_cost = 0
						sum_juyoo_card_price = 0
						sum_card_cost = 0
						sum_cash_tot_cost = 0
						sum_return_cash = 0
						sum_repair_cost = 0
						sum_cost_sum = 0

						tot_general_cnt = 0 
						tot_general_cost = 0 
						tot_overtime_cnt = 0	 
						tot_overtime_cost = 0
						tot_fare_cost = 0	 
						tot_tot_km = 0
						tot_tot_cost = 0
						tot_somopum_cost = 0
						tot_oil_cash_cost = 0
						tot_parking_cost = 0
						tot_toll_cost = 0
						tot_juyoo_card_price = 0
						tot_card_cost = 0
						tot_cash_tot_cost = 0
						tot_return_cash = 0
						tot_cost_sum = 0
						if rs_emp.eof or rs_emp.bof then
							bi_team = ""
					      else						  
							if isnull(rs_emp("emp_reside_place")) or rs_emp("emp_reside_place") = "" then	
								bi_team = ""
							  else
								bi_team = rs_emp("emp_reside_place")
							end if
						end if
							
						do until rs_emp.eof
							if isnull(rs_emp("emp_reside_place")) or rs_emp("emp_reside_place") = "" then
								emp_reside_place = ""
							  else
							  	emp_reside_place = rs_emp("emp_reside_place")
							end if
							
							if bi_team <> emp_reside_place then
								sum_emp = emp_cnt
								tot_emp = tot_emp + sum_emp
								emp_cnt = 0
						%>
							<tr>
								<td bgcolor="#EEFFFF" class="first">소계</td>
								<td bgcolor="#EEFFFF"><%=sum_emp%>명</td>
								<td bgcolor="#EEFFFF">&nbsp;</td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_overtime_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_general_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_fare_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_tot_km,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_tot_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_somopum_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_oil_cash_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_parking_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_toll_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_cash_tot_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_juyoo_card_price,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_return_cash,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_repair_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_card_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_cost_sum,0)%></td>
							</tr>
                        <%
								sum_general_cnt = 0 
								sum_general_cost = 0 
								sum_overtime_cnt = 0	 
								sum_overtime_cost = 0
								sum_fare_cost = 0	 
								sum_tot_km = 0
								sum_tot_cost = 0
								sum_somopum_cost = 0
								sum_oil_cash_cost = 0
								sum_parking_cost = 0
								sum_toll_cost = 0
								sum_juyoo_card_price = 0
								sum_card_cost = 0
								sum_cash_tot_cost = 0
								sum_return_cash = 0
								sum_repair_cost = 0
								sum_cost_sum = 0
								bi_team = emp_reside_place
							end if														

							sql = "select * from person_cost where cost_month = '"&cost_month&"' and emp_no = '"&rs_emp("emp_no")&"'"
							set rs=dbconn.execute(sql)
							if rs.eof or rs.bof then
								general_cnt = 0 
								general_cost = 0 
								overtime_cnt = 0	 
								overtime_cost = 0
								gas_km = 0
								gas_unit = 0
								gas_cost = 0
								gasol_km = 0
								gasol_unit = 0 
								gasol_cost = 0	 
								diesel_km = 0
								diesel_unit = 0
								diesel_cost = 0
								somopum_cost = 0
								fare_cost = 0	 
								oil_cash_cost = 0
								repair_cost = 0
								parking_cost = 0
								toll_cost = 0
								juyoo_card_cost = 0
								juyoo_card_cost_vat = 0
								card_cost = 0
								card_cost_vat = 0
								return_cash = 0
								repair_cost = 0
								tot_km = gas_km + diesel_km + gasol_km
								tot_cost = gas_cost + diesel_cost + gasol_cost
								juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat
								card_price = card_cost + card_cost_vat
								cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
								cost_sum = 0
								car_owner = "."
							  else
								emp_end = rs("emp_end")
								car_owner = rs("car_owner")
								if car_owner = "없음" then
									car_owner = "."
								end if
								general_cnt = rs("general_cnt")	 
								general_cost = rs("general_cost")	 
								overtime_cnt = rs("overtime_cnt")	 
								overtime_cost = rs("overtime_cost")	 
								gas_km = rs("gas_km")	 
								gas_unit = rs("gas_unit")	 
								gas_cost = rs("gas_cost")	 
								gasol_km = rs("gasol_km")	 
								gasol_unit = rs("gasol_unit")	 
								gasol_cost = rs("gasol_cost")	 
								diesel_km = rs("diesel_km")	 
								diesel_unit = rs("diesel_unit")	 
								diesel_cost = rs("diesel_cost")	 
								somopum_cost = rs("somopum_cost")	 
								fare_cost = rs("fare_cost")	 		 
								oil_cash_cost = rs("oil_cash_cost")	 
								repair_cost = rs("repair_cost")	 
								parking_cost = rs("parking_cost")	 
								toll_cost = rs("toll_cost")	 
								juyoo_card_cost = rs("juyoo_card_cost")	 
								juyoo_card_cost_vat = rs("juyoo_card_cost_vat")	 
								juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat
								card_cost = rs("card_cost")	 
								card_cost_vat = rs("card_cost_vat")	 
								return_cash = rs("return_cash")	 
								repair_cost = rs("repair_cost")	 
								tot_km = gas_km + diesel_km + gasol_km
								tot_cost = gas_cost + diesel_cost + gasol_cost
								cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
								if rs("car_owner") = "개인" then
									cost_sum = cash_tot_cost + card_cost + overtime_cost
								  else
								  	cost_sum = cash_tot_cost + repair_cost + card_cost + overtime_cost
								end if
							end if

							sum_general_cnt = sum_general_cnt + general_cnt 
							sum_general_cost = sum_general_cost + general_cost 
							sum_overtime_cnt = sum_overtime_cnt + overtime_cnt	 
							sum_overtime_cost = sum_overtime_cost + overtime_cost
							sum_fare_cost = sum_fare_cost + fare_cost	 
							sum_tot_km = sum_tot_km + tot_km
							sum_tot_cost = sum_tot_cost + tot_cost
							sum_somopum_cost = sum_somopum_cost + somopum_cost
							sum_oil_cash_cost = sum_oil_cash_cost + oil_cash_cost
							sum_parking_cost = sum_parking_cost + parking_cost
							sum_toll_cost = sum_toll_cost + toll_cost
							sum_juyoo_card_price = sum_juyoo_card_price + juyoo_card_price
							sum_card_cost = sum_card_cost + card_cost
							sum_cash_tot_cost = sum_cash_tot_cost + cash_tot_cost
							sum_return_cash = sum_return_cash + return_cash
							sum_repair_cost = sum_repair_cost + repair_cost
							sum_cost_sum = sum_cost_sum + cost_sum

							tot_general_cnt = tot_general_cnt + general_cnt 
							tot_general_cost = tot_general_cost + general_cost 
							tot_overtime_cnt = tot_overtime_cnt + overtime_cnt	 
							tot_overtime_cost = tot_overtime_cost + overtime_cost
							tot_fare_cost = tot_fare_cost + fare_cost	 
							tot_tot_km = tot_tot_km + tot_km
							tot_tot_cost = tot_tot_cost + tot_cost
							tot_somopum_cost = tot_somopum_cost + somopum_cost
							tot_oil_cash_cost = tot_oil_cash_cost + oil_cash_cost
							tot_parking_cost = tot_parking_cost + parking_cost
							tot_toll_cost = tot_toll_cost + toll_cost
							tot_juyoo_card_price = tot_juyoo_card_price + juyoo_card_price
							tot_card_cost = tot_card_cost + card_cost
							tot_cash_tot_cost = tot_cash_tot_cost + cash_tot_cost
							tot_return_cash = tot_return_cash + return_cash
							tot_repair_cost = tot_repair_cost + repair_cost
							tot_cost_sum = tot_cost_sum + cost_sum
'							if cost_sum <> 0 then
							
							if emp_end = "근무" then
								emp_view = rs_emp("emp_name") + " " + rs_emp("emp_job")
							  else
							  	emp_view = "퇴사 " + rs_emp("emp_name")
							end if
							if emp_end = "근무" or ( emp_end = "퇴사" and juyoo_card_price > 0 ) then
						%>
							<tr>
								<td class="first"><%=rs_emp("emp_reside_place")%>&nbsp;</td>
						<% if emp_end = "퇴사" then	%>
								<td bgcolor="#FFCCFF"><%=emp_view%></td>
                        <%   else	%>
								<td><%=emp_view%></td>
                        <% end if	%>
								<td><%=car_owner%></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="야특근"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(overtime_cost,0)%></a></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="일반경비"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(general_cost,0)%></a></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="대중교통"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(fare_cost,0)%></a></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="주행거리"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(tot_km,0)%></a></td>
								<td class="right"><%=formatnumber(tot_cost,0)%></td>
								<td class="right"><%=formatnumber(somopum_cost,0)%></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="주유비"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(oil_cash_cost,0)%></a></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="주차료"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(parking_cost,0)%></a></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="통행료"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(toll_cost,0)%></a></td>
								<td class="right"><%=formatnumber(cash_tot_cost,0)%></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="주유카드"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(juyoo_card_price,0)%></a></td>
								<td class="right"><%=formatnumber(return_cash,0)%></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="차량수리비"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(repair_cost,0)%></a></td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_detail_view.asp?cost_yymm=<%=cost_month%>&cost_id=<%="법인카드"%>&emp_no=<%=rs_emp("emp_no")%>','person_cost_detail_view_pop','scrollbars=yes,width=900,height=500')"><%=formatnumber(card_cost,0)%></a></td>
								<td class="right"><%=formatnumber(cost_sum,0)%></td>
							</tr>
						<%
							end if
							emp_cnt = emp_cnt + 1
							rs_emp.movenext()
						loop
						if tot_cost_sum > 0 then
							sum_emp = emp_cnt
							tot_emp = tot_emp + sum_emp
						%>
							<tr>
								<td bgcolor="#EEFFFF" class="first">소계</td>
								<td bgcolor="#EEFFFF"><%=sum_emp%>명</td>
								<td bgcolor="#EEFFFF">&nbsp;</td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_overtime_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_general_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_fare_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_tot_km,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_tot_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_somopum_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_oil_cash_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_parking_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_toll_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_cash_tot_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_juyoo_card_price,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_return_cash,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_repair_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_card_cost,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_cost_sum,0)%></td>
							</tr>
						<% end if	%>
						<% if saupbu <> "본인" then	%> 
							<tr>
								<td bgcolor="#FFE8E8" class="first">총계</td>
								<td bgcolor="#FFE8E8"><%=tot_emp%>명</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_overtime_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_general_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_fare_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_tot_km,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_tot_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_somopum_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_oil_cash_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_parking_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_toll_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_cash_tot_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_juyoo_card_price,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_return_cash,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_repair_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_card_cost,0)%></td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(tot_cost_sum,0)%></td>
							</tr>
						<% end if	%>
						</tbody>
						</table>
                        </DIV>
						</td>
                    </tr>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="reside_person_cal_excel.asp?cost_month=<%=cost_month%>&reside_company=<%=reside_company%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				  </table>
			</form>
				<br>
		</div>				
	</div>        				
	</body>
</html>

