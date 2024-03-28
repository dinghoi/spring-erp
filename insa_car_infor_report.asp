<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	from_date=request("from_date")
    to_date=request("to_date")
end if

if view_condi = "" then
	view_condi = ""
	ck_sw = "n"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
end if

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_drv = Server.CreateObject("ADODB.Recordset")
Set Rs_insu = Server.CreateObject("ADODB.Recordset")
Set Rs_pen = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


sql = "select * from car_info where car_no = '"&view_condi&"'"
Rs.Open Sql, Dbconn, 1


title_line = "차량 토탈 정보 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "7 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("차량번호를 입력하세요.");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_infor_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>차량번호 : </strong>
                              <%
								Sql="select * from car_info where (end_date = '1900-01-01' or isNull(end_date)) ORDER BY car_no ASC"
	                            rs_car.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
                                    <option value="선택" <%If view_condi = "" then %>selected<% end if %>>선택</option>
                			  <% 
								do until rs_car.eof 
			  				  %>
                					<option value='<%=rs_car("car_no")%>' <%If view_condi = rs_car("car_no") then %>selected<% end if %>><%=rs_car("car_no")%></option>
                			  <%
									rs_car.movenext()  
								loop 
								rs_car.Close()
							  %>
            					</select>
                                </label>
                               	<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색">&nbsp;&nbsp;※ 반드시 기간조회 날짜를 입력하시고 검색을 클릭하십시요!</a>
                            </p>
						</dd>
					</dl>
				</fieldset>	
                <h3 class="stit">전체 현황은 엑셀다운로드 하십시요!, 차량번호별은 출력으로 하시면 됩니다</h3>			
                <div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
                            <col width="10%" >
							<col width="6%" >
							<col width="12%" >
							<col width="8%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
							<col width="12%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
                            <tr>
                                <th class="first" scope="col">차량번호</th>
								<th scope="col">차종</th>
                                <th scope="col">연식</th>
								<th scope="col">유종</th>
								<th scope="col">소유</th>
								<th scope="col">구매<br>구분</th>
								<th scope="col">차량등록일</th>
                                <th scope="col">보험료</th>
                                <th scope="col">보험기간</th>
								<th scope="col">운행자</th>
								<th scope="col">최종KM</th>
								<th scope="col">최종검사일</th>
						    </tr>
						</thead>
                        <tbody>
                            
						<%
						do until rs.eof
						
						     owner_emp_name = ""
							 owner_emp_no = rs("owner_emp_no")
						     if rs("owner_emp_name") = "" or isnull(rs("owner_emp_name")) then
							     Sql="select * from emp_master where emp_no = '"&owner_emp_no&"'"
	                             Set rs_emp=DbConn.Execute(Sql)
								 if not rs_emp.EOF or not rs_emp.BOF then
								       owner_emp_name = rs_emp("emp_name")
									else
									   owner_emp_name = ""
							     end if
							   else 
							     owner_emp_name = rs("owner_emp_name")
							 end if
							if rs("last_check_date") = "1900-01-01"  then
	                               last_check_date = ""
							   else 
							       last_check_date = rs("last_check_date")
	                        end if
	                        if rs("end_date") = "1900-01-01" then
	                               end_date = ""
							   else 
							       end_date = rs("end_date")
	                        end if
							if rs("car_year") = "1900-01-01" then
	                               car_year = ""
							   else 
							       car_year = rs("car_year")
	                        end if
						
	           			%>
							<tr>
                                <td class="first"><%=rs("car_no")%>&nbsp;</td>
								<td class="left"><%=rs("car_name")%></td>
                                <td class="left"><%=car_year%>&nbsp;</td>
								<td><%=rs("oil_kind")%></td>
								<td><%=rs("car_owner")%></td>
								<td><%=rs("buy_gubun")%>&nbsp;<%=rs("rental_company")%></td>
								<td><%=rs("car_reg_date")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("insurance_amt"),0)%>&nbsp;</td>
                                <td><%=rs("insurance_date")%>&nbsp;</td>
                                <td><%=owner_emp_name%>(<%=rs("owner_emp_no")%>)&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("last_km"),0)%>&nbsp;</td>
								<td><%=last_check_date%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>  
						</tbody>
                </table>      
         <%
						sql = "select * from car_insurance where ins_car_no = '"&view_condi&"' ORDER BY ins_car_no,ins_date ASC"
                        Rs_insu.Open Sql, Dbconn, 1
						if not Rs_insu.EOF or not Rs_insu.BOF then
		 %>
                <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="7%" >
                              <col width="6%" >
							  <col width="10%" >
                              <col width="6%" >
                              <col width="7%" >
                              <col width="7%" >
                              <col width="6%" >
                              <col width="6%" >
                              <col width="6%" >
                              <col width="6%" >
                              <col width="7%" >
                              <col width="6%" >
                              <col width="4%" >
                              <col width="*" >
                           </colgroup>
                           <thead>
                              <tr>
                                <th class="first" scope="col">차량번호</th>
                                <th scope="col">가입일</th>
                                <th scope="col">보험사</th>
                                <th scope="col">보험기간</th>
                                <th scope="col">보험료</th>
                                <th scope="col">대인1</th>
                                <th scope="col">대인2</th>
                                <th scope="col">대물</th>
                                <th scope="col">자기보험</th>
                                <th scope="col">무상해</th>
                                <th scope="col">자차</th>
                                <th scope="col">연령</th>
                                <th scope="col">긴급<br>출동</th>
                                <th scope="col">계약내용</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until Rs_insu.eof
                              car_no = Rs_insu("ins_car_no")
							  
							  Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
                              Set rs_car = DbConn.Execute(SQL)
							  if not rs_car.eof then
									car_name = rs_car("car_name")
									car_year = rs_car("car_year")
									car_reg_date = rs_car("car_reg_date")
	                             else
								    car_name = ""
									car_year = ""
									car_reg_date = ""
                              end if
                              rs_car.close()
	           			%>
							<tr>
                                <td><%=Rs_insu("ins_car_no")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_date")%>&nbsp;</td>
								<td><%=Rs_insu("ins_company")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_last_date")%>&nbsp;</td>
                                <td><%=formatnumber(Rs_insu("ins_amount"),0)%>&nbsp;</td>
                                <td><%=Rs_insu("ins_man1")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_man2")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_object")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_self")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_injury")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_self_car")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_age")%>&nbsp;</td>
                                <td><%=Rs_insu("ins_scramble")%>&nbsp;</td>
                         <% if Rs_insu("ins_contract_yn") = "Y" then %>
                                <td class="left">계약내용포함&nbsp;</td>
                         <%    else %>
                                <td class="left">계약내용미포함(<%=Rs_insu("ins_comment")%>)&nbsp;</td>
                         <% end if %>
							</tr>
						<%
							Rs_insu.movenext()
						loop
						%>
						</tbody>
                </table>   
         <% 
		                Rs_insu.close()
			  end if %>
         <%
						tot_fare = 0
                        tot_oil_price = 0
						tot_parking = 0
                        tot_toll = 0
                        sql = "select * from transit_cost where car_no = '"&view_condi&"' and run_date >= '"+from_date+"' and run_date <= '"+to_date+"' ORDER BY car_no,run_date,run_seq ASC"
						Rs.Open Sql, Dbconn, 1
                        do until rs.eof
                              tot_fare = tot_fare + int(rs("fare"))
	                          tot_oil_price = tot_oil_price + int(rs("oil_price"))
							  tot_parking = tot_parking + int(rs("parking"))
							  tot_toll = tot_toll + int(rs("toll"))
	                       rs.movenext()
                        loop
                        rs.close()	
						
						sql = "select * from transit_cost where car_no = '"&view_condi&"' and run_date >= '"+from_date+"' and run_date <= '"+to_date+"' ORDER BY car_no,run_date,run_seq ASC"
                        Rs_drv.Open Sql, Dbconn, 1
						if not Rs_drv.EOF or not Rs_drv.BOF then
		 %>                
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
						do until Rs_drv.eof
						    emp_no = Rs_drv("mg_ce_id")
							Sql = "select * from emp_master where emp_no = '"+emp_no+"'"
	                        Set Rs_emp = DbConn.Execute(SQL)
	                        if not Rs_emp.EOF or not Rs_emp.BOF then
			                       drv_owner_emp_name = rs_emp("emp_name")
                               else
                                   drv_owner_emp_name = emp_no
							end if
							
							if Rs_drv("start_km") = "" or isnull(Rs_drv("start_km")) then
								start_view = 0
							  else
							  	start_view = Rs_drv("start_km")
							end if
							if Rs_drv("end_km") = "" or isnull(Rs_drv("end_km")) then
								end_view = 0
							  else
							  	end_view = Rs_drv("end_km")
							end if
							run_km = Rs_drv("far")

	           			%>
							<tr>
                                <td class="first"><%=Rs_drv("car_no")%></td>
                                <td><%=Rs_drv("run_date")%></td>
								<td><%=drv_owner_emp_name%></td>
								<td><%=Rs_drv("car_owner")%></td>
								<td>
								<% if Rs_drv("car_owner") = "대중교통" then %>
								       <%=Rs_drv("transit")%>
								<%   else	%>                                
								       <%=Rs_drv("oil_kind")%>
								<% end if %>
                                </td>
								<td><%=Rs_drv("start_company")%>&nbsp;</td>
								<td class="left"><%=Rs_drv("start_point")%></td>
								<td class="right"><%=formatnumber(start_view,0)%></td>
								<td><%=Rs_drv("end_company")%>&nbsp;</td>
								<td class="left"><%=Rs_drv("end_point")%></td>
								<td class="right"><%=formatnumber(end_view,0)%></td>
								<td ><%=Rs_drv("run_memo")%></td>
								<td class="right"><%=formatnumber(Rs_drv("fare"),0)%></td>
								<td class="right"><%=formatnumber(Rs_drv("oil_price"),0)%></td>
								<td class="right"><%=formatnumber(Rs_drv("parking"),0)%></td>
								<td class="right"><%=formatnumber(Rs_drv("toll"),0)%></td>
							</tr>
						<%
							Rs_drv.movenext()
						loop
						%>
                            <tr>
								<td colspan="12" class="first" style="background:#ffe8e8;">총계</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_fare,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_oil_price,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_parking,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_toll,0)%></td>
							</tr>                        
						</tbody>
                   </table>  
         <% 
		                Rs_drv.close()
			  end if %>

         <%
						tot_amount = 0
                        sql = "select * from car_as where as_car_no = '"&view_condi&"' and as_date >= '"+from_date+"' and as_date <= '"+to_date+"' ORDER BY as_car_no,as_date,as_seq ASC"
						Rs.Open Sql, Dbconn, 1
                        do until rs.eof
                              tot_amount = tot_amount + int(rs("as_amount"))
	                       rs.movenext()
                        loop
                        rs.close()	
						
						sql = "select * from car_as where as_car_no = '"&view_condi&"' and as_date >= '"+from_date+"' and as_date <= '"+to_date+"' ORDER BY as_car_no,as_date,as_seq ASC"
                        Rs_as.Open Sql, Dbconn, 1
						if not Rs_as.EOF or not Rs_as.BOF then
		 %>                
                <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="8%" >
                              <col width="10%" >
							  <col width="12%" >
                              <col width="8%" >
							  <col width="15%" >
							  <col width="*" >
							  <col width="8%" >
                              <col width="6%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <th class="first" scope="col">차량번호</th>
                                <th scope="col">차종</th>
								<th scope="col">운행자</th>
                                <th scope="col">AS일자</th>
								<th scope="col">AS증상</th>
								<th scope="col">수리내용</th>
								<th scope="col">수리비용</th>
                                <th scope="col">결재구분</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until Rs_as.eof
	           			%>
							<tr>
                                <td class="first"><%=Rs_as("as_car_no")%></td>
                                <td><%=Rs_as("as_car_name")%></td>
                                <td><%=Rs_as("as_owner_emp_name")%>(<%=Rs_as("as_owner_emp_no")%>)</td>
                                <td><%=Rs_as("as_date")%></td>
								<td class="left"><%=Rs_as("as_cause")%></td>
								<td class="left"><%=Rs_as("as_solution")%></td>
                                <td class="right"><%=formatnumber(Rs_as("as_amount"),0)%></td>
                                <td><%=Rs_as("as_amount_sign")%></td>
							</tr>
						<%
							Rs_as.movenext()
						loop
						%>
                            <tr>
								<td colspan="6" class="first" style="background:#ffe8e8;">총계</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_amount,0)%>&nbsp;</td>
                                <td style="background:#ffe8e8;">&nbsp;</td>
							</tr>
						</tbody>
                   </table>  
         <% 
		                Rs_as.close()
			  end if %>
              
         <%
						tot_amount = 0
                        tot_in_amt = 0
                        sql = "select * from car_penalty where pe_car_no = '"&view_condi&"' and pe_date >= '"+from_date+"' and pe_date <= '"+to_date+"' ORDER BY pe_car_no,pe_date,pe_seq ASC"
						Rs.Open Sql, Dbconn, 1
                        do until rs.eof
                              tot_amount = tot_amount + int(rs("pe_amount"))
	                          tot_in_amt = tot_in_amt + int(rs("pe_in_amt"))
	                       rs.movenext()
                        loop
                        rs.close()	
						jan_amount = tot_amount - tot_in_amt
						
						sql = "select * from car_penalty where pe_car_no = '"&view_condi&"' and pe_date >= '"+from_date+"' and pe_date <= '"+to_date+"' ORDER BY pe_car_no,pe_date,pe_seq ASC"
                        Rs_pen.Open Sql, Dbconn, 1
						if not Rs_pen.EOF or not Rs_pen.BOF then
		 %>                
                <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="7%" >
                              <col width="6%" >
							  <col width="8%" >
							  <col width="8%" >
                              <col width="6%" >
							  <col width="8%" >
                              <col width="6%" >
							  <col width="*" >
							  <col width="6%" >
                              <col width="6%" >
                              <col width="8%" >
                              <col width="6%" >
                              <col width="8%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <th class="first" scope="col">차량번호</th>
                                <th scope="col">차종</th>
								<th scope="col">운행자</th>
								<th scope="col">부서</th>
                                <th scope="col">위반일자</th>
								<th scope="col">위반내용</th>
								<th scope="col">과태료</th>
								<th scope="col">위반장소</th>
                                <th scope="col">납입일자</th>
                                <th scope="col">통보일자</th>
                                <th scope="col">통보방법</th>
                                <th scope="col">미납</th>
                                <th scope="col">비고</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until Rs_pen.eof
						
						  car_no = Rs_pen("pe_car_no")
						  if Rs_pen("pe_in_date") = "1900-01-01"  then
	                               pe_in_date = ""
							   else 
							       pe_in_date = Rs_pen("pe_in_date")
	                       end if
	                       if Rs_pen("pe_notice_date") = "1900-01-01" then
	                               pe_notice_date = ""
							   else 
							       pe_notice_date = Rs_pen("pe_notice_date")
	                       end if
							  
		                   Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
                           Set rs_car = DbConn.Execute(SQL)
		                   if not rs_car.eof then
		                        	car_name = rs_car("car_name")
		                    		car_year = rs_car("car_year")
			                    	car_reg_date = rs_car("car_reg_date")
		                    		car_use_dept = rs_car("car_use_dept")
	                    			car_company = rs_car("car_company")
	                     			car_use = rs_car("car_use")
									car_owner = rs_car("car_owner")
	                    			owner_emp_name = rs_car("owner_emp_name")
	                    			owner_emp_no = rs_car("owner_emp_no")
	                     			oil_kind = rs_car("oil_kind")
	                          else
	                     		    car_name = ""
	                    			car_year = ""
			                      	car_reg_date = ""
			                    	car_use_dept = ""
		                    		car_company = ""
		                    		car_use = ""
									car_owner = ""
	                    			owner_emp_name = ""
		                    		owner_emp_no = ""
	                    			oil_kind = ""
                           end if
                           rs_car.close()
						
	           			%>
							<tr>
                                <td class="first"><%=Rs_pen("pe_car_no")%>&nbsp;</td>
                                <td><%=car_name%>&nbsp;</td>
                                <td><%=owner_emp_name%>(<%=owner_emp_no%>)&nbsp;</td>
                                <td><%=car_use_dept%>&nbsp;</td>
                                <td><%=Rs_pen("pe_date")%>&nbsp;</td>
								<td class="left"><%=Rs_pen("pe_comment")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(Rs_pen("pe_amount"),0)%>&nbsp;</td>
                                <td class="left"><%=Rs_pen("pe_place")%>&nbsp;</td>
                                <td><%=pe_in_date%>&nbsp;</td>
                                <td><%=pe_notice_date%>&nbsp;</td>
                                <td class="left"><%=Rs_pen("pe_notice")%>&nbsp;</td>
                                <td class="left"><%=Rs_pen("pe_default")%>&nbsp;</td>
                                <td class="left"><%=Rs_pen("pe_bigo")%>&nbsp;</td>
							</tr>
						<%
							Rs_pen.movenext()
						loop
						%>
                            <tr>
								<td colspan="4" class="first" style="background:#ffe8e8;">총계</td>
                                <td style="background:#ffe8e8;">과태료 계</td>
                                <td colspan="2" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_amount,0)%>&nbsp;</td>
                                <td style="background:#ffe8e8;">납입 계</td>
                                <td colspan="2" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_in_amt,0)%>&nbsp;</td>
                                <td style="background:#ffe8e8;">미납 계</td>
                                <td colspan="2" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(jan_amount,0)%>&nbsp;</td>
							</tr>
						</tbody>
                   </table>  
         <% 
		                Rs_pen.close()
			  end if %>
                   
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<td width="15%">
					<div class="btnleft">
                    <a href="insa_excel_car_info_total.asp?car_no=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
                    <td width="20%">
                    <div class="btnRight">
                    <a href="#" onClick="pop_Window('insa_car_infor_print.asp?car_no=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>','car_infor_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">출력</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

