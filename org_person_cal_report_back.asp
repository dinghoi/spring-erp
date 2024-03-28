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
team=Request.form("team")
saupbu=Request.form("saupbu")
if position = "팀장" then
	team = request.cookies("nkpmg_user")("coo_team")
	org_view = "조직명 : " + team
end if
if position = "사업부장" then
	saupbu = request.cookies("nkpmg_user")("coo_saupbu")
	org_view = "조직명 : " + saupbu
end if
if position = "본부장" and saupbu = "" then
	saupbu = "본인"
'	sql = "select org_name from emp_org_mst where (org_level = '사업부') and (org_bonbu = '"&bonbu&"') group by org_name order by org_saupbu asc"
'    rs_org.Open sql, Dbconn, 1	
'	saupbu = rs_org("org_name")
'	rs_org.close()
end if
if cost_grade = "0" and saupbu = "" then
	sql = "select org_saupbu from emp_org_mst where (org_level = '사업부') group by org_name order by org_saupbu asc"
    rs_org.Open sql, Dbconn, 1	
	saupbu = rs_org("org_saupbu")
	rs_org.close()
end if

if cost_month = "" then
	cost_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
end If

if position = "팀장" then
	sql = "select * from emp_master where emp_team = '"&team&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&from_date&"') order by emp_team, emp_name"
  else
	if saupbu = "본인" then
		sql = "select * from emp_master where emp_no = '"&user_id&"'"
	  else
		sql = "select * from emp_master where emp_saupbu = '"&saupbu&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&from_date&"') order by emp_team, emp_name"
	end if
end if 
rs_emp.Open sql, Dbconn, 1
	
title_line = "조직별 개인별 비용 정산 현황"
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
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="org_person_cal_report.asp" method="post" name="frm">
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
							<% if position = "팀장" or position = "사업부장" then	%>
                                <label>
								<strong><%=org_view%></strong>
								</label>
							<% end if	%>
<% if position = "본부장" or cost_grade = "0" then	%>
                                <label>
								<strong>사업부</strong>
							<%
								if position = "본부장" then
									sql_org="select org_name from emp_org_mst where (org_level = '사업부') and (org_bonbu = '"&bonbu&"') group by org_name order by org_saupbu asc"
								  else
									sql_org="select org_name from emp_org_mst where (org_level = '사업부') group by org_name order by org_saupbu asc"
								end if							  
                                rs_org.Open sql_org, Dbconn, 1
                                %>
                                <select name="saupbu" id="saupbu" style="width:150px">
							<% if position = "본부장" then	%>
                                    <option value="본인" <%If saupbu = "본인" then %>selected<% end if %>>본인</option>
							<% end if %>
<% 
								do until rs_org.eof
							%>
          							<option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = saupbu  then %>selected<% end if %>><%=rs_org("org_name")%></option>
          					<%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							%>
                                </select>
								</label>
							<% end if	%>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="3%" >
							<col width="6%" >
							<col width="3%" >
							<col width="6%" >
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
							<col width="7%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="3" class="first" scope="col">팀</th>
								<th rowspan="3" scope="col">사용자</th>
								<th style=" border-bottom:1px solid #e3e3e3;" scope="col">야특근</th>
								<th colspan="9" style=" border-bottom:1px solid #e3e3e3;" scope="col">현금 사용</th>
								<th rowspan="3" scope="col">주유카드</th>
								<th rowspan="3" scope="col">정산금액</th>
								<th rowspan="3" scope="col">수리비</th>
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
							  <th scope="col">금액</th>
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
						<tbody>
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
						sum_card_price = 0
						sum_cash_tot_cost = 0
						sum_return_cash = 0

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
						tot_card_price = 0
						tot_cash_tot_cost = 0
						tot_return_cash = 0

						if isnull(rs_emp("emp_team")) or rs_emp("emp_team") = "" then	
							bi_team = ""
						  else
							bi_team = rs_emp("emp_team")
						end if
						do until rs_emp.eof
							if isnull(rs_emp("emp_team")) or rs_emp("emp_team") = "" then
								emp_team = ""
							  else
							  	emp_team = rs_emp("emp_team")
							end if
							
							if bi_team <> emp_team then
						%>
							<tr>
								<td colspan="2" bgcolor="#EEFFFF" class="first">소계</td>
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
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_card_price,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_return_cash,0)%></td>
								<td bgcolor="#EEFFFF" class="right">&nbsp;</td>
								<td bgcolor="#EEFFFF" class="right">&nbsp;</td>
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
								sum_card_price = 0
								sum_cash_tot_cost = 0
								sum_return_cash = 0
								bi_team = emp_team
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
								card_cost = 0
								card_cost_vat = 0
								return_cash = 0
								tot_km = gas_km + diesel_km + gasol_km
								tot_cost = gas_cost + diesel_cost + gasol_cost
								card_price = card_cost + card_cost_vat
								cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
							  else
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
								card_cost = rs("card_cost")	 
								card_cost_vat = rs("card_cost_vat")	 
								return_cash = rs("return_cash")	 
								tot_km = gas_km + diesel_km + gasol_km
								tot_cost = gas_cost + diesel_cost + gasol_cost
								card_price = card_cost + card_cost_vat
								cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost
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
							sum_card_price = sum_card_price + card_price
							sum_cash_tot_cost = sum_cash_tot_cost + cash_tot_cost
							sum_return_cash = sum_return_cash + return_cash

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
							tot_card_price = tot_card_price + card_price
							tot_cash_tot_cost = tot_cash_tot_cost + cash_tot_cost
							tot_return_cash = tot_return_cash + return_cash
						%>
							<tr>
								<td class="first"><%=rs_emp("emp_team")%>&nbsp;</td>
								<td><%=rs_emp("emp_name")%>&nbsp;<%=rs_emp("emp_job")%></td>
								<td class="right"><%=formatnumber(overtime_cost,0)%></td>
								<td class="right"><%=formatnumber(general_cost,0)%></td>
								<td class="right"><%=formatnumber(fare_cost,0)%></td>
								<td class="right"><%=formatnumber(tot_km,0)%></td>
								<td class="right"><%=formatnumber(tot_cost,0)%></td>
								<td class="right"><%=formatnumber(somopum_cost,0)%></td>
								<td class="right"><%=formatnumber(oil_cash_cost,0)%></td>
								<td class="right"><%=formatnumber(parking_cost,0)%></td>
								<td class="right"><%=formatnumber(toll_cost,0)%></td>
								<td class="right"><%=formatnumber(cash_tot_cost,0)%></td>
								<td class="right"><%=formatnumber(card_price,0)%></td>
								<td class="right"><%=formatnumber(return_cash,0)%></td>
								<td class="right">&nbsp;</td>
								<td class="right">&nbsp;</td>
							</tr>
						<%
							rs_emp.movenext()
						loop
						%>
							<tr>
								<td colspan="2" bgcolor="#EEFFFF" class="first">소계</td>
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
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_card_price,0)%></td>
								<td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_return_cash,0)%></td>
								<td bgcolor="#EEFFFF" class="right">&nbsp;</td>
								<td bgcolor="#EEFFFF" class="right">&nbsp;</td>
							</tr>
						<% if saupbu <> "본인" then	%> 
							<tr>
								<th colspan="2" class="first">총계</th>
								<th class="right"><%=formatnumber(tot_overtime_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_general_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_fare_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_tot_km,0)%></th>
								<th class="right"><%=formatnumber(tot_tot_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_somopum_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_oil_cash_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_parking_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_toll_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_cash_tot_cost,0)%></th>
								<th class="right"><%=formatnumber(tot_card_price,0)%></th>
								<th class="right"><%=formatnumber(tot_return_cash,0)%></th>
								<th class="right">&nbsp;</th>
								<th class="right">&nbsp;</th>
							</tr>
						<% end if	%>
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
			</form>
		</div>				
	</div>        				
	</body>
</html>

