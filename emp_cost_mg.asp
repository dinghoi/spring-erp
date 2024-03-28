<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim win_sw
dim day_sum(31,10,2)
dim day_tab(31)

for i = 1 to 31
	day_tab(i) = ""
next

for i = 1 to 31
	for j = 1 to 10
		for k = 1 to 2
			day_sum(i,j,k) = 0
		next
	next
next
	 
slip_month=Request.form("slip_month")

if slip_month = "" then
	slip_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
end If

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
sign_month = slip_month

for i = 1 to 31
	if i < 10 then
		d = "0" + cstr(i)
	  else
	  	d = cstr(i)
	end if
	work_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-" + d
	day_tab(i) = work_date
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_sign = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 차량 정보
sql = "select * from car_info where owner_emp_no ='"&user_id&"'"
set rs_car=dbconn.execute(sql)
if rs_car.eof then
	car_info = "차량없음"
  else  	
	car_info = rs_car("car_owner") + "차량 , 차종 : " + rs_car("car_name") + " , 유종 : " + rs_car("oil_kind")
end if
rs_car.Close()		

' 일반비용
for i = 1 to 31
	sql = "select pay_method,pay_yn,count(slip_seq) as c_cnt,sum(cost) as cost from general_cost where (emp_no='"&user_id&"') "& _
	"and (slip_gubun = '비용') and (tax_bill_yn = 'N' or isnull(tax_bill_yn)) and (cancel_yn = 'N') and (slip_date ='"&day_tab(i)&"') group by pay_method,pay_yn"
'Response.write sql&"<br>"
	rs.Open sql, Dbconn, 1
	do until rs.eof
		if rs("pay_method") = "현금" then
			if rs("pay_yn") = "N" then
				day_sum(i,1,1) = day_sum(i,1,1) + cint(rs("c_cnt"))
				day_sum(i,1,2) = day_sum(i,1,2) + cdbl(rs("cost"))
			  else
				day_sum(i,2,1) = day_sum(i,2,1) + cint(rs("c_cnt"))
				day_sum(i,2,2) = day_sum(i,2,2) + cdbl(rs("cost"))
			end if
		end if			  													  
		rs.movenext()
	loop
	rs.close()
next

' 야특근
for i = 1 to 31
	sql = "select cancel_yn,count(work_date) as c_cnt,sum(overtime_amt) as cost from overtime where (mg_ce_id='"&user_id&"') "& _
	"and (work_date ='"&day_tab(i)&"') and (cancel_yn = 'N') group by cancel_yn"
'	response.write(sql)
	rs.Open sql, Dbconn, 1
	do until rs.eof
		day_sum(i,3,1) = day_sum(i,3,1) + cint(rs("c_cnt"))
		day_sum(i,3,2) = day_sum(i,3,2) + cdbl(rs("cost"))
		rs.movenext()
	loop
	rs.close()
next

' 교통비
for i = 1 to 31
	sql = "select * from transit_cost where (mg_ce_id='"&user_id&"') and (run_date ='"&day_tab(i)&"') and (cancel_yn = 'N')"
	rs.Open sql, Dbconn, 1
	do until rs.eof
		if rs("car_owner") = "개인" then
			day_sum(i,4,1) = day_sum(i,4,1) + rs("far")
			day_sum(i,4,2) = day_sum(i,4,2) + int(rs("far")) * 25
		end if

		day_sum(i,5,2) = day_sum(i,5,2) + rs("fare")

		if rs("car_owner") = "회사" then
			day_sum(i,6,2) = day_sum(i,6,2) + rs("oil_price")
			day_sum(i,7,2) = day_sum(i,7,2) + rs("repair_cost")
		end if

		day_sum(i,8,2) = day_sum(i,8,2) + rs("parking")
		day_sum(i,9,2) = day_sum(i,9,2) + rs("toll")

		rs.movenext()
	loop
	rs.close()
next

' 카드사용
for i = 1 to 31
	sql = "select count(*) as c_cnt,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (emp_no='"&user_id&"') and (slip_date ='"&day_tab(i)&"')"
		
	Set rs = Dbconn.Execute (sql)
	if cint(rs("c_cnt")) <>  0 then
		day_sum(i,10,1) = cint(rs("c_cnt"))
		day_sum(i,10,2) = cdbl(rs("cost")) + cdbl(rs("cost_vat"))
	  else
		day_sum(i,10,1) = 0
		day_sum(i,10,2) = 0  
	end if
	rs.close()
next

for i = 1 to 31
	for j = 1 to 10
		for k = 1 to 2
			day_sum(0,j,k) = day_sum(0,j,k) + day_sum(i,j,k)
		next
	next
next

title_line = "개인별 일자별 비용 사용현황"
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
				if (document.frm.slip_month.value == "") {
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
				<form action="emp_cost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>발생년월&nbsp;</strong>(예201401) : 
                                	<input name="slip_month" type="text" value="<%=slip_month%>" style="width:70px">
								</label>
								<label>
								<strong>차량정보 : </strong><%=car_info%>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="4%" >
							<col width="7%" >
							<col width="4%" >
							<col width="7%" >
							<col width="4%" >
							<col width="7%" >
							<col width="5%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="4%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="3" class="first" scope="col">일자</th>
								<th colspan="4" style=" border-bottom:1px solid #e3e3e3;" scope="col">일 반 비 용</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">야 특 근</th>
								<th colspan="7" style=" border-bottom:1px solid #e3e3e3;" scope="col">교 통 비</th>
								<th colspan="2" rowspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">법인카드</th>
							</tr>
							<tr>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;border-left:1px solid #e3e3e3;">현금미지급</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">선지급</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">신청금액</th>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">주행비용</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">현 금 사 용</th>
						  </tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">건수</th>
							  <th scope="col">금액</th>
							  <th scope="col">건수</th>
							  <th scope="col">금액</th>
							  <th scope="col">수량</th>
							  <th scope="col">금액</th>
							  <th scope="col">주행</th>
							  <th scope="col">소모품</th>
							  <th scope="col">대중교통</th>
							  <th scope="col">주유비</th>
							  <th scope="col">수리비</th>
							  <th scope="col">주차비</th>
							  <th scope="col">통행료</th>
							  <th scope="col">건수</th>
							  <th scope="col">금액</th>
						  </tr>
						</thead>
						<tbody>
						<%
						for i = 1 to 31
							j = 0
							for k = 1 to 10
								j = j + day_sum(i,k,2)
							next
							if j <> 0 then
						%>
							<tr>
								<td class="first"><%=replace(mid(day_tab(i),6),"-","/")%></td>
								<td class="right"><%=formatnumber(day_sum(i,1,1),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,1,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,2,1),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,2,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,3,1),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,3,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,4,1),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,4,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,5,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,6,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,7,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,8,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,9,2),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,10,1),0)%></td>
								<td class="right"><%=formatnumber(day_sum(i,10,2),0)%></td>
							</tr>
						<%
							end if
						next
						%>
							<tr>
								<th class="first">계</th>
								<th class="right"><%=formatnumber(day_sum(0,1,1),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,1,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,2,1),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,2,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,3,1),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,3,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,4,1),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,4,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,5,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,6,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,7,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,8,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,9,2),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,10,1),0)%></th>
								<th class="right"><%=formatnumber(day_sum(0,10,2),0)%></th>
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
			</form>
		</div>				
	</div>        				
	</body>
</html>

