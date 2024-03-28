<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim win_sw
dim ce_sum(1000,3,11)
dim ce_tab(1000,2)

for i = 1 to 1000
	ce_tab(i,1) = ""
	ce_tab(i,2) = ""
next

for i = 1 to 1000
	for j = 1 to 3
		for k = 1 to 11
			ce_sum(i,j,k) = 0
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

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Set rs_sign = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

i = 0
if position = "팀장" then
	sql = "select * from memb where emp_company ='"&emp_company&"' and bonbu ='"&bonbu&"' and saupbu ='"&saupbu&"' and team ='"&team& _
	"' order by user_name" 
	rs_memb.Open sql, Dbconn, 1
	do until rs_memb.eof
		i = i + 1
		ce_tab(i,1) = rs_memb("user_id")
		ce_tab(i,2) = rs_memb("user_name")
		rs_memb.movenext()
	loop
end if

' 일반비용
for i = 1 to 1000
	sql = "select pay_method,pay_yn,count(slip_seq) as c_cnt,sum(cost) as cost from general_cost where (emp_no='"&ce_tab(i,1)&"') "& _
	"and (slip_gubun = '비용') and (cancel_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by pay_method,pay_yn"
	rs.Open sql, Dbconn, 1
	do until rs.eof
		if rs("pay_method") = "현금" then
			if rs("pay_yn") = "N" then
				ce_sum(i,1,1) = ce_sum(i,1,1) + cint(rs("c_cnt"))
				ce_sum(i,1,2) = ce_sum(i,1,2) + cdbl(rs("cost"))
			  else
				ce_sum(i,1,3) = ce_sum(i,1,3) + cint(rs("c_cnt"))
				ce_sum(i,1,4) = ce_sum(i,1,4) + cdbl(rs("cost"))
			end if
		  else
			ce_sum(i,1,5) = ce_sum(i,1,5) + cint(rs("c_cnt"))
			ce_sum(i,1,6) = ce_sum(i,1,6) + cdbl(rs("cost"))
		end if			  													  
		rs.movenext()
	loop
	rs.close()
next

' 야특근
for i = 1 to 1000
	sql = "select cancel_yn,count(work_date) as c_cnt,sum(overtime_amt) as cost from overtime where (mg_ce_id='"&ce_tab(i,1)&"') "& _
	"and (work_date >='"&from_date&"' and work_date <='"&to_date&"') and (cancel_yn = 'N') group by cancel_yn"
'	response.write(sql)
	rs.Open sql, Dbconn, 1
	do until rs.eof
		ce_sum(i,2,1) = ce_sum(i,2,1) + cint(rs("c_cnt"))
		ce_sum(i,2,2) = ce_sum(i,2,2) + cdbl(rs("cost"))
		rs.movenext()
	loop
	rs.close()
next

' 교통비
for i = 1 to 1000
	sql = "select * from transit_cost where (mg_ce_id='"&ce_tab(i,1)&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cancel_yn = 'N')"
	rs.Open sql, Dbconn, 1
	do until rs.eof
		if rs("car_owner") = "개인" then
			ce_sum(i,3,1) = ce_sum(i,3,1) + rs("far")
		end if
		if rs("payment") = "현금" then		
			ce_sum(i,3,2) = ce_sum(i,3,2) + rs("fare")
		  else
			ce_sum(i,3,7) = ce_sum(i,3,7) + rs("fare")
		end if
		if rs("car_owner") = "회사" then
			if rs("oil_pay") = "현금" then		
				ce_sum(i,3,3) = ce_sum(i,3,3) + rs("oil_price")
			  else
				ce_sum(i,3,8) = ce_sum(i,3,8) + rs("oil_price")
			end if
		end if
		if rs("car_owner") = "회사" then
			if rs("repair_pay") = "현금" then		
				ce_sum(i,3,4) = ce_sum(i,3,4) + rs("repair_cost")
			  else
				ce_sum(i,3,9) = ce_sum(i,3,9) + rs("repair_cost")
			end if
		end if
		if rs("parking_pay") = "현금" then		
			ce_sum(i,3,5) = ce_sum(i,3,5) + rs("parking")
		  else
			ce_sum(i,3,10) = ce_sum(i,3,10) + rs("parking")
		end if
		if rs("toll_pay") = "현금" then		
			ce_sum(i,3,6) = ce_sum(i,3,6) + rs("toll")
		  else
			ce_sum(i,3,11) = ce_sum(i,3,11) + rs("toll")
		end if
		rs.movenext()
	loop
	rs.close()
next

for i = 1 to 1000
	for j = 1 to 3
		for k = 1 to 11
			ce_sum(0,j,k) = ce_sum(0,j,k) + ce_sum(i,j,k)
		next
	next
next

title_line = "팀별 개인별 비용 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
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
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="org_emp_cost_mg.asp" method="post" name="frm">
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
								<strong>조직정보 : </strong><%=org_name%>&nbsp;<%=position%>
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
							<col width="3%" >
							<col width="5%" >
							<col width="3%" >
							<col width="5%" >
							<col width="3%" >
							<col width="5%" >
							<col width="3%" >
							<col width="5%" >
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="3" class="first" scope="col">일자</th>
								<th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3;">일 반 비 용</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">야 특 근</th>
								<th colspan="13" scope="col" style=" border-bottom:1px solid #e3e3e3;">교 통 비</th>
							</tr>
							<tr>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;border-left:1px solid #e3e3e3;">현금미지급</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">선지급</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">법인카드</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">신청금액</th>
							  <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">주행비용</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">현 금</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">법인카드</th>
						  </tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">건수</th>
							  <th scope="col">금액</th>
							  <th scope="col">건수</th>
							  <th scope="col">금액</th>
							  <th scope="col">건수</th>
							  <th scope="col">금액</th>
							  <th scope="col">수량</th>
							  <th scope="col">금액</th>
							  <th scope="col">주행</th>
							  <th scope="col">주유비</th>
							  <th scope="col">소모품</th>
							  <th scope="col">교통비</th>
							  <th scope="col">유류비</th>
							  <th scope="col">수리비</th>
							  <th scope="col">주차비</th>
							  <th scope="col">통행료</th>
							  <th scope="col">교통비</th>
							  <th scope="col">유류비</th>
							  <th scope="col">수리비</th>
							  <th scope="col">주차비</th>
							  <th scope="col">통행료</th>
                          </tr>
						</thead>
						<tbody>
						<%
						for i = 1 to 1000
							j = 0
							j = ce_sum(i,1,1) + ce_sum(i,1,3) + ce_sum(i,1,5) + ce_sum(i,2,1) + ce_sum(i,2,3)					
							for k = 1 to 11
								j = j + ce_sum(i,3,k)
							next
							if j <> 0 then
						%>
							<tr>
								<td class="first"><%=ce_tab(i,1)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,1,1),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,1,2),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,1,3),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,1,4),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,1,5),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,1,6),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,2,1),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,2,2),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,1),0)%></td>
								<td class="right">&nbsp;</td>
								<td class="right"><%=formatnumber(ce_sum(i,3,1)*25,0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,2),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,3),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,4),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,5),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,6),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,7),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,8),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,9),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,10),0)%></td>
								<td class="right"><%=formatnumber(ce_sum(i,3,11),0)%></td>
							</tr>
						<%
							end if
						next
						%>
							<tr>
								<th class="first">계</th>
								<th class="right"><%=formatnumber(ce_sum(0,1,1),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,1,2),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,1,3),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,1,4),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,1,5),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,1,6),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,2,1),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,2,2),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,1),0)%></th>
								<th class="right">&nbsp;</th>
								<th class="right"><%=formatnumber(ce_sum(0,3,1)*25,0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,2),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,3),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,4),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,5),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,6),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,7),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,8),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,9),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,10),0)%></th>
								<th class="right"><%=formatnumber(ce_sum(0,3,11),0)%></th>
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

