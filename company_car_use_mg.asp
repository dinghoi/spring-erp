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
	
if cost_month = "" then
	cost_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
end If

from_date = mid(cost_month,1,4) + "-" + mid(cost_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
	
be_yy = int(mid(cost_month,1,4))
be_mm = int(mid(cost_month,5) - 1)
if be_mm = 0 then
	be_month = cstr(be_yy - 1) + "12"
  else
	be_month = cstr(be_yy) + right("0" + cstr(be_mm),2)
end if

sql = "select car_no, car_name, oil_kind, mg_ce_id, user_name, team, org_name, sum(far) as sum_far, sum(oil_price) as sum_price, sum(repair_cost) as sum_repair  from transit_cost where (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cancel_yn = 'N') and car_owner = '회사' group by car_no, mg_ce_id order by car_no, user_name "

title_line = "회사 차량 추정비용 대비 실 사용금액"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
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
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/account_cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="company_car_use_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>발생년월&nbsp;</strong>(예201401) : 
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px" maxlength="6">
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
							<col width="6%" >
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="10%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
							  <th colspan="3" class="first" style=" border-bottom:1px solid #e3e3e3;" scope="col">차량 정보</th>
							  <th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">운전자 정보</th>
							  <th rowspan="2" scope="col">주행거리</th>
							  <th colspan="4" style=" border-bottom:1px solid #e3e3e3;" scope="col">추정 금액</th>
							  <th colspan="4" style=" border-bottom:1px solid #e3e3e3;" scope="col">실 사용금액</th>
							  <th rowspan="2" scope="col">차액</th>
						  </tr>
							<tr>
							  <th class="first" scope="col">차량번호</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">차종</th>
							  <th scope="col">유종</th>
							  <th scope="col">운전자</th>
							  <th scope="col">조직명</th>
							  <th scope="col">단가</th>
							  <th scope="col">금액</th>
							  <th scope="col">소모품</th>
							  <th scope="col">소계</th>
							  <th scope="col">현금주유비</th>
							  <th scope="col">주유카드</th>
							  <th scope="col">수리비</th>
							  <th scope="col">소계</th>
						  </tr>
						</thead>
						<tbody>
					<%
						rs.Open sql, Dbconn, 1
						do until rs.eof
							if rs("team") = "본사팀" or rs("team") = "SM1팀" or rs("team") = "Repair팀" or rs("team") = "SM2팀" then
								oil_unit_id = "1"
							  else
								oil_unit_id = "2"
							end if
							
							sql = "select * from oil_unit where oil_unit_month = '"&cost_month&"' and oil_unit_id = '"&oil_unit_id&"' and oil_kind = '"&rs("oil_kind")&"'"
							set rs_etc=dbconn.execute(sql)
							if rs_etc.eof or rs_etc.bof then
								oil_unit = 1
							  else
								oil_unit = rs_etc("oil_unit_average")
							end if								
							rs_etc.close()

							if oil_kind = "가스" then
								oil_cost = round(cdbl(rs("sum_far")) * oil_unit / 7)
							  else
								oil_cost = round(cdbl(rs("sum_far")) * oil_unit / 10)
							end if
							somopum = cdbl(rs("sum_far")) * 25
							
' 주유 카드사용
							juyoo_card_price = 0
							sql = "select count(*) as c_cnt,sum(price) as price from card_slip where (emp_no='"&user_id&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and card_type like '%주유%'"
								
							Set rs_etc = Dbconn.Execute (sql)
							if cint(rs_etc("c_cnt")) <>  0 then
								juyoo_card_price = cdbl(rs_etc("price"))
							end if
							rs_etc.close()
					%>
							<tr>
								<td class="first"><%=rs("car_no")%></td>
								<td><%=rs("car_name")%></td>
								<td><%=rs("oil_kind")%></td>
								<td><%=rs("user_name")%></td>
								<td><%=rs("org_name")%></td>
								<td class="right"><%=formatnumber(rs("sum_far"),0)%></td>
								<td class="right"><%=formatnumber(oil_unit,0)%></td>
								<td class="right"><%=formatnumber(oil_cost,0)%></td>
								<td class="right"><%=formatnumber(somopum,0)%></td>
								<td class="right"><%=formatnumber(oil_cost + somopum,0)%></td>
								<td class="right"><%=formatnumber(rs("sum_price"),0)%></td>
								<td class="right"><%=formatnumber(juyoo_card_price,0)%></td>
								<td class="right"><%=formatnumber(rs("sum_repair"),0)%></td>
								<td class="right"><%=formatnumber(cdbl(rs("sum_price")) + juyoo_card_price + cdbl(rs("sum_repair")),0)%></td>
								<td class="right"><%=formatnumber(cdbl(rs("sum_price")) + juyoo_card_price + cdbl(rs("sum_repair")) - oil_cost - somopum,0)%></td>
							</tr>
					<%
							rs.movenext()
						loop
					%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
					<div class="btnCenter">
                    <a href="company_car_use_excel.asp?cost_month=<%=cost_month%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td width="85%">
                    </td>
			      </tr>
				  </table>
				<br>
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

