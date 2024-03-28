<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	Dim Rs
	Dim Repeat_Rows
	Dim from_date
	Dim to_date
	Dim win_sw
	
	win_sw = "close"
	
	ck_sw=Request("ck_sw")
	Page=Request("page")
	
	if ck_sw = "y" Then
		bill_id = request("bill_id")
		bill_month = request("bill_month")
		cost_reg_yn = request("cost_reg_yn")
		end_yn = request("end_yn")
	else
		bill_id = request.form("bill_id")
		bill_month = request.form("bill_month")
		cost_reg_yn = request.form("cost_reg_yn")
		end_yn = request.form("end_yn")
	end if
	
	if bill_month = "" then
		bill_month = mid(now(),1,4) + mid(now(),6,2)
		bill_id = "1"
		cost_reg_yn = "T"
		end_yn = "T"
	end if
'	response.write(end_yn)	
	from_date = mid(bill_month,1,4) + "-" + mid(bill_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))
	
	pgsize = 10 ' 화면 한 페이지 
	
	If Page = "" Then
		Page = 1
		start_page = 1
	End If
	stpage = int((page - 1) * pgsize)
	
	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	Set RsCount = Server.CreateObject("ADODB.Recordset")
	Set Rscost = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	base_sql = "select * from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (bill_id = '"&bill_id&"') "
	
	if cost_reg_yn = "T" then
		cost_reg_sql = " "
	  else
		cost_reg_sql = " and ( cost_reg_yn = '"&cost_reg_yn&"' ) "
	end if
	if end_yn = "T" then
		end_sql = " "
	  else
		end_sql = " and ( end_yn = '"&end_yn&"' ) "
	end if
	
	order_sql = " ORDER BY bill_date ASC"
' 비용 등록 여부 확인	
	sql = "select count(*) from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (cost_reg_yn = 'Y') and (bill_id = '"&bill_id&"') "
	Set rscost = Dbconn.Execute (sql)
	
	cost_record = cint(rscost(0)) 'Result.RecordCount
' 비용등록 확인 끝
' 비용 미등록 여부 확인	
	sql = "select count(*) from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (cost_reg_yn = 'N') and (bill_id = '"&bill_id&"') "
	Set rsmicost = Dbconn.Execute (sql)
	
	mi_record = cint(rsmicost(0)) 'Result.RecordCount
' 비용 미등록 확인 끝
	sql = "select count(*) from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (bill_id = '"&bill_id&"') " + cost_reg_sql + end_sql
	Set RsCount = Dbconn.Execute (sql)
	
	tottal_record = cint(RsCount(0)) 'Result.RecordCount
	
	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF
	
	sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (bill_id = '"&bill_id&"') " + cost_reg_sql + end_sql
	Set rs_sum = Dbconn.Execute (sql)
	if isnull(rs_sum("price")) then
		sum_price = 0
		sum_cost = 0
		sum_cost_vat = 0
	  else
		sum_price = cdbl(rs_sum("price"))
		sum_cost = cdbl(rs_sum("cost"))
		sum_cost_vat = cdbl(rs_sum("cost_vat"))
	end if
	
	sql = base_sql + cost_reg_sql + end_sql + order_sql + " limit "& stpage & "," &pgsize 
	Rs.Open Sql, Dbconn, 1
'response.write sql	

	title_line = "이세로 세금계산서 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.bill_id.value == "") {
					alert ("계산서 유형을 선택하세요");
					return false;
				}	
				if (document.frm.bill_month.value == "") {
					alert ("년월을 선택하세요");
					return false;
				}	
				if (document.frm.cost_reg_yn.value == "") {
					alert ("비용등록 여부를 선택하세요");
					return false;
				}	
				return true;
			}
			function upload_cancel() 
				{
				a=confirm('업로드를 취소하겠습니까?')
				if (a==true) {
					document.frm.action = "tax_bill_upload_cancel.asp";
               		document.frm.submit();
					return true;
				}
				return false;
				}
			function end_process() 
				{
				a=confirm('마감하시겠습니까?')
				if (a==true) {
					document.frm.action = "tax_esero_end.asp";
               		document.frm.submit();
					return true;
				}
				return false;
				}
			function cancel_process() 
				{
				a=confirm('취소하시겠습니까?')
				if (a==true) {
					document.frm.action = "tax_esero_end_cancel.asp";
               		document.frm.submit();
					return true;
				}
				return false;
				}
		</script>
	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/tax_bill_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="tax_esero_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조회조건</dt>
                        <dd>
                            <p>
								<label>
								<strong>계산서 유형 : </strong>
                              	<input type="radio" name="bill_id" value="1" <% if bill_id = "1" then %>checked<% end if %> style="width:25px">매입
                                <input type="radio" name="bill_id" value="2" <% if bill_id = "2" then %>checked<% end if %> style="width:25px">매출
								</label>
								<label>
								<strong>발행년월 : </strong>
                                	<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
								<label>
								<strong>비용등록여부 : </strong>
                              	<input type="radio" name="cost_reg_yn" value="T" <% if cost_reg_yn = "T" then %>checked<% end if %> style="width:25px">전체
                                <input type="radio" name="cost_reg_yn" value="Y" <% if cost_reg_yn = "Y" then %>checked<% end if %> style="width:25px">등록
                                <input type="radio" name="cost_reg_yn" value="N" <% if cost_reg_yn = "N" then %>checked<% end if %> style="width:25px">미등록
								</label>
								<label>
								<strong>마감여부 : </strong>
                              	<input type="radio" name="end_yn" value="T" <% if end_yn = "T" then %>checked<% end if %> style="width:25px">전체
                                <input type="radio" name="end_yn" value="Y" <% if end_yn = "Y" then %>checked<% end if %> style="width:25px">Yes
                                <input type="radio" name="end_yn" value="N" <% if end_yn = "N" then %>checked<% end if %> style="width:25px">No
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
							<col width="10%" >
							<col width="7%" >
							<col width="11%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="3%" >
							<col width="12%" >
							<col width="*" >
							<col width="3%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">발행일</th>
								<th scope="col">계산서소유회사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">상호</th>
								<th scope="col">대표자명</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">청구</th>
								<th scope="col">계산서이메일</th>
								<th scope="col">품목명</th>
								<th scope="col">마감</th>
								<th scope="col">비용</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><%=formatnumber(tottal_record,0)%>&nbsp;건</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(sum_price,0)%></td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
								<td class="right"><%=formatnumber(sum_cost_vat,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
'						end_yn = "N"
						do until rs.eof
'							end_yn = rs("end_yn")
							if rs("cost_reg_yn") = "Y" then
								cost_reg_view = "등록"
							  else
							  	cost_reg_view = "미등록"
							end if
							if bill_id = "1" then
								email_view = rs("send_email")
							  else
							  	email_view = rs("receive_email")
							end if
						%>
							<tr>
								<td class="first"><%=rs("bill_date")%></td>
								<td><%=rs("owner_company")%></td>
								<td><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%></td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("trade_owner")%></td>
								<td class="right"><%=formatnumber(rs("price"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("bill_collect")%></td>
								<td><%=email_view%>&nbsp;</td>
								<td class="left"><%=rs("tax_bill_memo")%></td>
								<td><%=rs("end_yn")%></td>
								<td><%=cost_reg_view%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="24%">
					<div class="btnCenter">
				<% if cost_record = 0 and tottal_record > 0 then	%>
					<a href="#" onClick="upload_cancel()" class="btnType04">업로드취소</a>
				<% end if	%>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="tax_esero_mg.asp?page=<%=first_page%>&bill_id=<%=bill_id%>&bill_month=<%=bill_month%>&cost_reg_yn=<%=cost_reg_yn%>&end_yn=<%=end_yn%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="tax_esero_mg.asp?page=<%=intstart -1%>&bill_id=<%=bill_id%>&bill_month=<%=bill_month%>&cost_reg_yn=<%=cost_reg_yn%>&end_yn=<%=end_yn%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="tax_esero_mg.asp?page=<%=i%>&bill_id=<%=bill_id%>&bill_month=<%=bill_month%>&cost_reg_yn=<%=cost_reg_yn%>&end_yn=<%=end_yn%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="tax_esero_mg.asp?page=<%=intend+1%>&bill_id=<%=bill_id%>&bill_month=<%=bill_month%>&cost_reg_yn=<%=cost_reg_yn%>&end_yn=<%=end_yn%>&ck_sw=<%="y"%>">[다음]</a> 
                        <a href="tax_esero_mg.asp?page=<%=total_page%>&bill_id=<%=bill_id%>&bill_month=<%=bill_month%>&cost_reg_yn=<%=cost_reg_yn%>&end_yn=<%=end_yn%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnCenter">
				<% if (tottal_record > 0 and end_yn = "N") then	%>
					<a href="#" onClick="end_process()" class="btnType04">마감처리</a>
				<% end if	%>
				<% if (cost_record <> 0 and mi_record <> 0) then	%>
					<a href="#" onClick="end_process()" class="btnType04">부분마감처리</a>
				<% end if	%>
				<% if cost_record = 0 and end_yn = "Y" then	%>
					<a href="#" onClick="cancel_process()" class="btnType04">마감처리취소</a>
				<% end if	%>
					</div>                  
                    </td>
			      </tr>
				  </table>
				</form>
		</div>				
	</div>        				
	</body>
</html>

