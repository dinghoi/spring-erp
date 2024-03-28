<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim sum_amt(10)
dim tot_amt(10)
dim detail_tab(30)
dim cost_amt(30,10)
dim cost_tab
dim arr_company(1000)

cost_tab = array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","전산")

cost_month  = Request.form("cost_month")  ' 발생년월
view_sw     = Request.form("view_sw")     ' 1: 회사별, 2: 그룹별
company     = Request.form("company")     ' 회사명 (1: 회사별)
group_name  = Request.form("group_name")  ' 그룹명 (2: 그룹별)

if cost_month = "" then
	before_date = dateadd("m",-1,now())
	cost_month = mid(cstr(before_date),1,4) + mid(cstr(before_date),6,2)
	company = "선택"
	group_name = "선택"
	view_sw = "1"
end If

cost_year = mid(cost_month,1,4)
cost_mm   = mid(cost_month,5)
if cost_mm = "01" then
	before_year = cstr(int(cost_year) - 1)
	before_mm = "12"
  else
  	before_year = cost_year
	before_mm = right("0" + cstr(int(cost_mm) - 1),2)
end if
before_month = cstr(before_year) + cstr(before_mm)

c_month = cstr(cost_year) + "-" + cstr(cost_mm)
b_month = cstr(before_year) + "-" + cstr(before_mm)

if view_sw = "1" then
	condi_sql = " and company = '"&company&"'"
  else
  	condi_sql = " and group_name ='"&group_name&"'"
end if

for i = 0 to 10
	sum_amt(i) = 0
	tot_amt(i) = 0
next

sql = "select sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&b_month&"'" & condi_sql
Set rs_sum = Dbconn.Execute (sql)	
if isnull(rs_sum(0)) then
	before_sales_amt = 0 
  else
	before_sales_amt = clng(rs_sum(0)) 
end if

sql = "select sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&c_month&"'" & condi_sql
Set rs_sum = Dbconn.Execute (sql)	
if isnull(rs_sum(0)) then
	curr_sales_amt = 0 
  else
	curr_sales_amt = clng(rs_sum(0)) 
end if



  Sql="select company from company_profit_loss where (cost_year = '"&cost_year&"') group by company order by company asc"
  rs_org.Open Sql, Dbconn, 1
  arr_company_cnt = -1
  do until rs_org.eof
      arr_company_cnt = arr_company_cnt + 1
      arr_company(arr_company_cnt) = rs_org("company") 
      rs_org.movenext()
  loop
  rs_org.close()            

                                
title_line = "고객사별 손익 현황"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
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
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("조회년월을 입력하세요.");
					return false;
				}	
				return true;
			}
			function condi_view() {

				if (eval("document.frm.view_sw[0].checked")) {
					document.getElementById('group_name').value = '';
				}	
				if (eval("document.frm.view_sw[1].checked")) {
					document.getElementById('company').value = '';
				}	
			}
			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body onload="condi_view();">
		<div id="wrap">			
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="company_profit_loss_report.asp" method="post" name="frm">
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
							<label>
								<input type="radio" name="view_sw" value="1" <% if view_sw = "1" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="condi_view()" ><strong>회사별</strong>
								<input name="company" id="company" list="company_view" value="<% if view_sw = "1" then %><%=company%><% end if %>"  style="width:200px">
								<input type="radio" name="view_sw" value="2" <% if view_sw = "2" then %>checked<% end if %> style="width:30px" id="Radio4" onClick="condi_view()" ><strong>그룹별</strong>
								<input name="group_name"  id="group_name" list="group_name_view" value="<% if view_sw = "2" then %><%=group_name%><% end if %>">
							</label>
							<label>
							
							
                            <datalist  id="company_view" style="width:150px " >
                              <option value="선택" <% if company = "선택" then %>selected<% end if %>>선택</option>
                              <%
                              for i=0 to arr_company_cnt
                            	%><option value='<%=arr_company(i)%>' <%If company = arr_company(i) then %>selected<% end if %>><%=arr_company(i)%></option><%
                              next						
                              %>
                            </datalist>
                            <datalist  id="group_name_view" style="width:150px; ">
                              <option value="선택" <% if company = "선택" then %>selected<% end if %>>선택</option>
                              <%
                                Sql="select group_name from company_profit_loss where cost_year = '"&cost_year&"' and group_name <> '' group by group_name order by group_name asc"
                                rs_org.Open Sql, Dbconn, 1
                                do until rs_org.eof
                                	%>
                              		<option value='<%=rs_org("group_name")%>' <%If group_name = rs_org("group_name") then %>selected<% end if %>><%=rs_org("group_name")%></option>
                              		<%
                                    rs_org.movenext()
                                loop
                                rs_org.close()						
                              %>
                            </datalist>
							</label>
                                 <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<strong>* 기준 설정이 확정되지 않아 아직 확정된 손익이 아닙니다. 참조바랍니다!!</strong>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
                    	<td>
      					<DIV id="topLine2" style="width:1200px;overflow:hidden;">
						
						<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="*" >
							<col width="7%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="8%" >
							<col width="6%" >
							<col width="1%" >
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">비용항목</th>
							  <th rowspan="2" scope="col">세부내역</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">전 월</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">당 월</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">증감</th>
							  <th rowspan="2" scope="col"></th>
						  </tr>
							<tr>
							  <th scope="col" style="border-left:1px solid #e3e3e3;">상주직접비</th>
							  <th scope="col">직접비</th>
							  <th scope="col">전사공통비</th>
							  <th scope="col">부문공통비</th>
							  <th scope="col">계</th>
							  <th scope="col">상주직접비</th>
							  <th scope="col">직접비</th>
							  <th scope="col">전사공통비</th>
							  <th scope="col">부문공통비</th>
							  <th scope="col">계</th>
							  <th scope="col">금액</th>
							  <th scope="col">율</th>
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
							<col width="5%" >
							<col width="*" >
							<col width="7%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="8%" >
							<col width="6%" >
							<col width="1%" >
						</colgroup>
						<tbody>
						<tr bgcolor="#FFFFCC">
							  <td colspan="2" class="first" scope="col"><strong>매출계</strong></td>
							  <td colspan="5" scope="col"><strong><%=formatnumber(before_sales_amt,0)%></strong></td>
							  <td colspan="5" scope="col"><strong><%=formatnumber(curr_sales_amt,0)%></strong></td>
						<%
						   incr_amt = curr_sales_amt - before_sales_amt
						   if before_sales_amt = 0 and curr_sales_amt = 0 then
						   		incr_per = 0 
							  elseif before_sales_amt = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / before_sales_amt * 100
						   end if
						%>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
					<%
					for jj = 0 to 10
						rec_cnt = 0

						for i = 1 to 30
							detail_tab(i) = ""
							for j = 1 to 10
								cost_amt(i,j) = 0
								sum_amt(j) = 0
							next
						next
						if cost_tab(jj) = "인건비" then
							sql = "select cost_detail from saupbu_cost_account where cost_id ='"&cost_tab(jj)&"' order by view_seq"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rs("cost_detail")
								rs.movenext()
							loop
							rs.close()
						else
							sql = "select cost_detail from company_profit_loss where (cost_year ='"&cost_year&"' or cost_year ='"&before_year&"') and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rs("cost_detail")
								rs.movenext()
							loop
							rs.close()
						end if
						if rec_cnt <> 0 then
' 전월 금액 SUM
							sql = "select cost_center,cost_detail,sum(cost_amt_"&before_mm&") as cost from company_profit_loss where cost_year ='"&before_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_center,cost_detail order by cost_center, cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								for i = 1 to 30
									if rs("cost_detail") = detail_tab(i) then
										select case rs("cost_center")
											case "상주직접비"
												j = 1
											case "직접비"
												j = 2
											case "전사공통비"
												j = 3
											case "부문공통비"
												j = 4									
										end select
										cost_amt(i,j) = cost_amt(i,j) + clng(rs("cost"))
										cost_amt(i,5) = cost_amt(i,5) + clng(rs("cost"))
										sum_amt(j) = sum_amt(j) + clng(rs("cost"))
										sum_amt(5) = sum_amt(5) + clng(rs("cost"))
										tot_amt(j) = tot_amt(j) + clng(rs("cost"))
										tot_amt(5) = tot_amt(5) + clng(rs("cost"))
										exit for
									end if
								next								
								rs.movenext()
							loop
							rs.close()
' 당월 금액 SUM
							sql = "select cost_center,cost_detail,sum(cost_amt_"&cost_mm&") as cost from company_profit_loss where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_center,cost_detail order by cost_center, cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								for i = 1 to 30
									if rs("cost_detail") = detail_tab(i) then
										select case rs("cost_center")
											case "상주직접비"
												j = 6
											case "직접비"
												j = 7
											case "전사공통비"
												j = 8
											case "부문공통비"
												j = 9									
										end select
										cost_amt(i,j) = cost_amt(i,j) + clng(rs("cost"))
										cost_amt(i,10) = cost_amt(i,10) + clng(rs("cost"))
										sum_amt(j) = sum_amt(j) + clng(rs("cost"))
										sum_amt(10) = sum_amt(10) + clng(rs("cost"))
										tot_amt(j) = tot_amt(j) + clng(rs("cost"))
										tot_amt(10) = tot_amt(10) + clng(rs("cost"))
										exit for
									end if
								next								
								rs.movenext()
							loop
							rs.close()
						%>
							<tr>
							  	<td rowspan="<%=rec_cnt + 1%>" class="first">
						<% if jj = 2 or jj = 3 then	%>
                        	  	<%=cost_tab(jj)%><br>(현금사용)
						<%   else	%>
                        	  	<%=cost_tab(jj)%>
                        <% end if	%>
                              	</td>
								<td class="left"><%=detail_tab(1)%></td>
						<% for j = 1 to 10	%>
								<td class="right"><%=formatnumber(cost_amt(1,j),0)%></td>
						<% next	%>
						<%
						   incr_amt = cost_amt(1,10) - cost_amt(1,5)
						   if cost_amt(1,5) = 0 and cost_amt(1,10) = 0 then
						   		incr_per = 0 
							  elseif cost_amt(1,5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / cost_amt(1,5) * 100
						   end if
						%>
								<td class="right"><%=formatnumber(incr_amt,0)%></td>
				        		<td class="right"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right">&nbsp;</td>
							</tr>
						<% for i = 2 to rec_cnt	%>
                        	<tr>
								<td class="left" style=" border-left:1px solid #e3e3e3;"><%=detail_tab(i)%></td>
						<%   for j = 1 to 10	%>
								<td class="right"><%=formatnumber(cost_amt(i,j),0)%></td>
						<%   next	%>
						<%
						   incr_amt = cost_amt(i,10) - cost_amt(i,5)
						   if cost_amt(i,5) = 0 and cost_amt(i,10) = 0 then
						   		incr_per = 0 
							  elseif cost_amt(i,5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / cost_amt(i,5) * 100
						   end if
						%>
					      		<td class="right"><%=formatnumber(incr_amt,0)%></td>
								<td class="right"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right">&nbsp;</td>
							</tr>
						<% next	%>
							<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
						<% for j = 1 to 10	%>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(j),0)%></td>
						<% next	%>
						<%
						   incr_amt = sum_amt(10) - sum_amt(5)
						   if sum_amt(5) = 0 and sum_amt(10) = 0 then
						   		incr_per = 0 
							  elseif sum_amt(5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / sum_amt(5) * 100
						   end if
						%>
					      		<td class="right" bgcolor="#EEFFFF"><%=formatnumber(incr_amt,0)%></td>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right" bgcolor="#EEFFFF">&nbsp;</td>
							</tr>
					<%
						end if
					next
					%>
					<tr bgcolor="#FFFFCC">
							  <td colspan="2" class="first" scope="col"><strong>비용합계</strong></td>
						<% for j = 1 to 10	%>
							  <td scope="col" class="right"><%=formatnumber(tot_amt(j),0)%></td>
						<% next	%>
						<%
						   incr_amt = tot_amt(10) - tot_amt(5)
						   if tot_amt(5) = 0 and tot_amt(10) = 0 then
						   		incr_per = 0 
							  elseif tot_amt(5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / tot_amt(5) * 100
						   end if
						%>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
						<tr bgcolor="#FFDFDF">
							  <td colspan="2" class="first" scope="col"><strong>손익</strong></td>
						<%
						   	be_profit_loss = before_sales_amt - tot_amt(5)	
						   	curr_profit_loss = curr_sales_amt - tot_amt(10)	
						   	incr_amt = curr_profit_loss - be_profit_loss
						   	if be_profit_loss = 0 and curr_profit_loss = 0 then
						   		incr_per = 0 
							  elseif be_profit_loss = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / be_profit_loss * 100
						   	end if
							if be_profit_loss < 0 then
								incr_per = incr_per * -1
							end if 
						%>
							  <td scope="col" colspan="5" class="right"><strong><%=formatnumber(be_profit_loss,0)%></strong></td>
							  <td scope="col" colspan="5" class="right"><strong><%=formatnumber(curr_profit_loss,0)%></strong></td>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
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
                    <a href="company_profit_loss_excel.asp?cost_month=<%=cost_month%>&view_sw=<%=view_sw%>&company=<%=company%>&group_name=<%=group_name%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

