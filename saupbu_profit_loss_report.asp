<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	Dim sum_amt(10)
	Dim tot_amt(10)
	Dim detail_tab(30)
	Dim cost_amt(30,10)
	Dim cost_tab

	cost_tab = array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

	'cost_month=Request.form("cost_month")
	'sales_saupbu=Request.form("sales_saupbu")
	sales_saupbu=Request("sales_saupbu")

	'if sales_saupbu = "기타사업부" then
	'	sales_saupbu = ""
	'end if
	'if cost_month = "" then
	'	before_date = dateadd("m",-1,now())
	'	cost_month = mid(cstr(before_date),1,4) + mid(cstr(before_date),6,2)
	'	sales_saupbu = "전체"
	'end If

	'cost_year = mid(cost_month,1,4)
	'cost_mm = mid(cost_month,5)

	cost_year = request("cost_year")
	cost_mm = right("0" + cstr(request("cost_mm")),2)
	cost_month = cstr(cost_year) + cstr(cost_mm)

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

	if sales_saupbu = "전체" then
		condi_sql = ""
	else
		condi_sql = " AND saupbu ='"&sales_saupbu&"'"
	end if
	if sales_saupbu = "기타사업부" then
  		condi_sql = " AND (saupbu ='' OR saupbu = '기타사업부')"
	end if
	if (sales_saupbu = "한진" OR sales_saupbu = "한진그룹") then
		condi_sql = " AND saupbu IN ('한진', '한진그룹')"
	end if

	for i = 0 to 10
		sum_amt(i) = 0
		tot_amt(i) = 0
	next

	'/매출계(전월)
	sql = "SELECT SUM(cost_amt) AS sales_amt "&_
	      "  FROM saupbu_sales "&_
	      " WHERE SUBSTRING(SALES_DATE,1,7) = '"&b_month&"'"&condi_sql
	Set rs_sum = Dbconn.Execute (sql)

	if isnull(rs_sum(0)) then
		before_sales_amt = 0
  	else
		before_sales_amt = cdbl(rs_sum(0))
	end if

	'/매출계(당월)
	sql = "SELECT SUM(cost_amt) AS sales_amt "&_
	      "  FROM saupbu_sales "&_
	      " WHERE SUBSTRING(sales_date,1,7) = '"&c_month&"'"&condi_sql
	Set rs_sum = Dbconn.Execute (sql)

	if isnull(rs_sum(0)) then
		curr_sales_amt = 0
  	else
		curr_sales_amt = cdbl(rs_sum(0))
	end if

	title_line = sales_saupbu + " 손익 현황"

	if sales_saupbu = "" then
		title_line = "기타사업부 손익 현황"
	end if
	'Response.write sql &"<br>" & curr_sales_amt

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				return "2 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("조회년을 입력하세요.");
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
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td>
							<div id="topLine2" style="width:1200px;overflow:hidden;">
								<div class="gView">
									<table cellpadding="0" cellspacing="0" class="tableList">
										<colgroup>
											<col width="4%" >
											<col width="*" >
											<col width="8%" >
											<col width="6%" >
											<col width="6%" >
											<col width="7%" >
											<col width="9%" >
											<col width="8%" >
											<col width="6%" >
											<col width="6%" >
											<col width="7%" >
											<col width="9%" >
											<col width="7%" >
											<col width="5%" >
											<col width="1%" >
										</colgroup>
										<thead>
											<tr>
											  <th rowspan="2" class="first" scope="col">비용항목</th>
											  <th rowspan="2" scope="col">세부내역</th>
											  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">전 월&nbsp;(<%=before_year%>년<%=before_mm%>월)</th>
											  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">당 월&nbsp;(<%=cost_year%>년<%=cost_mm%>월)</th>
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
								</div>
							</div>
						</td>
					</tr>
					<tr>
          				<td valign="top">
				    	<DIV id="mainDisplay2" style="width:1200;height:470px;overflow:scroll" onscroll="scrollAll()">
				    	<table cellpadding="0" cellspacing="0" class="scrollList">
				    		<colgroup>
								<col width="6%" >
								<col width="*" >
								<col width="8%" >
								<col width="6%" >
								<col width="6%" >
								<col width="7%" >
								<col width="9%" >
								<col width="8%" >
								<col width="6%" >
								<col width="6%" >
								<col width="7%" >
								<col width="9%" >
								<col width="7%" >
								<col width="5%" >
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
								for jj = 0 to 10   ' cost_tab = array(0"인건비",1"야특근",2"일반경비",3"교통비",4"법인카드",5"임차료",6"외주비",7"자재",8"장비",9"운반비",10"상각비")

									rec_cnt = 0

									for i = 1 to 30
										detail_tab(i) = ""

										for j = 1 to 10
											cost_amt(i,j) = 0
											sum_amt(j) = 0
										next
									next

									'Response.write cost_tab(jj) & "<br>"
									if (cost_tab(jj) = "인건비")then
										sql = "   SELECT cost_detail "&_
											  "     FROM SAUPBU_COST_ACCOUNT "&_
											  "    WHERE cost_id = '인건비' "&_
											  " ORDER BY view_seq"
										rs.Open sql, Dbconn, 1
										'Response.write sql & "<br>"

										do until rs.eof
											rec_cnt = rec_cnt + 1
											detail_tab(rec_cnt) = rs("cost_detail")
											rs.movenext()
										loop
										rs.close()
									else
										sql = "   SELECT cost_detail "&_
											  "     FROM SAUPBU_PROFIT_LOSS "&_
											  "    WHERE (cost_year = '"& cost_year &"' OR cost_year = '"& before_year &"') "&_
											  "      AND cost_id ='"& cost_tab(jj) &"'"& condi_sql &_
											  " GROUP BY cost_detail "&_
											  " ORDER BY cost_detail"
										rs.Open sql, Dbconn, 1
										'Response.write sql & "<br>"

										do until rs.eof
											rec_cnt = rec_cnt + 1
											detail_tab(rec_cnt) = rs("cost_detail")
											rs.movenext()
										loop
										rs.close()
									end if

									if rec_cnt <> 0 then
										' 전월 금액 SUM
										sql = "  SELECT cost_center "&_
										  	  "       , cost_detail "&_
											  "       , SUM(cost_amt_"& before_mm &") AS cost " &_
											  "    FROM SAUPBU_PROFIT_LOSS "&_
											  "   WHERE cost_year = '"& before_year &"' "&_
											  "     AND cost_id   = '"& cost_tab(jj) &"'"&condi_sql &_
											  " GROUP BY cost_center, cost_detail "&_
											  " ORDER BY cost_center, cost_detail"
										'      if (cost_tab(jj) = "인건비")then
										'sql = sql & "   AND cost_id IN ('인건비', '야특근') "&condi_sql
										'    	else
										'sql = sql & "   AND cost_id ='"&cost_tab(jj)&"'"&condi_sql
										'    	end if
										'sql = sql & " GROUP BY cost_center, cost_detail "&_
										'      			" ORDER BY cost_center, cost_detail"
										rs.Open sql, Dbconn, 1
										'Response.write sql & ";<br>"
										
										do until rs.eof
											for i = 1 to 30
												'Response.write i & " : " & detail_tab(i) & "<br>" ' 전월에는 있지만 detail_tab에 없다면 cost_detail은 나오지 않는다..
												if rs("cost_detail") = detail_tab(i) then
													select case rs("cost_center")
														case "상주직접비" : j = 1
														case "직접비"     : j = 2
														case "전사공통비" : j = 3
														case "부문공통비" : j = 4
													end select
													
													cost_amt(i,j) = cost_amt(i,j) + Cdbl(rs("cost"))
													cost_amt(i,5) = cost_amt(i,5) + Cdbl(rs("cost"))
													sum_amt(j) = sum_amt(j) + Cdbl(rs("cost"))
													sum_amt(5) = sum_amt(5) + Cdbl(rs("cost"))
													tot_amt(j) = tot_amt(j) + Cdbl(rs("cost"))
													tot_amt(5) = tot_amt(5) + Cdbl(rs("cost"))
													
													exit for
												end if
											next
											rs.movenext()
										loop
										rs.close()
													
										' 당월 금액 SUM
										sql = "    SELECT cost_center "&_
											  "         , cost_detail "&_
											  "         , SUM(cost_amt_"&cost_mm&") AS cost "&_
											  "      FROM  SAUPBU_PROFIT_LOSS "&_
											  "     WHERE  cost_year ='"& cost_year &"' "&_
											  "       AND  cost_id   ='"& cost_tab(jj) &"'"&condi_sql&" "&_
											  " GROUP  BY cost_center, cost_detail "&_
											  " ORDER  BY cost_center, cost_detail"
										rs.Open sql, Dbconn, 1
										'Response.write sql & ";<br>"
										
										do until rs.eof
											for i = 1 to 30
												'Response.write i & " : " & detail_tab(i) & "<br>" ' 전월에는 있지만 detail_tab에 없다면 cost_detail은 나오지 않는다..
												if rs("cost_detail") = detail_tab(i) then
													select case rs("cost_center")
														case "상주직접비"	: j = 6
														case "직접비"	    : j = 7
														case "전사공통비"	: j = 8
														case "부문공통비"	: j = 9
													end select
													
													cost_amt(i,j) = cost_amt(i,j) + Cdbl(rs("cost"))
													cost_amt(i,10) = cost_amt(i,10) + Cdbl(rs("cost"))
													sum_amt(j) = sum_amt(j) + Cdbl(rs("cost"))
													sum_amt(10) = sum_amt(10) + Cdbl(rs("cost"))
													tot_amt(j) = tot_amt(j) + Cdbl(rs("cost"))
													tot_amt(10) = tot_amt(10) + Cdbl(rs("cost"))

													exit for
												end if
											next
											rs.movenext()
										loop
										rs.close()
										%>
										<tr>
							  				<td rowspan="<%=rec_cnt + 1%>" class="first">
												<%
												if jj = 2 or jj = 3 then
													Response.write cost_tab(jj) & "<br>(현금사용)"
												else
	                  								Response.write cost_tab(jj)
	                  							end if
	                  							%>
                  							</td>
											<td class="left"><%=detail_tab(1)%></td>

											<%
											for j = 1 to 10
												if j = 5 or j = 10 then
													Response.write "<td class='right'><strong>"&formatnumber(cost_amt(1,j),0)&"</strong></td>"
												else
													Response.write "<td class='right'>" ' [["&jj&"]][[cost_amt(1,"&j&")="&cost_amt(1,j)&"]] 
													if jj < 2	then
														Response.write formatnumber(cost_amt(1,j),0)
													else
														if (j = 1 or j = 2 or j = 6 or j = 7) and (jj > 1) and (cost_amt(1,j) <> 0)	then
														%>
			                  								<a href="#" onClick="pop_Window('profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(1)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')"><%=formatnumber(cost_amt(1,j),0)%></a>
														<%	  
														else
			                  								Response.write formatnumber(cost_amt(1,j),0)
			                  							end if
			                  						end if
			                  						%>
		                  							</td>
												<%   
												end if	
											next	

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
											<%
											for j = 1 to 10
												if j = 5 or j = 10 then
													Response.write "<td class='right'><strong>"&formatnumber(cost_amt(i,j),0)&"</strong></td>"
												else
													%>
													<td class="right">
														<%	if jj < 2	then	'//2016-08-23 알바비 상세조회 링크 추가
															If detail_tab(i)="알바비" Then
															%>
																<a href="#" onClick="pop_Window('profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(i)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')"><%=formatnumber(cost_amt(i,j),0)%></a>
															<%
															Else
																Response.write formatnumber(cost_amt(i,j),0)
															End IF
															%>
														<%	else	%>
															<%		
															if (j = 1 or j = 2 or j = 6 or j = 7) and jj > 1 and (cost_amt(i,j) <> 0) then	%>
																<a href="#" onClick="pop_Window('profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(i)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')"><%=formatnumber(cost_amt(i,j),0)%></a>
															<%
															else	
															%>
																<%=formatnumber(cost_amt(i,j),0)%>
															<%		
															end if	%>
														<%	end if	%>
													</td>
													<%
												end if
											next
											
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
											<%   if j = 5 or j = 10 then	%>
											<td class="right" bgcolor="#EEFFFF"><strong><%=formatnumber(sum_amt(j),0)%></strong></td>
											<% 	 else	%>
											<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(j),0)%></td>
											<%   end if	%>
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
									<%   if j = 5 or j = 10 then	%>
									<td class="right"><strong><%=formatnumber(tot_amt(j),0)%></strong></td>
									<% 	 else	%>
									<td class="right"><%=formatnumber(tot_amt(j),0)%></td>
									<%   end if	%>
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
									<td colspan="2" bgcolor="#FFDFDF" class="first" scope="col"><strong>손익</strong></td>
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
									<td scope="col" colspan="5"><strong><%=formatnumber(be_profit_loss,0)%></strong></td>
									<td scope="col" colspan="5"><strong><%=formatnumber(curr_profit_loss,0)%></strong></td>
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
				    	<td>
							<div class="btnCenter">
			            		<a href="saupbu_profit_loss_excel.asp?cost_year=<%=cost_year%>&cost_mm=<%=cost_mm%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">화면 엑셀다운로드</a>
			            		<a href="cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">상주비/직접비 엑셀다운로드</a>
			            		<a href="saupbu_sales_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">매출액 엑셀다운로드</a>
								<% if sales_grade = "0" then	%>
			            			<a href="cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%="전사공통비"%>" class="btnType04">전사공통비 엑셀다운로드</a>
			          				<a href="cost_center_detail_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%="부문공통비"%>" class="btnType04">부문공통비 엑셀다운로드</a>
								<% end if	%>
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
