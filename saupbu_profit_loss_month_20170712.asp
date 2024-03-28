<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim sum_amt(9)
dim tot_amt(9)
dim detail_tab(30)
dim cost_amt(30,9)
dim saupbu_tab(9)
dim sales_amt(9)
dim cost_tab

cost_tab = array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

i = 0
Sql="select saupbu from sales_org order by sort_seq asc"
rs_org.Open Sql, Dbconn, 1
do until rs_org.eof
	i = i + 1
	saupbu_tab(i) = rs_org(0)
	rs_org.movenext()
loop
rs_org.close()						
i = i + 1
'saupbu_tab(i) = ""
'i = i + 1
'saupbu_tab(i) = "소계"

cost_month=Request.form("cost_month")
if cost_month = "" then
	before_date = dateadd("m",-1,now())
	cost_month = mid(cstr(before_date),1,4) + mid(cstr(before_date),6,2)
end If

cost_year = mid(cost_month,1,4)
cost_mm = mid(cost_month,5)
c_month = cost_year + "-" + cost_mm
for i = 0 to 8
	sum_amt(i) = 0
	tot_amt(i) = 0
	sales_amt(i) = 0
next

sql = "select saupbu,sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&c_month&"' group by saupbu"
rs.Open Sql, Dbconn, 1
do until rs.eof
	bi_saupbu = rs("saupbu")
	if bi_saupbu = "기타사업부" then
		bi_saupbu = ""
	end if
	for i = 1 to 7
		if saupbu_tab(i) = bi_saupbu then
			sales_amt(i) = CCur(rs("sales_amt"))
			sales_amt(8) = sales_amt(8) + CCur(rs("sales_amt"))
			exit for
		end if
	next
	rs.movenext()
loop
rs.close()						

title_line = "사업부별 월별 손익 현황"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="saupbu_profit_loss_month.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>조회년월&nbsp;</strong>(예201401) : 
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
                    	<td>
      					<DIV id="topLine2" style="width:1200px;overflow:hidden;">
                <div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="70px" >
							<col width="170px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">비용항목</th>
							  <th rowspan="2" scope="col">세부내역</th>
						<% for i = 1 to 6	%>
							  <th scope="col"><%=saupbu_tab(i)%></th>
						<% next	%>
							  <th scope="col">기타사업부</th>
							  <th scope="col">소계</th>
                          </tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:470px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="70px" >
							<col width="170px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="*" >
						</colgroup>
						<tbody>
						<tr bgcolor="#FFFFCC">
							<td colspan="2" class="first" scope="col"><strong>매출</strong></td>
					<% for i = 1 to 8	%>				
                    		<td class="right" scope="col"><%=formatnumber(sales_amt(i),0)%></td>
 					<% next	%>
                         </tr>
					<%
					for jj = 0 to 10
						rec_cnt = 0

						for i = 1 to 30
							detail_tab(i) = ""
							for j = 1 to 8
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
							sql = "select cost_detail from saupbu_profit_loss where (cost_year ='"&cost_year&"') and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rs("cost_detail")
								rs.movenext()
							loop
							rs.close()
						end if
						if rec_cnt <> 0 then
' 당월 금액 SUM
							sql = "select saupbu,cost_detail,sum(cost_amt_"&cost_mm&") as cost from saupbu_profit_loss where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"' group by saupbu,cost_detail order by saupbu, cost_detail"
							rs.Open sql, Dbconn, 1
							do until rs.eof
								for i = 1 to 30
									if rs("cost_detail") = detail_tab(i) then
										for j = 1 to 7
											if saupbu_tab(j) = rs("saupbu") then
												cost_amt(i,j) = cost_amt(i,j) + Cdbl(rs("cost"))
												cost_amt(i,8) = cost_amt(i,8) + Cdbl(rs("cost"))
												sum_amt(j) = sum_amt(j) + Cdbl(rs("cost"))
												sum_amt(8) = sum_amt(8) + Cdbl(rs("cost"))
												tot_amt(j) = tot_amt(j) + Cdbl(rs("cost"))
												tot_amt(8) = tot_amt(8) + Cdbl(rs("cost"))
												exit for
											end if
										next
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
						<% for j = 1 to 8	%>
								<td class="right"><%=formatnumber(cost_amt(1,j),0)%></td>
						<% next	%>
						  </tr>
					  <% for i = 2 to rec_cnt	%>
                        	<tr>
								<td class="left" style=" border-left:1px solid #e3e3e3;"><%=detail_tab(i)%></td>
						<%   for j = 1 to 8	%>
								<td class="right"><%=formatnumber(cost_amt(i,j),0)%></td>
						<%   next	%>
							</tr>
						<% next	%>
							<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
						<% for j = 1 to 8	%>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(j),0)%></td>
						<% next	%>
						  </tr>
					<%
						end if
					next
					%>
					<tr bgcolor="#FFFFCC">
							  <td colspan="2" class="first" scope="col"><strong>비용합계</strong></td>
						<% for j = 1 to 8	%>
								<td class="right"><%=formatnumber(tot_amt(j),0)%></td>
						<% next	%>
                         </tr>
						<tr bgcolor="#FFDFDF">
							  <td colspan="2" bgcolor="#FFDFDF" class="first" scope="col"><strong>손익</strong></td>
						<%
						 for j = 1 to 8	
						 	cal_amt = sales_amt(j) - tot_amt(j)
						 %>
								<td class="right"><%=formatnumber(cal_amt,0)%></td>
						<%
						 next	
						 %>
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
                    <a href="saupbu_profit_loss_month_excel.asp?cost_month=<%=cost_month%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="profit_loss_detail_excel.asp?cost_month=<%=cost_month%>" class="btnType04">매입세금계산서다운로드</a>
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

