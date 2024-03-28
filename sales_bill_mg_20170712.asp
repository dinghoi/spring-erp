<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
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
		sales_month = request("sales_month")
		sales_saupbu = request("sales_saupbu")
		field_check = request("field_check")
		field_view = request("field_view")
	else
		sales_month = request.form("sales_month")
		sales_saupbu = request.form("sales_saupbu")
		field_check = request.form("field_check")
		field_view = request.form("field_view")
	end if
	
	if sales_month = "" then
		sales_month = mid(now(),1,4) + mid(now(),6,2)
		sales_saupbu = "전체"
		field_check = "total"
		field_view = ""
	end if

	if field_check = "total" then
		field_view = ""
	end if		
	
	sales_yymm = mid(sales_month,1,4) + "-" + mid(sales_month,5,2)
	
	pgsize = 10 ' 화면 한 페이지 
	
	If Page = "" Then
		Page = 1
		start_page = 1
	End If
	stpage = int((page - 1) * pgsize)

	base_sql = "select * from saupbu_sales where (substring(sales_date,1,7) = '"&sales_yymm&"')"

	if field_check = "total" then
		field_sql = " "
	  else
		field_sql = " and ("&field_check&" like '%"&field_view&"%') "
	end if
	if sales_saupbu = "전체" then
		saupbu_sql = " "
	  else
		saupbu_sql = " and (saupbu = '"&sales_saupbu&"') "
	end if
	
	order_sql = " ORDER BY sales_date ASC"

	sql = "select count(*) from saupbu_sales where (substring(sales_date,1,7) = '"&sales_yymm&"') " + field_sql + saupbu_sql
	Set RsCount = Dbconn.Execute (sql)
	
	tottal_record = cint(RsCount(0)) 'Result.RecordCount
	
	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF
	
	sql = "select sum(sales_amt) as price,sum(cost_amt) as cost,sum(vat_amt) as cost_vat from saupbu_sales where (substring(sales_date,1,7) = '"&sales_yymm&"') " + field_sql + saupbu_sql
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
	
	sql = base_sql + field_sql + saupbu_sql + order_sql + " limit "& stpage & "," &pgsize 
	Rs.Open Sql, Dbconn, 1

	title_line = "매출 업로드 내역 관리"
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
				return "2 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.sales_month.value == "") {
					alert ("매출년월을 선택하세요");
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
				<form action="sales_bill_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조회조건</dt>
                        <dd>
                            <p>
								<label>
								<strong>매출년월 : </strong>
                                	<input name="sales_month" type="text" value="<%=sales_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
                                <label>
								<strong>영업사업부</strong>
                                <select name="sales_saupbu" id="sales_saupbu" style="width:150px">
                                  <option value="전체" <% if sales_saupbu = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="회사간거래" <% if sales_saupbu = "회사간거래" then %>selected<% end if %>>회사간거래</option>
                                  <option value="기타사업부" <% if sales_saupbu = "기타사업부" then %>selected<% end if %>>기타사업부</option>
                                  <%
									Sql="select saupbu from sales_org order by sort_seq asc"
									rs_org.Open Sql, Dbconn, 1
									do until rs_org.eof
                                    %>
                                  <option value='<%=rs_org("saupbu")%>' <%If sales_saupbu = rs_org("saupbu") then %>selected<% end if %>><%=rs_org("saupbu")%></option>
                                  <%
                                        rs_org.movenext()
                                    loop
                                    rs_org.close()						
                                    %>
                                </select>
                                </label>
                                <label>
								<strong>세부조건</strong>
                                <select name="field_check" id="field_check" style="width:100px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                    <option value="sales_company" <% if field_check ="sales_company" then %>selected<% end if %>>매출회사</option>
                                    <option value="company" <% if field_check = "company" then %>selected<% end if %>>고객사</option>
                                    <option value="trade_no" <% if field_check = "trade_no" then %>selected<% end if %>>사업자번호</option>
                                    <option value="emp_name" <% if field_check = "emp_name" then %>selected<% end if %>>담당자</option>
                                    <option value="sales_memo" <% if field_check = "sales_memo" then %>selected<% end if %>>품목명</option>                                </select>
								</label>
                                <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:150px" id="field_view" >
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
							<col width="9%" >
							<col width="8%" >
							<col width="12%" >
							<col width="8%" >
							<col width="9%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="5%" >
							<col width="*" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">매출일자</th>
								<th scope="col">매출회사</th>
								<th scope="col">영업사업부</th>
								<th scope="col">고객사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">그룹</th>
								<th scope="col">합계금액</th>
								<th scope="col">공급가액</th>
								<th scope="col">세액</th>
								<th scope="col">담당자</th>
								<th scope="col">품목명</th>
								<th scope="col">변경</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><%=formatnumber(tottal_record,0)%>&nbsp;건</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(sum_price,0)%></td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
								<td class="right"><%=formatnumber(sum_cost_vat,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
						end_sw = "N"
						do until rs.eof
						%>
							<tr>
								<td class="first"><%=rs("sales_date")%></td>
								<td><%=rs("sales_company")%></td>
								<td><%=rs("saupbu")%></td>
								<td><%=rs("company")%></td>
								<td><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%></td>
								<td><%=rs("group_name")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("sales_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("vat_amt"),0)%></td>
								<td><%=rs("emp_name")%>&nbsp;</td>
								<td class="left"><%=rs("sales_memo")%></td>
								<td>
								<a href="#" onClick="pop_Window('sales_saupbu_mod.asp?approve_no=<%=rs("approve_no")%>','sales_saupbu_mod_pop','scrollbars=yes,width=800,height=250')">수정</a>
                                </td>
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
                    <a href="sales_report_excel.asp?sales_month=<%=sales_month%>&sales_saupbu=<%=sales_saupbu%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="sales_report.asp?page=<%=first_page%>&sales_month=<%=sales_month%>&sales_saupbu=<%=sales_saupbu%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="sales_report.asp?page=<%=intstart -1%>&sales_month=<%=sales_month%>&sales_saupbu=<%=sales_saupbu%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="sales_report.asp?page=<%=i%>&sales_month=<%=sales_month%>&sales_saupbu=<%=sales_saupbu%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="sales_report.asp?page=<%=intend+1%>&sales_month=<%=sales_month%>&sales_saupbu=<%=sales_saupbu%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[다음]</a> 
                        <a href="sales_report.asp?page=<%=total_page%>&sales_month=<%=sales_month%>&sales_saupbu=<%=sales_saupbu%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnCenter">
					</div>                  
                    </td>
			      </tr>
				  </table>
				</form>
		</div>				
	</div>        				
	</body>
</html>

