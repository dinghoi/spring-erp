<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim as_process
Dim field_check
Dim field_view
Dim win_sw
dim company_tab(160)

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	slip_id=Request("slip_id")
	view_date=Request("view_date")
	field_check=Request("field_check")
	field_view=Request("field_view")
  else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	slip_id=Request.form("slip_id")
	view_date=Request.form("view_date")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If to_date = "" or from_date = "" Then
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-31),1,10)
	field_check = "total"
	slip_id = "T"
	view_date = "sales_date"
End If

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select * from sales_slip "

date_sql = "where (sign_id = '1') and (sign_yn = 'Y') and (" + view_date + " >='" + from_date  + "' and " + view_date + " <= '" + to_date  + "') "

if slip_id = "T" then
	slip_sql = " "
  else
	slip_sql = " and slip_id = '"+ slip_id + "'"
end if

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

order_sql = " ORDER BY sales_date DESC"

Sql = "SELECT count(*) FROM sales_slip " + date_sql + slip_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

sql = base_sql + date_sql + slip_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
'response.write(sql)
Rs.Open Sql, Dbconn, 1

title_line = "대기 전표 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
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
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.field_check.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/sales_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="sales_slip_ing_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
								<strong>전표구분 : </strong>
                                <select name="slip_id" id="slip_id" style="width:80px">
                              		<option value="T" <% if slip_id = "T" then %>selected<% end if %>>전체</option>
                                    <option value="2" <% if slip_id = "2" then %>selected<% end if %>>수주전표</option>
                                    <option value="1" <% if slip_id = "1" then %>selected<% end if %>>대기전표</option>
                                </select>
								<strong>일자별 검색 : </strong>
                                <select name="view_date" id="view_date" style="width:150px">
                                    <option value="sales_date" <% if view_date = "sales_date" then %>selected<% end if %>>매출일자</option>
                                    <option value="bill_issue_date" <% if view_date = "bill_issue_date" then %>selected<% end if %>>계산서발행일</option>
                                    <option value="bill_due_date" <% if view_date = "bill_due_date" then %>selected<% end if %>>계산서발행일예정일</option>
                                    <option value="out_request_date" <% if view_date = "out_request_date" then %>selected<% end if %>>출고요청일</option>
                                    <option value="collect_due_date" <% if view_date = "collect_due_date" then %>selected<% end if %>>수금예정일</option>
                                    <option value="collect_date" <% if view_date = "collect_date" then %>selected<% end if %>>수금완료일</option>
                                </select>
								<label>
								일자
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								~
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <label>
								<strong>조건 : </strong>
                                <select name="field_check" id="field_check" style="width:80px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                    <option value="slip_no" <% if field_check = "slip_no" then %>selected<% end if %>>전표번호</option>
                                    <option value="trade_name" <% if field_check = "trade_name" then %>selected<% end if %>>거래처명</option>
                                    <option value="emp_name" <% if field_check = "emp_name" then %>selected<% end if %>>영업담당</option>
                                </select>
								</label>
                                <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:100px" id="field_view" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="5%" >
							<col width="6%" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="*" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">전표구분</th>
								<th scope="col">결재요청</th>
								<th scope="col">전표번호</th>
								<th scope="col">매출일자</th>
								<th scope="col">계산서<br>발행일</th>
								<th scope="col">계산서<br>발행예정일</th>
								<th scope="col">거래처명</th>
								<th scope="col">영업담당</th>
								<th scope="col">매입총액</th>
								<th scope="col">매출총액</th>
								<th scope="col">마진총액</th>
								<th scope="col">수금방법</th>
								<th scope="col">수금완료일</th>
								<th scope="col">수정</th>
							</tr>
						</thead>
						<tbody>
						<%
    					seq = total_record - ( page - 1 ) * pgsize
						do until rs.eof
							if rs("slip_id") = "2" then
								slip_id_view = "수주전표"
							end if
							if rs("slip_id") = "1" then
								slip_id_view = "대기전표"
							end if
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td><%=slip_id_view%></td>
								<td>결제요청</td>
								<td><%=rs("slip_no")%>-<%=rs("slip_seq")%></td>
								<td><%=rs("sales_date")%></td>
								<td><%=rs("bill_issue_date")%></td>
								<td><%=rs("bill_due_date")%></td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("emp_name")%></td>
								<td class="right"><%=formatnumber(rs("buy_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("margin_cost"),0)%></td>
								<td><%=rs("bill_collect")%></td>
								<td><%=rs("collect_date")%>&nbsp;</td>
							  	<td>
                                <a href="#" onClick="pop_Window('sales_slip_wait_mod.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','sales_slip_wait_mod_pop','scrollbars=yes,width=1230,height=600')">수정</a>
                                </td>
							</tr>
						<%
							rs.movenext()
  							seq = seq -1
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
				    <td width="15%">
					<div class="btnCenter">
                    <a class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="sales_slip_ing_mg.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="sales_slip_ing_mg.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="sales_slip_ing_mg.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="sales_slip_ing_mg.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[다음]</a> 
                        <a href="sales_slip_ing_mg.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
					<a href="#" onClick="pop_Window('sales_slip_wait_add.asp','sales_slip_wait_add_pop','scrollbars=yes,width=1230,height=600')" class="btnType04">대기전표등록</a>
					<a href="#" onClick="pop_Window('sales_slip_order_add.asp','sales_slip_order_add_pop','scrollbars=yes,width=1230,height=600')" class="btnType04">수주전표등록</a>
                    </div>                  
                    </td>
			      </tr>
				  </table>
				<input type="hidden" name="user_id">
				<input type="hidden" name="pass">
			</form>
		</div>				
	</div>        				
	</body>
</html>

