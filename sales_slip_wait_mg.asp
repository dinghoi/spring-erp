<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim as_process
Dim field_check
Dim field_view
Dim win_sw

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
	slip_id = "1"
	view_date = "total"
End If

If view_date = "total" Then
	to_date = ""
	from_date = ""
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

base_sql = "select * from sales_slip "

if view_date = "total" then
	date_sql = "where (sign_yn = 'Y') "
  else
  	date_sql = "where (sign_yn = 'Y') and (" + view_date + " >='" + from_date  + "' and " + view_date + " <= '" + to_date  + "') "
end if

'if view_date = "reg_date" then
'	date_sql = "where (sign_yn = 'Y') and (CAST(reg_date as date) >= '"&from_date&"' and CAST(reg_date as date) <= '"&to_date&"') "
'  else
'  	date_sql = "where (sign_yn = 'Y') and ("&view_date&" >='"&from_date&"' and "&view_date&" <= '"&to_date&"') "
'end if

slip_sql = " and slip_id = '1' "

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

order_sql = " ORDER BY slip_no DESC"

Sql = "SELECT count(*) FROM sales_slip " + date_sql + slip_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

sql = base_sql + date_sql + slip_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
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
				<form action="sales_slip_wait_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
								<strong>검색일자 : </strong>
                                <select name="view_date" id="view_date" style="width:150px">
                                    <option value="total" <% if view_date = "total" then %>selected<% end if %>>전체</option>
                                    <option value="sales_date" <% if view_date = "sales_date" then %>selected<% end if %>>매출일자</option>
                                    <option value="bill_issue_date" <% if view_date = "bill_issue_date" then %>selected<% end if %>>계산서발행일</option>
                                    <option value="bill_due_date" <% if view_date = "bill_due_date" then %>selected<% end if %>>계산서발행일예정일</option>
                                    <option value="out_request_date" <% if view_date = "out_request_date" then %>selected<% end if %>>출고요청일</option>
                                    <option value="collect_due_date" <% if view_date = "collect_due_date" then %>selected<% end if %>>수금예정일</option>
                                    <option value="collect_date" <% if view_date = "collect_date" then %>selected<% end if %>>수금완료일</option>
                                </select>
								<label>
								<strong>일자 : </strong>
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
							<col width="10%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="4%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">전표번호</th>
								<th scope="col">전표상태</th>
								<th scope="col">매출일자</th>
								<th scope="col">계산서<br>발행일</th>
								<th scope="col">계산서<br>발행예정일</th>
								<th scope="col">거래처명</th>
								<th scope="col">영업담당</th>
								<th scope="col">매출구분</th>
								<th scope="col">매입총액</th>
								<th scope="col">매출총액</th>
								<th scope="col">마진총액</th>
								<th scope="col">취소</th>
								<th scope="col">수주전표</th>
							</tr>
						</thead>
						<tbody>
						<%
    					seq = total_record - ( page - 1 ) * pgsize
						do until rs.eof
							Sql="select slip_no from sales_slip where slip_no = '"&rs("slip_no")&"' and slip_id = '2' and slip_seq = '"&rs("slip_seq")&"'"
							Set rs_etc=DbConn.Execute(Sql)
							if rs_etc.eof or rs_etc.bof then
								exchange_yn = "N"
							  else
							  	exchange_yn = "Y"
							end if
							rs_etc.close()
							
							if rs("sales_yn") = "Y" then
								sales_yn_view = "매출"
							  else
								sales_yn_view = "비매출"
							end if
							
							if rs("slip_stat") = "2" then	
								slip_stat_view = "부분수주"
							  elseif rs("slip_stat") = "3" then	
								slip_stat_view = "수주완료"
							  elseif rs("slip_stat") = "4" then	
								slip_stat_view = "취소진행"
							  elseif rs("slip_stat") = "5" then	
								slip_stat_view = "전표취소"
							  else
								slip_stat_view = "단독전표"
							end if
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td>
								<a href="#" onClick="pop_Window('sales_slip_view.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>&cancel_yn=<%="N"%>','sales_slip_view_pop','scrollbars=yes,width=1230,height=600')"><%=rs("slip_no")%>-<%=rs("slip_seq")%></a>&nbsp;
                        <% if rs("slip_stat") = "2" or rs("slip_stat") = "3" then	%>
                                <a href="#" onClick="pop_Window('sales_slip_history.asp?slip_no=<%=rs("slip_no")%>','sales_slip_history_pop','scrollbars=yes,width=1230,height=600')"><img src="image/hist.gif" width="24" height="11" border="0"></a>
						<% end if	%>
                                </td>
							  	<td><%=slip_stat_view%></td>
								<td><%=rs("sales_date")%>&nbsp;</td>
								<td>
						<% if rs("sales_yn") = "N" or rs("collect_stat") = "영수" then	%>
								미발행
                       	<%   else	%>
						<% 		if rs("bill_issue_date") = "0000-00-00" or isnull(rs("bill_issue_date")) then	%>
                                <a href="#" onClick="pop_Window('sales_bill_issue_date.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>','sales_bill_issue_date_pop','scrollbars=yes,width=400,height=300')">등록</a>
                        <%   	  else	%>
                        		<%=rs("bill_issue_date")%>&nbsp;
               			<% 		end if	%>
						<% end if	%>
                                </td>
								<td><%=rs("bill_due_date")%>&nbsp;</td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=sales_yn_view%></td>
								<td class="right"><%=formatnumber(rs("buy_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("margin_cost"),0)%></td>
								<td>
						<% if rs("slip_stat") = "1" then	%>
                                <a href="#" onClick="pop_Window('sales_slip_view.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>&cancel_yn=<%="Y"%>','sales_slip_view_pop','scrollbars=yes,width=1230,height=650')">취소</a>
						<%   else	%>
								불가
                        <% end if	%>
                                </td>
								<td>
						<% if rs("slip_stat") < "3" then	%>
                                <a href="#" onClick="pop_Window('sales_slip_exchange.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','sales_slip_exchange_pop','scrollbars=yes,width=1230,height=650')">수주등록</a>
						<%   else	%>
                        		-
              			<% end if	%>
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
                    <a href="sales_slip_excel.asp?sign_yn=<%="Y"%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="sales_slip_wait_mg.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="sales_slip_wait_mg.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="sales_slip_wait_mg.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
        <% if 	intend < total_page then %>
                        <a href="sales_slip_wait_mg.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[다음]</a> 
                        <a href="sales_slip_wait_mg.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
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

