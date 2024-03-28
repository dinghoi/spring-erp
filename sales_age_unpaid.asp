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
	field_check=Request("field_check")
	field_view=Request("field_view")
	view_sw=Request("view_sw")
  else
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	view_sw=Request.form("view_sw")
End if

If field_check = "" Then
	field_check = "total"
	view_sw = 0
End If

If field_check = "total" Then
	field_view = ""
End If
curr_date = mid(cstr(now()),1,10)
curr_month = mid(cstr(now(),1,7)

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select company,substring(collect_due_date,1,7) as collect_month, count(*) as uppaid_cnt,sum(sales_amt-collect_tot_amt) as sum_amt from saupbu_sales where (sales_amt > collect_tot_amt) "

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

order_sql = "group by company, substring(collect_due_date,1,7) ORDER BY company ASC"

Sql = "SELECT count(*) FROM saupbu_sales where (sales_amt > collect_tot_amt) " + field_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

sql = "select substring(collect_due_date,1,7) as collect_month, count(*) as uppaid_cnt,sum(sales_amt-collect_tot_amt) as sum_amt from saupbu_sales where (sales_amt > collect_tot_amt) " + field_sql + "group by substring(collect_due_date,1,7)"
Rs_sum.Open Sql, Dbconn, 1
do until rs_sum.eof
	rs_sum.movenext()
loop
rs_sum.close()

sql = base_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "미수금 관리"
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
				return "1 1";
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
			<!--#include virtual = "/include/sales_unpaid_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="sales_unpaid_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
							<label>
                               <select name="field_check" id="field_check" style="width:80px">
                           		<option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                <option value="company" <% if field_check = "company" then %>selected<% end if %>>거래처명</option>
                                <option value="emp_name" <% if field_check = "emp_name" then %>selected<% end if %>>영업담당</option>
                               </select>
							</label>
							<label>
								<strong>기준일자 : <%=curr_date%></strong>
                            </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="16%" >
							<col width="6%" >
							<col width="8%" >
							<col width="7%" >
							<col width="8%" >
							<col width="*" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">전표번호</th>
								<th scope="col">매출일자</th>
								<th scope="col">수금예정일</th>
								<th scope="col">미수금 예정일</th>
								<th scope="col">거래처명</th>
								<th scope="col">영업담당</th>
								<th scope="col">매출총액</th>
								<th scope="col">총수금액</th>
								<th scope="col">잔액</th>
								<th scope="col">미수금 사유</th>
								<th scope="col">수금</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><strong><%=formatnumber(total_record,0)%>건<strong></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(tot_sales_amt,0)%></td>
								<td class="right"><%=formatnumber(tot_collect_tot_amt,0)%></td>
								<td class="right"><%=formatnumber(tot_sales_amt - tot_collect_tot_amt,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
    					seq = total_record - ( page - 1 ) * pgsize
						do until rs.eof
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td><%=mid(rs("slip_no"),1,17)%></td>
								<td><%=rs("sales_date")%></td>
								<td><%=rs("collect_due_date")%></td>
								<td><%=rs("unpaid_due_date")%>&nbsp;</td>
								<td><%=rs("company")%></td>
								<td><%=rs("emp_name")%></td>
								<td class="right"><%=formatnumber(rs("sales_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("collect_tot_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_amt")-rs("collect_tot_amt"),0)%></td>
								<td><%=rs("unpaid_memo")%>&nbsp;</td>
							  	<td>
                                <a href="#" onClick="pop_Window('sales_collect_add.asp?slip_no=<%=rs("slip_no")%>&u_type=<%="U"%>','sales_collect_add_pop','scrollbars=yes,width=700,height=550')">등록</a>
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
                    <a href="sales_unpaid_excel.asp?field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="sales_unpaid_mg.asp?page=<%=first_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="sales_unpaid_mg.asp?page=<%=intstart -1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
               	  <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="sales_unpaid_mg.asp?page=<%=i%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
               	  <% if 	intend < total_page then %>
                        <a href="sales_unpaid_mg.asp?page=<%=intend+1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[다음]</a> 
                        <a href="sales_unpaid_mg.asp?page=<%=total_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[마지막]</a>
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

