<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	field_check=Request("field_check")
	field_view=Request("field_view")
	page_cnt=Request("page_cnt")

Else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	page_cnt=Request.form("page_cnt")
End if

If to_date = "" or from_date = "" Then
'	curr_dd = cstr(datepart("d",now))
'	from_date = mid(cstr(now()-curr_dd+1),1,10)
'	from_date = cstr(dateadd("m",-1,from_date))
'	to_date = dateadd("m",1,from_date)
'	to_date = cstr(dateadd("d",-1,to_date))
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	field_check = "total"
End If
If field_check = "total" Then
	field_view = ""
End If

bill_id = "매입"

pgsize = 10 ' 화면 한 페이지 
'pgsize = page_cnt ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 조건별 조회.........
if field_check = "total" then
	sql = "select * from general_cost where tax_bill_yn = 'Y' and (slip_date >= '" + from_date  + "' and slip_date <= '" + to_date  + "') ORDER BY customer, slip_gubun, slip_date ASC"
  else
	sql = "select * from general_cost where tax_bill_yn = 'Y' and (slip_date >= '" + from_date  + "' and slip_date <= '" + to_date  + "') and (" + field_check + " like '%" + field_view + "%' ) ORDER BY customer, slip_gubun, slip_date ASC"
end if
Rs.Open Sql, Dbconn, 1
'Response.write sql&"<br>"

title_line = "매입세금계산서 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
//				if (document.frm.bill_id.value == "") {
//					alert ("필드조건을 선택하시기 바랍니다");
//					return false;
//				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/tax_bill_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="tax_bill_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<strong>조건조회</strong>
                                <select name="field_check" id="field_check" style="width:120px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                    <option value="customer" <% if field_check = "customer" then %>selected<% end if %>>거래처</option>
                                    <option value="account" <% if field_check = "account" then %>selected<% end if %>>유형</option>
                                    <option value="org_name" <% if field_check = "org_name" then %>selected<% end if %>>발행부서</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:120px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="12%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
							<col width="*" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">거래처</th>
								<th scope="col">사업자번호</th>
								<th scope="col">발행일</th>
								<th scope="col">유형</th>
								<th scope="col">항목</th>
								<th scope="col">발행내역</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">발행회사</th>
                                <th scope="col">발행부서</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							i = i + 1
							if rs("end_yn") = "Y" then
								end_yn = "마감"
							  else
							  	end_yn = "진행"
							end if
							customer_no = mid(rs("customer_no"),1,3) + "-" + mid(rs("customer_no"),4,2) + "-" + right(rs("customer_no"),5)
						%>
							<tr>
								<td class="first"><%=rs("customer")%></td>
								<td><%=customer_no%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("slip_gubun")%></td>
								<td><%=rs("account_item")%></td>
								<td><%=rs("slip_memo")%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("emp_company")%></td>
								<td><%=rs("org_name")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr>
								<th class="first">총계</th>
								<th colspan="1"><%=i%>&nbsp;건</th>
							  	<th colspan="3">합계 :&nbsp;<%=formatnumber(price_sum,0)%></th>
							  	<th colspan="3">공급가액 :&nbsp;<%=formatnumber(cost_sum,0)%></th>
								<th colspan="3">부가세 :&nbsp;<%=formatnumber(cost_vat_sum,0)%></th>
							</tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="tax_bill_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&bill_id=<%=bill_id%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
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

