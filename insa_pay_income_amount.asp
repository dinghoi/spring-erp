<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	inc_yyyy=request("inc_yyyy")
  else
	inc_yyyy=Request.form("inc_yyyy")
End if

If inc_yyyy = "" Then
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	inc_yyyy = mid(cstr(from_date),1,4)
End If

' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

' 분기 테이블 생성
curr_mm = mid(now(),6,2)
if curr_mm > 0 and curr_mm < 4 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "1"
end if
if curr_mm > 3 and curr_mm < 7 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "2"
end if
if curr_mm > 6 and curr_mm < 10 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "3"
end if
if curr_mm > 9 and curr_mm < 13 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "4"
end if

quarter_tab(8,2) = cstr(mid(quarter_tab(8,1),1,4)) + "년 " + cstr(mid(quarter_tab(8,1),5,1)) + "/4분기"

for i = 7 to 1 step -1
	cal_quarter = cint(quarter_tab(i+1,1)) - 1
	if cstr(mid(cal_quarter,5,1)) = "0" then
		quarter_tab(i,1) = cstr(cint(mid(cal_quarter,1,4))-1) + "4"
	  else
		quarter_tab(i,1) = cal_quarter
	end if	 
	quarter_tab(i,2) = cstr(mid(quarter_tab(i,1),1,4)) + "년 " + cstr(mid(quarter_tab(i,1),5,1)) + "/4분기"
next

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_into = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select count(*) from pay_income_amount where (inc_yyyy = '"+inc_yyyy+"')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "select * from pay_income_amount where inc_yyyy = '"+inc_yyyy+"' ORDER BY inc_yyyy,inc_seq ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "근로소득 간이세액표"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
			}
		</script>
		<script type="text/javascript">
		    $(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.inc_yyyy.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_rule_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_income_amount.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>년도 검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>귀속년도 : </strong>
                                    <select name="inc_yyyy" id="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:90px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If inc_yyyy = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                    </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset> 
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
                            <col width="3%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
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
                            <col width="6%" >
                            <col width="3%" >
						</colgroup>
						<thead>
				            <tr>
				               <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">년도</th>
				               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">순번</th>
                               <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">월 급여액</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">과세표준액</th>
				               <th colspan="11" scope="col" style=" border-bottom:1px solid #e3e3e3;">공제 가족수에 준한 간이세액</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">변경</th>
			               </tr>
                           <tr>
				              <th scope="col" style=" border-left:1px solid #e3e3e3;">이상</th>
				              <th scope="col" style=" border-bottom:1px solid #e3e3e3;">미만</th>
                              <th scope="col">1인</th>
				              <th scope="col">2인</th>
				              <th scope="col">3인</th>
                              <th scope="col">4인</th>
                              <th scope="col">5인</th>
                              <th scope="col">6인</th>
                              <th scope="col">7인</th>
                              <th scope="col">8인</th>
                              <th scope="col">9인</th>
                              <th scope="col">10인</th>
                              <th scope="col">11인</th>
                           </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
						   if rs("inc_seq") = "998" then
					          inc_comment1 = rs("inc_comment")
						   end if
						   if rs("inc_seq") = "999" then
					          inc_comment2 = rs("inc_comment")
						   end if
						   
						%>
							<tr>
								<td class="first"><%=rs("inc_yyyy")%>&nbsp;</td>
								<td ><%=rs("inc_seq")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("inc_from_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_to_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_st_amt"),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("inc_incom1"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom2"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom3"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom4"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom5"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom6"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom7"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom8"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom9"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom10"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("inc_incom11"),0)%>&nbsp;</td>
								<td>
                                 <a href="#" onClick="pop_Window('insa_pay_income_amount_add.asp?inc_yyyy=<%=rs("inc_yyyy")%>&inc_seq=<%=rs("inc_seq")%>&u_type=<%="U"%>','pay_income_amount_popup','scrollbars=yes,width=750,height=400')"></a>
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()   
						%>
                           <tr>
                                <th colspan="3" class="first">10,000천원 초과~14,000천원 이하</th>
								<td class="left" colspan="14"><%=inc_comment1%>&nbsp;</td>
                           </tr>
                           <tr>
                                <th colspan="3" class="first">14,000천원 초과</th>
								<td class="left" colspan="14"><%=inc_comment2%>&nbsp;</td>
                           </tr>
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
				    <td>
                    <div id="paging">
                        <a href="insa_pay_income_amount.asp?page=<%=first_page%>&inc_yyyy=<%=inc_yyyy%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_income_amount.asp?page=<%=intstart -1%>&inc_yyyy=<%=inc_yyyy%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_income_amount.asp?page=<%=i%>&inc_yyyy=<%=inc_yyyy%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_pay_income_amount.asp?page=<%=intend+1%>&inc_yyyy=<%=inc_yyyy%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_income_amount.asp?page=<%=total_page%>&inc_yyyy=<%=inc_yyyy%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td> 
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

