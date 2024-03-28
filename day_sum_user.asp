<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim company_tab(150)
dim acpt_cnt_tab(31,6)
dim acpt_date(30)
dim acpt_per(30)
dim per_cnt

for i = 0 to 30
	acpt_per(i) = 0
	for j = 0 to 6
		acpt_cnt_tab(i,j) = 0
	next
next

if c_grade = "5" then
	company = user_name
  else
  	company = reside_place
end if

to_date=Request.form("to_date")
company=Request.form("company")

If to_date = ""  Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid((cstr(dateadd("d",-30,now()))),1,10)
	company = "전체"
End If

from_date=cstr(dateadd("d",-30,to_date))

per_cnt = 500

for i = 0 to 30 
	acpt_date(i) = mid(cstr(dateadd("d",i,from_date)),1,10)
next

grade_sql = " and (company = '"+company+"') "

if reside = "9" then
	k = 0
	Sql="select * from trade where use_sw = 'Y' and group_name = '"+user_name+"' order by trade_name asc"
	rs_trade.Open Sql, Dbconn, 1
	do until rs_trade.eof
		k = k + 1
		company_tab(k) = rs_trade("trade_name")
		rs_trade.movenext()
	loop
	rs_trade.close()						
end if

if reside <> "9" then
	company = user_name
end if
if reside = "9" and company = "전체" then
	com_sql = "(company = '" + company_tab(1) + "'"	
	for kk = 2 to k
		com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
	next
	condi_sql = com_sql + ") "
  else
	condi_sql = " (company = '" + reside_company + "' or company = '" + company + "') "
end if

'총접수
sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and " + condi_sql
sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
Rs.Open Sql, Dbconn, 1
ii = 0
do until rs.eof
	for i = ii to 30
		if cstr(rs("com_date")) = acpt_date(i) then				
			acpt_cnt_tab(i,0) = cint(rs("acpt_cnt"))
	  		acpt_per(i) = acpt_cnt_tab(i,0) / per_cnt * 100
			exit for
		end if
	next
	ii = i
	rs.movenext()
loop
rs.close()

' 본사
sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (reside = '0') and " + condi_sql
sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		

Rs.Open Sql, Dbconn, 1
ii = 0
do until rs.eof
	for i = ii to 30
		if cstr(rs("com_date")) = acpt_date(i) then				
			acpt_cnt_tab(i,1) = cint(rs("acpt_cnt"))
			exit for
		end if
	next
	ii = i
	rs.movenext()
loop
rs.close()
			
' 상주
sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and reside = '1' and " + condi_sql
sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		

Rs.Open Sql, Dbconn, 1
ii = 0
do until rs.eof
	for i = ii to 30
		if cstr(rs("com_date")) = acpt_date(i) then				
			acpt_cnt_tab(i,2) = cint(rs("acpt_cnt"))
			exit for
		end if
	next
	ii = i
	rs.movenext()
loop
rs.close()

' 원격처리
sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type = '원격처리') and " + condi_sql
sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		

Rs.Open Sql, Dbconn, 1
ii = 0
do until rs.eof
	for i = ii to 30
		if cstr(rs("com_date")) = acpt_date(i) then				
			acpt_cnt_tab(i,4) = cint(rs("acpt_cnt"))
			exit for
		end if
	next
	ii = i
	rs.movenext()
loop
rs.close()
			
' 방문
sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리') and " + condi_sql
sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		

Rs.Open Sql, Dbconn, 1

ii = 0
do until rs.eof
	for i = ii to 30
		if cstr(rs("com_date")) = acpt_date(i) then				
			acpt_cnt_tab(i,5) = cint(rs("acpt_cnt"))
			exit for
		end if
	next
	ii = i
	rs.movenext()
loop
rs.close()

' 지방 방문 자체접수
sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리') and (company = '"+company+"')  and (acpt_man = mg_ce)"
sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		

Rs.Open Sql, Dbconn, 1

do until rs.eof
	for i = 0 to 30
		if cstr(rs("com_date")) = acpt_date(i) then				
			acpt_cnt_tab(i,6) = cint(rs("acpt_cnt"))
			exit for
		end if
	next
	rs.movenext()
loop
rs.close()

title_line = "일자별 접수 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}
		</script>
		<style type="text/css">
			.style_red {color: #FF0000; font-weight: bold}
			.style_green {color: #006600; font-weight: bold}
			.style_blue {color: #000099; font-weight: bold}
        </style>
	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/user_header.asp" -->
			<!--#include virtual = "/include/report_menu_user.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=day_sum_user.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" style="width:70px" value="<%=from_date%>" readonly="true">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker">
								</label>
 							<%
							if reside = "9" then
							%>
                                <label>
								<strong>회사</strong>
								<%
								sql_trade="select * from trade where (use_sw = 'Y') and (group_name = '"+user_name+"') order by trade_name asc"
                                rs_trade.Open sql_trade, Dbconn, 1
                                %>
                                <select name="company" id="company" style="width:150px">
 									<option value="전체">전체</option> 
          					<% 
								While not rs_trade.eof 
							%>
          							<option value='<%=rs_trade("trade_name")%>' <%If rs_trade("trade_name") = company  then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
          					<%
									rs_trade.movenext()  
								Wend 
								rs_trade.Close()
							%>
                                </select>
								</label>
							<%
							  else
							%>
								<strong>회사 : </strong><%=user_name%>
							<%
							end if
							%>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
						<% for i = 0 to 31 %>
							<col width="3%" >
                        <% next	%>
						</colgroup>
						<tbody>
							<tr valign="bottom">
                                <td class="first" height="200" valign="middle" style="background:#CFF"><strong>0<br>~<br><%=per_cnt%><br>기준</strong></td>
                  				<% 
								for i = 0 to 30 
									acpt_pro = int(acpt_per(i)*200/100)
								%>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro%>" align="center"></td>
							  <%
								next
								%>
                                <td>&nbsp;</td>
							</tr>
							<tr>
                                <th class="first">총괄</th>
                  				<% 
								sum_cnt = 0
								for i = 0 to 30 
									sum_cnt = sum_cnt + acpt_cnt_tab(i,0)
								%>
                                
                                <td><strong><%=acpt_cnt_tab(i,0)%></strong></td>
							  <%
								next
								%>
                                <td><strong><%=sum_cnt%></strong></td>
							</tr>
							<tr>
                                <td class="first" style="background:#FFC">본사</td>
                  				<% 
								sum_cnt = 0
								for i = 0 to 30 
								sum_cnt = sum_cnt + acpt_cnt_tab(i,1)
								%>

                                <td><%=acpt_cnt_tab(i,1)%></td>
							  <%
								next
								%>
                                <td><%=sum_cnt%></td>
							</tr>
							<tr>
                                <td class="first" style="background:#FFC">상주</td>
                  				<% 
								sum_cnt = 0
								for i = 0 to 30 
								sum_cnt = sum_cnt + acpt_cnt_tab(i,2)
								%>

                                <td><%=acpt_cnt_tab(i,1)%></td>
							  <%
								next
								%>
                                <td><%=sum_cnt%></td>
							</tr>
							<tr>
                                <td class="first" style="background:#6F9">자체</td>
                  				<% 
								sum_cnt = 0
								for i = 0 to 30 
								sum_cnt = sum_cnt + acpt_cnt_tab(i,6)
								%>

                                <td><%=acpt_cnt_tab(i,6)%></td>
						    <%
								next
								%>
                                <td><%=sum_cnt%></td>
							</tr>
							<tr>
                                <td class="first" style="background:#CFF">원격</td>
                  				<% 
								sum_cnt = 0
								for i = 0 to 30 
								sum_cnt = sum_cnt + acpt_cnt_tab(i,4)
								%>

                                <td><%=acpt_cnt_tab(i,4)%></td>
							  <%
								next
								%>
                                <td><%=sum_cnt%></td>
							</tr>
							<tr>
                                <td class="first" style="background:#CFF">방문</td>
                  				<% 
								sum_cnt = 0
								for i = 0 to 30 
								sum_cnt = sum_cnt + acpt_cnt_tab(i,5)
								%>

                                <td><%=acpt_cnt_tab(i,5)%></td>
							  <%
								next
								%>
                                <td><%=sum_cnt%></td>
							</tr>
							<tr>
                                <td class="first" style="background:#FFC"><strong>일자</strong></td>
                  				<% 
								for i = 0 to 30 
								%>
                                <td><strong><%=cstr(datepart("d",dateadd("d",i,from_date)))%></strong></td>
							  <%
								next
								%>
                                <td><strong>계</strong></td>
							</tr>
						</tbody>
					</table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

