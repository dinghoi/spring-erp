<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
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

c_name = "전체"

to_date=Request.form("to_date")
'from_date=Request.form("from_date")
company=Request.form("company")

'If to_date = "" or from_date = "" Then
If to_date = ""  Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid((cstr(dateadd("d",-30,now()))),1,10)
	company = "전체"
	acpt_place = "총괄"
End If

from_date=cstr(dateadd("d",-30,to_date))

if	c_grade = "0" or c_grade = "1" or c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
	c_name = "전체"
end if
per_cnt = 1000
if c_grade = "7" then 
	per_cnt = 400
end if
if c_grade = "8" or c_grade = "5" then 
	per_cnt = 300
end if

if	c_grade = "8" or ( c_grade = "5" and c_reside = "0" ) then
	c_name = request.cookies("asmg_user")("coo_name")
	company = c_name
end if

for i = 0 to 30 
	acpt_date(i) = mid(cstr(dateadd("d",i,from_date)),1,10)
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if c_name = "전체" then
	k = 0
	company_tab(0) = "전체"
	if	c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
		Sql="select * from etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' and group_name = '"+user_name+"' order by etc_name asc"
		  else
		Sql="select * from etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' order by etc_name asc"
	end if
	Rs_etc.Open Sql, Dbconn, 1
	while not rs_etc.eof
		k = k + 1
		company_tab(k) = rs_etc("etc_name")
		rs_etc.movenext()
	Wend
rs_etc.close()						
end if				

grade_sql = ""
if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
	com_sql = "company = '" + company_tab(1) + "'"	
	for kk = 1 to k
		com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
	next
	grade_sql = " and (" + com_sql + ")"
end if
kkk = k

'총접수
if company = "전체" then
	if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "+ grade_sql
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
  else
	if c_grade = "7" or c_grade = "8" then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (company = '"+company+"') "
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (company = '"+company+"') "
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
end if
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
if company = "전체" then
	if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (reside_place = '본사')"+ grade_sql
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (reside_place = '본사')"
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
  else
	if c_grade = "7" or c_grade = "8" then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and reside_place = '본사' and (company = '"+company+"') "
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and reside_place = '본사' and (company = '"+company+"') "
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
end if


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
if company = "전체" then
	if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (reside_place <> '본사')"+ grade_sql
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (reside_place <> '본사')"
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
  else
	if c_grade = "7" or c_grade = "8" then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and reside_place <> '본사' and (company = '"+company+"') "
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and reside_place <> '본사' and (company = '"+company+"') "
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
end if


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

' 인터넷 접수
'if company = "전체" then
'	if c_grade = "7" then
'		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (acpt_man = '인터넷')"+ grade_sql
'	  else
'		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (acpt_man = '인터넷')"
'	end if
'	sql = sql + " group by CAST(acpt_date as date)"		
'  else
'	if c_grade = "7" or c_grade = "8" then
'		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (acpt_man = '인터넷') and (company = '"+company+"') "
'	  else
'		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (acpt_man = '인터넷') and (company = '"+company+"') "
'	end if
'	sql = sql + " group by CAST(acpt_date as date)"		
'end if


'Rs.Open Sql, Dbconn, 1
'do until rs.eof
'	for i = 0 to 30
'		if cstr(rs("com_date")) = acpt_date(i) then				
'			acpt_cnt_tab(i,3) = cint(rs("acpt_cnt"))
'			exit for
'		end if
'	next
'	rs.movenext()
'loop
'rs.close()
' 콜센터에서 인터넷 접수 제외
'for i = 0 to 30
'	acpt_cnt_tab(i,1) = acpt_cnt_tab(i,1) - acpt_cnt_tab(i,3)
'next


' 원격처리
if company = "전체" then
	if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type = '원격처리')"+ grade_sql
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type = '원격처리')"
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
  else
	if c_grade = "7" or c_grade = "8" then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type = '원격처리') and (company = '"+company+"') "
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type = '원격처리') and (company = '"+company+"') "
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
end if


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
if company = "전체" then
	if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리')"+ grade_sql
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리')"
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
  else
	if c_grade = "7" or c_grade = "8" then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리') and (company = '"+company+"') "
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리') and (company = '"+company+"') "
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
end if

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
if company = "전체" then
	if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리') and (acpt_man = mg_ce)"+ grade_sql
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리') and (acpt_man = mg_ce)"
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
  else
	if c_grade = "7" or c_grade = "8" then
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리') and (company = '"+company+"')  and (acpt_man = mg_ce)"
	  else
		sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (as_type <> '원격처리') and (company = '"+company+"')  and (acpt_man = mg_ce)"
	end if
	sql = sql + " group by CAST(acpt_date as date) order by CAST(acpt_date as date) asc"		
end if

'sql = "select CAST(acpt_date as date) as com_date, count(*) as acpt_cnt from as_acpt WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') and (mg_company = '(주)유니디아' or mg_company = '케이원') and (acpt_man = mg_ce)"		  
'sql = sql + " group by CAST(acpt_date as date)"

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
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/asset_header.asp" -->
			<!--#include virtual = "/include/asset_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=day_sum_asset.asp" method="post" name="frm">
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
								<strong>회사</strong>
							  	<%
                                    k = 0
                                    company_tab(0) = "전체"
                
                                    Sql="select * from etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' order by etc_name asc"
                                    Rs_etc.Open Sql, Dbconn, 1
                                    while not rs_etc.eof
                                        k = k + 1
                                        company_tab(k) = rs_etc("etc_name")
                                        rs_etc.movenext()
                                    Wend
                                    rs_etc.close()						
                                %>
                              	<select name="company" id="company" style="width:150px">
                                <% 
                                	for kk = 0 to k
                                %>
                                	<option value='<%=company_tab(kk)%>' <%If company_tab(kk) = company then %>selected<% end if %>><%=company_tab(kk)%></option>
                                <%
                                    next
                                %>
                              	</select>
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
                                <td class="first" style="background:#FFC">자체</td>
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
                                <td class="first" style="background:#FFC">인터넷</td>
                  				<% 
								sum_cnt = 0
								for i = 0 to 30 
								sum_cnt = sum_cnt + acpt_cnt_tab(i,3)
								%>

                                <td><%=acpt_cnt_tab(i,3)%></td>
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

