<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim pro_name(7)
dim pro_cnt(7)
dim err_name
dim company_tab(150)

for i = 0 to 7
	pro_cnt(i) = 0
next

pro_name(0) = "당일처리"
pro_name(1) = "익일처리"
pro_name(2) = "2일 처리"
pro_name(3) = "3일~ 6일"
pro_name(4) = "7일 이상"
pro_name(5) = "처리예정"
pro_name(6) = "입고중"
pro_name(7) = "미처리"

c_name = "전체"

'If ck_sw = "n" Then
'	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	company = request.form("company")
	as_type = request.form("as_type")
'  Else
'	from_date=Request("from_date")
'	to_date=Request("to_date")
'	company = "전체"
'End if

If to_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	as_type = "방문처리"
	company = "전체"
End If

curr_dd = cstr(datepart("d",to_date))
from_date = mid(to_date,1,8) + "01"

last_year = mid(to_date,1,4) - 1
last_month = mid(to_date,6,2) - 1

curr_year = mid(to_date,1,4)
if last_month = 0 then
	last_month = 12
	curr_year = last_year
end if

curr_month = mid(to_date,6,2)

if as_type = "전체" then
	type_sql = ""
  else
  	type_sql = " and (as_type ='"+as_type+"')"
end if

'당월 처리 내용 (총접수)
if company = "전체" then
	sql = "select count(*) as acpt_tot from as_acpt "
	sql = sql + "WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') " + type_sql
  else
	sql = "select count(*) as acpt_tot from as_acpt "
	sql = sql + "WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
	sql = sql + " and company = '" + company + "'" + type_sql
end if

Rs.Open Sql, Dbconn, 1
acpt_tot = cint(rs("acpt_tot"))
if rs.eof then
	acpt_tot = 0
end if
rs.close()

'전월 처리 내용 (총접수)
if company = "전체" then
	sql = "select count(*) as acpt_tot from as_acpt "
	sql = sql + "WHERE month(acpt_date) = "&last_month&" and year(acpt_date) ="&curr_year + type_sql
  else
	sql = "select count(*) as acpt_tot from as_acpt "
	sql = sql + "WHERE month(acpt_date) = "&last_month&" and year(acpt_date) ="&curr_year
	sql = sql + " and company = '" + company + "'" + type_sql
end if

Rs.Open Sql, Dbconn, 1

if rs.eof then
	last_tot = 0
  else
 	last_tot =cint(rs("acpt_tot"))
end if
rs.close()

'전년 당월 처리 내용 (총접수)
if company = "전체" then
	sql = "select count(*) as acpt_tot from as_acpt "
	sql = sql + "WHERE month(acpt_date) = "&curr_month&" and year(acpt_date) ="&last_year&type_sql
  else 
	sql = "select count(*) as acpt_tot from as_acpt "
	sql = sql + "WHERE month(acpt_date) = "&curr_month&" and year(acpt_date) ="&last_year
	sql = sql + " and company = '" + company + "'" + type_sql
end if

Rs.Open Sql, Dbconn, 1

if rs.eof then
	last_year = 0
  else
 	last_year =cint(rs("acpt_tot"))
end if
rs.close()

' 당월 처리 완료건
if company = "전체" then
	sql = "select CAST(acpt_date as date) as acpt_day, CAST((acpt_date + interval 10 DAY_HOUR) as date) as com_date, visit_date, substring(visit_time,1,2) as visit_hh, count(*) as err_cnt from as_acpt "
	sql = sql + " WHERE (as_process = '대체' or as_process = '완료' or as_process = '취소')"
	sql = sql + " and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')" + type_sql
	sql = sql + " GROUP BY CAST(acpt_date as date), CAST((acpt_date + interval 10 DAY_HOUR) as date), visit_date, substring(visit_time,1,2) Order By visit_date Asc"
  else
	sql = "select CAST(acpt_date as date) as acpt_day, CAST((acpt_date + interval 10 DAY_HOUR) as date) as com_date, visit_date, substring(visit_time,1,2) as visit_hh, count(*) as err_cnt from as_acpt "
	sql = sql + " WHERE (as_process = '대체' or as_process = '완료' or as_process = '취소') and (company ='" + company + "')"
	sql = sql + " and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')" + type_sql
	sql = sql + " GROUP BY CAST(acpt_date as date), CAST((acpt_date + interval 10 DAY_HOUR) as date), visit_date, substring(visit_time,1,2) Order By visit_date Asc"
end if  
Rs.Open Sql, Dbconn, 1

do until rs.eof

  	visit_date = datevalue(rs("visit_date"))
  	visit_day = datevalue(rs("visit_date"))

	if cstr(rs("visit_hh")) > "12" then
		visit_date = dateadd("d",1,visit_date)
	end if
	
	dd = datediff("d", rs("com_date"), visit_date)

	if cstr(visit_day) = cstr(rs("acpt_day")) then
		dd = 0
	end if

	if dd < 0 then
		dd = 0 
	end if
	

'휴일 계산
	if dd > 0 then
		a = datediff("d", rs("acpt_day"), visit_day)
		b = datepart("w",rs("acpt_day"))
		c = a + b
		d = a
		if a > 1 then
			if c > 7 then
				d = a - 2
			end if
		end if
		
'		visit_date = rs("visit_date")
		com_date = datevalue(rs("acpt_day"))
	
		do until com_date > visit_day
			sql_hol = "select * from holiday where holiday = '" + cstr(com_date) + "'"
			Set rs_hol=DbConn.Execute(SQL_hol)
			if rs_hol.eof or rs_hol.bof then
				d = d
			  else 
				d = d -1
			end if
			com_date = dateadd("d",1,com_date)
			rs_hol.close()
		loop
' 2012-02-06
		if d = 1 then
			visit_hh = int(rs("visit_hh"))
			if rs("acpt_day") <> rs("com_date") and visit_hh < 12 then
				d = 0
			end if
		end if
' 2012-02-06 end
		if d > 2 and d < 7 then
			d = 3
		end if
		if d > 6 then
			d = 4
		end if
		if d < 0 then
			d = 0
		end if
		pro_cnt(d) = pro_cnt(d) + cint(rs("err_cnt"))	
	  else

' 휴일 계산 끝
		pro_cnt(0) = pro_cnt(0) + cint(rs("err_cnt"))	
	end if
	rs.movenext()
loop
rs.close()
end_tot = pro_cnt(0) + pro_cnt(1) + pro_cnt(2) + pro_cnt(3) + pro_cnt(4)
pro_cnt(7) = acpt_tot - end_tot


'당월 처리 내용 (처리예정)
if company = "전체" then
	sql = "select count(*) as end_tot from as_acpt "
	sql = sql + "WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '접수' or as_process = '연기') and (request_date > '"+ to_date +"')" + type_sql
  else
	sql = "select count(*) as end_tot from as_acpt "
	sql = sql + "WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '접수' or as_process = '연기') and (request_date > '"+ to_date +"')"
	sql = sql + " and company = '" + company + "'" + type_sql
end if

Rs.Open Sql, Dbconn, 1
pro_cnt(5) = cint(rs("end_tot"))
pro_cnt(7) = pro_cnt(7) - pro_cnt(5)
if rs.eof then
	pro_cnt(5) = 0
end if
rs.close()

'당월 처리 내용 (입고)
if company = "전체" then
	sql = "select count(*) as end_tot from as_acpt "
	sql = sql + "WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '입고')" + type_sql
  else
	sql = "select count(*) as end_tot from as_acpt "
	sql = sql + "WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '입고')"
	sql = sql + " and company = '" + company + "'" + type_sql
end if

Rs.Open Sql, Dbconn, 1
pro_cnt(6) = cint(rs("end_tot"))
'pro_cnt(7) = pro_cnt(7) - pro_cnt(6)
if rs.eof then
	pro_cnt(6) = 0
end if
rs.close()

title_line = "처리 기간별 접수현황"
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
				return "2 1";
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
				if (document.frm.from_date.value > document.frm.to_date.value) {
					alert ("시작일이 종료일보다 클수가 없습니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/ceo_header.asp" -->
			<!--#include virtual = "/include/ceo_as_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=ceo_as_term_sum.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%="1900-01-01"%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<strong>회사</strong>
							  	<%
									sql="select * from trade where use_sw = 'Y'  and (trade_id = '매출' or trade_id = '공통') order by trade_name asc"
                                    rs_trade.Open Sql, Dbconn, 1
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
								<strong>처리유형</strong>
                                <select name="as_type" id="as_type" style="width:120px">
                                    <option value="전체" <%If as_type = "전체" then %>selected<% end if %>>전체</option>
                                    <option value="원격처리" <%If as_type = "원격처리" then %>selected<% end if %>>원격처리</option>
                                    <option value="방문처리" <%If as_type = "방문처리" then %>selected<% end if %>>방문처리</option>
                                    <option value="신규설치" <%If as_type = "신규설치" then %>selected<% end if %>>신규설치</option>
                                    <option value="신규설치공사" <%If as_type = "신규설치공사" then %>selected<% end if %>>신규설치공사</option>
                                    <option value="이전설치" <%If as_type = "이전설치" then %>selected<% end if %>>이전설치</option>
                                    <option value="이전설치공사" <%If as_type = "이전설치공사" then %>selected<% end if %>>이전설치공사</option>
                                    <option value="랜공사" <%If as_type = "랜공사" then %>selected<% end if %>>랜공사</option>
                                    <option value="이전랜공사" <%If as_type = "이전랜공사" then %>selected<% end if %>>이전랜공사</option>
                                    <option value="장비회수" <%If as_type = "장비회수" then %>selected<% end if %>>장비회수</option>
                                    <option value="예방점검" <%If as_type = "예방점검" then %>selected<% end if %>>예방점검</option>
                                    <option value="기타" <%If as_type = "기타" then %>selected<% end if %>>기타</option>
                                </select>
                              <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="12%" >
							<col width="12%" >
							<col width="12%" >
							<col width="12%" >
							<col width="12%" >
							<col width="12%" >
							<col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">당월접수</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">전 월</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">전 년</th>
								<th rowspan="2" scope="col">처리완료</th>
								<th rowspan="2" scope="col">미처리</th>
								<th rowspan="2" scope="col">처리율</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">전월접수</th>
							  <th scope="col">증감율</th>
							  <th scope="col">전년접수</th>
							  <th scope="col">증감율</th>
			              </tr>
						</thead>
						<tbody>
							<tr>
                                <th class="first"><%=formatnumber(clng(acpt_tot),0)%></th>
                                <th><%=formatnumber(clng(last_tot),0)%></th>
                                <th>
                            <% if last_tot = 0 then %>
                            	0%
                            <% else %>
                            	<%=formatnumber(((acpt_tot/last_tot * 100)-100),2)%>%
                            <% end if %>
                                </th>
                                <th><%=formatnumber(clng(last_year),0)%></th>
                                <th>
                            <% if last_year = 0 then %>
	                            0%
                            <% else %>
    	                        <%=formatnumber(((acpt_tot/last_year * 100)-100),2)%>%
                            <% end if %>
                                </th>
                                <th><%=formatnumber(clng(end_tot),0)%></th>
                                <th><%=formatnumber(clng(acpt_tot-end_tot),0)%></th>
                                <th>
                            <% if acpt_tot = 0 then %>
	                            0%
                            <% else %>
    	                        <%=formatnumber((end_tot/acpt_tot * 100),2)%>%
                            <% end if %>
                                </th>
							</tr>
						</tbody>
					</table>
					<h3 class="stit">* 시도별 내역</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="16%" >
							<col width="*" >
							<col width="12%" >
							<col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">처리기간</th>
								<th scope="col">그래프</th>
								<th scope="col">처리건수</th>
								<th scope="col">처리율(%)</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                              <th>총계</th>
                              <td class="left">&nbsp;</th>
                              <th><%=formatnumber(clng(acpt_tot),0)%></th>
                              <th>100%</th>
							</tr>
	                    <%
						for i = 0 to 7
							if	acpt_tot = 0 then
								pro_per = 0
							  else
								pro_per = formatnumber((pro_cnt(i)/acpt_tot * 100),2)
							end if
						%>
							<tr>
                              <td><%=pro_name(i)%></td>
                              <td class="left"><img src="image/graph02.gif" width="<%=pro_per*97/100%>%" height="15px" align="center"></th>
                              <td><%=formatnumber(clng(pro_cnt(i)),0)%></td>
                              <td><%=pro_per%>%</td>
							</tr>
                		<% 	
						next 
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

