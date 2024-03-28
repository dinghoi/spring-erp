<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_tab
dim com_sum(16)
dim ok_sum(16)
dim mi_sum(16)
dim com_cnt(16,9)
dim com_in(16,9)
dim sum_cnt(9)
dim sum_in(9)
dim company_tab(150)
dim end_tab(11)
dim mi_tab(11)
dim curr_mi_tab(11)
dim mi_in
com_tab = array("서울","경기","부산","대구","인천","광주","대전","울산","강원","경남","경북","충남","충북","세종","전남","전북","제주")

from_date=Request.form("from_date")
to_date=Request.form("to_date")
as_type=Request.form("as_type")
company=Request.form("company")
mg_group=Request.form("mg_group")

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	as_type = "방문처리"
	company = "전체"
	mg_group = "1"
End If

for i = 0 to 16
	com_sum(i) = 0
	ok_sum(i) = 0
	mi_sum(i) = 0
	for j = 0 to 9
		com_cnt(i,j) = 0
		com_in(i,j) = 0
		sum_cnt(j) = 0
		sum_in(j) = 0
	next
next
for i = 0 to 11
	end_tab(i) = 0
	mi_tab(i) = 0
	curr_mi_tab(i) = 0
next

curr_day = datevalue(mid(cstr(now()),1,10))
curr_date = datevalue(mid(dateadd("h",12,now()),1,10))

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

if company = "전체" then
	com_sql0 = ""
	com_sql = ""
  else
  	com_sql0 = " (company ='"+company+"') and "
  	com_sql = " (as_acpt.company ='"+company+"') and "
end if
if as_type = "전체" then
	type_sql0 = ""
	type_sql = ""
  else
  	type_sql0 = " (as_type ='"+as_type+"') and "
  	type_sql = " (as_acpt.as_type ='"+as_type+"') and "
end if
if mg_group = "0" then
	mg_group_sql = ""
  else
  	mg_group_sql = " (mg_group ='"+mg_group+"') and "
end if

tot_cnt = 0
' 처리완료
sql = "select as_type, count(*) as end_cnt from as_acpt"
sql = sql + " where "+com_sql0+mg_group_sql+" (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"') and (as_process = '대체' or as_process = '완료' or as_process = '취소') "
sql = sql + " GROUP BY as_type Order By as_type Asc"
Rs.Open Sql, Dbconn, 1 

do until rs.eof
	end_cnt = clng(rs("end_cnt"))
	end_tab(0) = end_tab(0) + end_cnt
	
	if rs("as_type") = "원격처리" then
		end_tab(1) = end_tab(1) + end_cnt
	end if
	if rs("as_type") = "방문처리" then
		end_tab(2) = end_tab(2) + end_cnt
	end if
	if rs("as_type") = "신규설치" then
		end_tab(3) = end_tab(3) + end_cnt
	end if
	if rs("as_type") = "신규설치공사" then
		end_tab(4) = end_tab(4) + end_cnt
	end if
	if rs("as_type") = "이전설치" then
		end_tab(5) = end_tab(5) + end_cnt
	end if
	if rs("as_type") = "이전설치공사" then
		end_tab(6) = end_tab(6) + end_cnt
	end if
	if rs("as_type") = "랜공사" then
		end_tab(7) = end_tab(7) + end_cnt
	end if
	if rs("as_type") = "이전랜공사" then
		end_tab(8) = end_tab(8) + end_cnt
	end if
	if rs("as_type") = "장비회수" then
		end_tab(9) = end_tab(9) + end_cnt
	end if
	if rs("as_type") = "예방점검" then
		end_tab(10) = end_tab(10) + end_cnt
	end if
	if rs("as_type") = "기타" then
		end_tab(11) = end_tab(11) + end_cnt
	end if

	rs.movenext()
loop
rs.close()

'현재까지 미처리
sql = "select as_type, count(*) as end_cnt from as_acpt"
sql = sql + " where "+com_sql0+mg_group_sql+" (Cast(acpt_date as date) <= '"+to_date+"') and (as_process = '접수' or as_process = '입고' or as_process = '연기') "
sql = sql + " GROUP BY as_type Order By as_type Asc"
Rs.Open Sql, Dbconn, 1 

do until rs.eof
	end_cnt = clng(rs("end_cnt"))
	curr_mi_tab(0) = curr_mi_tab(0) + end_cnt
	
	if rs("as_type") = "원격처리" then
		curr_mi_tab(1) = curr_mi_tab(1) + end_cnt
	end if
	if rs("as_type") = "방문처리" then
		curr_mi_tab(2) = curr_mi_tab(2) + end_cnt
	end if
	if rs("as_type") = "신규설치" then
		curr_mi_tab(3) = curr_mi_tab(3) + end_cnt
	end if
	if rs("as_type") = "신규설치공사" then
		curr_mi_tab(4) = curr_mi_tab(4) + end_cnt
	end if
	if rs("as_type") = "이전설치" then
		curr_mi_tab(5) = curr_mi_tab(5) + end_cnt
	end if
	if rs("as_type") = "이전설치공사" then
		curr_mi_tab(6) = curr_mi_tab(6) + end_cnt
	end if
	if rs("as_type") = "랜공사" then
		curr_mi_tab(7) = curr_mi_tab(7) + end_cnt
	end if
	if rs("as_type") = "이전랜공사" then
		curr_mi_tab(8) = curr_mi_tab(8) + end_cnt
	end if
	if rs("as_type") = "장비회수" then
		curr_mi_tab(9) = curr_mi_tab(9) + end_cnt
	end if
	if rs("as_type") = "예방점검" then
		curr_mi_tab(10) = curr_mi_tab(10) + end_cnt
	end if
	if rs("as_type") = "기타" then
		curr_mi_tab(11) = curr_mi_tab(11) + end_cnt
	end if

	rs.movenext()
loop
rs.close()
'기간내입고건
'sql = "select as_type, count(*) as end_cnt from as_acpt"
'sql = sql + " where "+com_sql0+mg_group_sql+" (Cast(acpt_date as date) <= '"+to_date+"') and (as_process = '입고') "
'sql = sql + " GROUP BY as_type Order By as_type Asc"
'Rs.Open Sql, Dbconn, 1 

'mi_in = 0
'do until rs.eof
'	end_cnt = clng(rs("end_cnt"))
'	mi_in = mi_in + end_cnt	
'	rs.movenext()
'loop
'rs.close()
'현재까지 입고건
sql = "select as_type, count(*) as end_cnt from as_acpt"
sql = sql + " where "+com_sql0+mg_group_sql+" (Cast(acpt_date as date) <= '"+to_date+"') and (as_process = '입고') "
sql = sql + " GROUP BY as_type Order By as_type Asc"
Rs.Open Sql, Dbconn, 1 

curr_mi_in = 0
do until rs.eof
	end_cnt = clng(rs("end_cnt"))
	curr_mi_in = curr_mi_in + end_cnt	
	rs.movenext()
loop
rs.close()

' 총계 SQL -> 전체 SUM 으로 변경
'sql = "select count(*) as err_cnt from as_acpt"
''sql = sql + " WHERE (mg_group='"+mg_group+"') and (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"') and (reside_place = '콜센터') and (as_type <> '원격처리')"
''sql = sql + " WHERE (mg_group='"+mg_group+"') and (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"') and (as_type = '방문처리')"
'sql = sql + " WHERE "+com_sql0+type_sql0+"(mg_group='"+mg_group+"') and (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"')"
'Rs.Open Sql, Dbconn, 1

'if rs.eof then
'	tot_cnt = 0
'  else
'  	tot_cnt = clng(rs("err_cnt"))
'end if
'rs.close()

' 완료건
sql = "select as_acpt.sido, Cast(acpt_date as date) as acpt_day, CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) as com_date, visit_date, substring(visit_time,1,2) as visit_hh, count(*) as err_cnt from as_acpt"
'sql = sql + " WHERE (as_acpt.mg_group='"+mg_group+"') and (k1_etc_code.etc_type = '81') and (as_acpt.as_process = '취소' or as_acpt.as_process = '완료') and (as_acpt.as_type <> '원격처리') and (reside_place = '콜센터')"
sql = sql + " WHERE "+com_sql+type_sql+mg_group_sql+" (as_acpt.as_process = '대체' or as_acpt.as_process = '완료' or as_acpt.as_process = '취소')"
sql = sql + " and (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"')"
sql = sql + " GROUP BY as_acpt.sido, Cast(acpt_date as date), CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date), visit_date, substring(visit_time,1,2) Order By as_acpt.sido Asc"
Rs.Open Sql, Dbconn, 1

do until rs.eof
	select case rs("sido")
		case "서울"
			i = 0
		case "경기"
			i = 1
		case "부산"
			i = 2
		case "대구"
			i = 3
		case "인천"
			i = 4
		case "광주"
			i = 5
		case "대전"
			i = 6
		case "울산"
			i = 7
		case "강원"
			i = 8
		case "경남"
			i = 9
		case "경북"
			i = 10
		case "충남"
			i = 11
		case "충북"
			i = 12
		case "세종"
			i = 13
		case "전남"
			i = 14
		case "전북"
			i = 15
		case "제주"
			i = 16
	end select	

  	visit_date = datevalue(rs("visit_date"))
' 1/19 추가
  	visit_day = datevalue(rs("visit_date"))
' 1/19 추가 end

	if cstr(rs("visit_hh")) > "12" then
		visit_date = dateadd("d",1,visit_date)
	end if
	
	dd = datediff("d", rs("com_date"), visit_date)

	if cstr(rs("visit_date")) = cstr(rs("acpt_day")) then
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
' 임시 0224		
		if	d < 0 then
			d = 0
		end if
' 임시 End
' 1/19 추가
'		ddd = datediff("d", rs("acpt_day"), visit_day)
'		if d > ddd then
'			d = ddd
'		end if
' 1/19 추가 end
		com_cnt(i,d) = com_cnt(i,d) + clng(rs("err_cnt"))	
	  else

' 휴일 계산 끝
		com_cnt(i,0) = com_cnt(i,0) + clng(rs("err_cnt"))
	end if
	tot_cnt = tot_cnt + clng(rs("err_cnt"))
	rs.movenext()
loop
rs.close()

' 미처리건
sql = "select as_acpt.sido, as_acpt.as_process, Cast(acpt_date as date) as acpt_day, CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) as com_date, count(*) as err_cnt from as_acpt"
sql = sql + " WHERE "+com_sql+type_sql+mg_group_sql+" (as_acpt.as_process = '접수' or as_acpt.as_process = '입고' or as_acpt.as_process = '연기')"
sql = sql + " and (Cast(acpt_date as date) <= '"+to_date+"')"
sql = sql + " GROUP BY as_acpt.sido, as_acpt.as_process, Cast(acpt_date as date), CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) Order By as_acpt.sido Asc"
Rs.Open Sql, Dbconn, 1

do until rs.eof
'	i = int(rs("etc_code")) - 8101
'	com_tab(i) = rs("sido")
	select case rs("sido")
		case "서울"
			i = 0
		case "경기"
			i = 1
		case "부산"
			i = 2
		case "대구"
			i = 3
		case "인천"
			i = 4
		case "광주"
			i = 5
		case "대전"
			i = 6
		case "울산"
			i = 7
		case "강원"
			i = 8
		case "경남"
			i = 9
		case "경북"
			i = 10
		case "충남"
			i = 11
		case "충북"
			i = 12
		case "세종"
			i = 13
		case "전남"
			i = 14
		case "전북"
			i = 15
		case "제주"
			i = 16
	end select	

	dd = datediff("d", rs("com_date"), curr_date)

	if dd < 0 then
		dd = 0 
	end if
	
	if cstr(curr_day) = cstr(rs("acpt_day")) then
		dd = 0
	end if

'휴일 계산
	if dd > 0 then
		a = datediff("d", rs("acpt_day"), curr_day)
		b = datepart("w",rs("acpt_day"))
		bb = datepart("w", curr_day)
		if bb = 1 then
			a = a -1
		end if
		c = a + b
		d = a
		if a > 1 then
			if c > 7 then
				d = a - 2
			end if
		end if
		
'		visit_date = rs("visit_date")
		com_date = datevalue(rs("acpt_day"))
'		act_date = com_date
	
		do until com_date > curr_day
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
			curr_hh = int(datepart("h",now()))
			if rs("acpt_day") <> rs("com_date") and curr_hh < 12 then
				d = 0
			end if
		end if
' 2012-02-06 end
		if d = 0 then
			j = 5
		  elseif d = 1 then
			j = 6
		  elseif d = 2 then
			j = 7
		  elseif d > 2 and d < 7  then
			j = 8
		  else
			j = 9
		end if
		com_cnt(i,j) = com_cnt(i,j) + clng(rs("err_cnt"))	

		if rs("as_process") = "입고" then		
			com_in(i,j) = com_in(i,j) + clng(rs("err_cnt"))
		end if
	  else
' 휴일 계산 끝
		com_cnt(i,5) = com_cnt(i,5) + clng(rs("err_cnt"))

		if rs("as_process") = "입고" then		
			com_in(i,5) = com_in(i,5) + clng(rs("err_cnt"))
		end if
	end if
	tot_cnt = tot_cnt + clng(rs("err_cnt"))
	rs.movenext()
loop
rs.close()

title_line = "지역별 처리 기간별 현황(접수일 기준)"
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
				return "3 1";
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/sum_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=area_term_pro.asp" method="post" name="frm">
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
									sql="select * from trade where use_sw = 'Y'  and (trade_id = '매출' or trade_id = '공용') order by trade_name asc"
                                    rs_trade.Open Sql, Dbconn, 1
                                %>
        						<select name="company" id="company">
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
                                <select name="as_type" id="as_type" style="width:100px">
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
								<strong>관리그룹</strong>
                              	<input type="radio" name="mg_group" value="0" <% if mg_group = "0" then %>checked<% end if %> style="width:25px" onClick="condi_view()">전체                               
                                <input type="radio" name="mg_group" value="1" <% if mg_group = "1" then %>checked<% end if %> style="width:25px" onClick="condi_view()">일반그룹
                                <input type="radio" name="mg_group" value="2" <% if mg_group = "2" then %>checked<% end if %> style="width:25px" onClick="condi_view()">한진그룹
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">처리유무</th>
								<th scope="col">합계</th>
								<th scope="col">원격처리</th>
								<th scope="col">방문처리</th>
								<th scope="col">신규설치</th>
								<th scope="col">신규설치.공사</th>
								<th scope="col">이전설치</th>
								<th scope="col">이전설치.공사</th>
								<th scope="col">랜공사</th>
								<th scope="col">이전랜공사</th>
								<th scope="col">장비회수</th>
								<th scope="col">예방점검</th>
								<th scope="col">기타</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                                <td class="first">처리완료</td>
                                <td><%=formatnumber(end_tab(0),0)%></td>
                                <td><%=formatnumber(end_tab(1),0)%></td>
                                <td><%=formatnumber(end_tab(2),0)%></td>
                                <td><%=formatnumber(end_tab(3),0)%></td>
                                <td><%=formatnumber(end_tab(4),0)%></td>
                                <td><%=formatnumber(end_tab(5),0)%></td>
                                <td><%=formatnumber(end_tab(6),0)%></td>
                                <td><%=formatnumber(end_tab(7),0)%></td>
                                <td><%=formatnumber(end_tab(8),0)%></td>
                                <td><%=formatnumber(end_tab(9),0)%></td>
                                <td><%=formatnumber(end_tab(10),0)%></td>
                                <td><%=formatnumber(end_tab(11),0)%></td>
							</tr>
							<tr>
                                <td class="first">전체미처리</td>
                                <td><%=formatnumber(curr_mi_tab(0),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(1),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(2),0)%>&nbsp;(<%=curr_mi_in%>)</td>
                                <td><%=formatnumber(curr_mi_tab(3),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(4),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(5),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(6),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(7),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(8),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(9),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(10),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(11),0)%></td>
							</tr>
						</tbody>
					</table>
					<h3 class="stit">* 시도별 내역</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">시도</th>
								<th scope="col" colspan="6" style=" border-bottom:1px solid #e3e3e3;">기간내 처리 완료</th>
								<th scope="col" colspan="6" style=" border-bottom:1px solid #e3e3e3;">현재까지 미처리 * 괄호는 입고건</th>
								<th scope="col" rowspan="2">시도계</th>
								<th scope="col" rowspan="2">백분율</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">당일</th>
								<th scope="col">익일</th>
								<th scope="col">2일</th>
								<th scope="col">3일~6일</th>
								<th scope="col">7일이상</th>
								<th scope="col">소계</th>
								<th scope="col">당일</th>
								<th scope="col">익일</th>
								<th scope="col">2일</th>
								<th scope="col">3일~6일</th>
								<th scope="col">7일이상</th>
								<th scope="col">소계</th>
							</tr>
						</thead>
						<tbody>
						<% 	
                    	if tot_cnt > 0 then
                        	k = 0
                      	  else
                        	k = 16
                    	end if
        
                    	for i = k to 16 
                        	if	com_tab(i) <> "" then
        
								for j = 0 to 4
									ok_sum(i) = ok_sum(i) + com_cnt(i,j)
									sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)				
								next
								for j = 5 to 9
									mi_sum(i) = mi_sum(i) + com_cnt(i,j)
									sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)				
									sum_in(j) = sum_in(j) + com_in(i,j)				
								next
								com_sum(i) = ok_sum(i) + mi_sum(i)
				
								sido = com_tab(i)
							end if
						next
                		%>
							<tr>
                              <th>계</th>
                              <th><%=formatnumber(clng(sum_cnt(0)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(1)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(2)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(3)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(4)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)),0)%>&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%="총괄"%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=0%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(5)),0)%></a>(<%=sum_in(5)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%="총괄"%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(6)),0)%></a>(<%=sum_in(6)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%="총괄"%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(7)),0)%></a>(<%=sum_in(7)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%="총괄"%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(8)),0)%></a>(<%=sum_in(8)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%="총괄"%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=7%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(9)),0)%></a>(<%=sum_in(9)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('as_michulri_popup.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%="총괄"%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)),0)%>(<%=sum_in(5)+sum_in(6)+sum_in(7)+sum_in(8)+sum_in(9)%>)&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)),0)%>&nbsp;</th>
                              <th>
                              <% if tot_cnt = 0 then %>
                                    0%
                                <% else %>
                                    <%=formatnumber(((sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9))/tot_cnt * 100),2)%>%
                                <% end if %>
                              &nbsp;
                              </th>
							</tr>
						<% 	
                    	if tot_cnt > 0 then
                        	k = 0
                      	  else
                        	k = 16
                    	end if
        
                    	for i = k to 16 
                        	if	com_tab(i) <> "" then
                		%>
							<tr>
                              <td><%=com_tab(i)%></td>
                              <td><%=formatnumber(clng(com_cnt(i,0)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(com_cnt(i,1)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(com_cnt(i,2)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(com_cnt(i,3)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(com_cnt(i,4)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(ok_sum(i)),0)%>&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=0%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,5)),0)%></a>(<%=com_in(i,5)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,6)),0)%></a>(<%=com_in(i,6)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,7)),0)%></a>(<%=com_in(i,7)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,8)),0)%></a>(<%=com_in(i,8)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=7%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,9)),0)%></a>(<%=com_in(i,9)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('as_michulri_popup.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(mi_sum(i)),0)%></a>(<%=com_in(i,5)+com_in(i,6)+com_in(i,7)+com_in(i,8)+com_in(i,9)%>)&nbsp;</td>
                              <td><%=formatnumber(clng(com_sum(i)),0)%>&nbsp;</td>
                              <td>
                              <% if tot_cnt = 0 then %>
                                    0%
                                <% else %>
                                    <%=formatnumber((com_sum(i)/tot_cnt * 100),2)%>%
                                <% end if %>
                              &nbsp;
                              </td>
							</tr>
                		<% 	
							end if
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

