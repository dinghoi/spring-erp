<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_tab
dim com_sum(7)
dim ok_sum(7)
dim mi_sum(7)
'dim com_cnt(7,9)
'dim com_in(7,9)
'dim sum_cnt(9)
'dim sum_in(9)
dim com_cnt(7,10)
dim com_in(7,10)
dim sum_cnt(10)
dim sum_in(10)
dim company_tab(150)
dim end_tab(11)
dim mi_tab(11)
dim curr_mi_tab(11)
dim mi_in
com_tab = array("본사","부산지사","대구지사","대전지사","광주지사","전주지사","강원지사","제주지사")

for i = 0 to 7
	com_sum(i) = 0
	ok_sum(i) = 0
	mi_sum(i) = 0
	'for j = 0 to 9
	for j = 0 to 10
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
to_date = mid(cstr(now()),1,10)
  as_type = "방문처리"
  company = "전체"
  mg_group = "1"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

type_sql = " (as_type ='방문처리') and "
'type_sql = " (as_acpt.as_type ='방문처리') and "
mg_group_sql = " (mg_group ='1') and "

tot_cnt = 0

' 미처리건
'sql = "select as_acpt.sido, as_acpt.as_process, Cast(acpt_date as date) as acpt_day, CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) as com_date, count(*) as err_cnt from as_acpt"
'sql = sql + " WHERE "+type_sql+mg_group_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
'sql = sql + " GROUP BY sido, as_process, Cast(acpt_date as date), CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) Order By as_acpt.sido Asc"


sql = "select as_acpt.sido, as_acpt.as_process, Cast(request_date as date) as acpt_day, CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date) as com_date, count(*) as err_cnt from as_acpt"
sql = sql + " WHERE "+type_sql+mg_group_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
sql = sql + " AND CAST(request_date AS DATE) <= now()"
sql = sql + " GROUP BY sido, as_process, Cast(request_date as date), CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date) Order By as_acpt.sido Asc"
Rs.Open Sql, Dbconn, 1

do until rs.eof
'	com_tab(i) = rs("sido")
	select case rs("sido")
		case "서울"
			i = 0
		case "경기"
			i = 0
		case "인천"
			i = 0
		case "부산"
			i = 1
		case "울산"
			i = 1
		case "경남"
			i = 1
		case "대구"
			i = 2
		case "경북"
			i = 2
		case "대전"
			i = 3
		case "충남"
			i = 3
		case "충북"
			i = 3
		case "세종"
			i = 3
		case "광주"
			i = 4
		case "전남"
			i = 4
		case "전북"
			i = 5
		case "강원"
			i = 6
		case "제주"
			i = 7
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
		'b = datepart("w",rs("acpt_day"))
		'bb = datepart("w", curr_day)
		'if bb = 1 then
		'	a = a -1
		'end if
		'c = a + b
		d = a
		'if a > 1 then
		'	if c > 7 then
		'		d = a - 2
		'	end if
		'end if
		
		com_date = datevalue(rs("acpt_day"))
		
		do until com_date > curr_day
			sql_hol = "select * from (select DAYOFWEEK('" + cstr(com_date) + "') as  dayweeks ) A where A.dayweeks in (1,7)" 
			Set rs_wek=DbConn.Execute(SQL_hol)
			if rs_wek.eof or rs_wek.bof then
				d = d
			  else 
				d = d -1
			end if
			com_date = dateadd("d",1,com_date)
			rs_wek.close()
		loop
		
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
'		  elseif d > 2 and d < 7  then
'			j = 8
'		  else
'			j = 9
		  elseif d = 3 then
			j = 8
		  elseif d = 4 then
			j = 9
		  else
			j = 10
		end if
		com_cnt(i,j) = com_cnt(i,j) + clng(rs("err_cnt"))	

		if rs("as_process") = "입고" then		
			com_in(i,j) = com_in(i,j) + clng(rs("err_cnt"))
		end if
	  else
' 휴일 계산 끝
		com_cnt(i,5) = com_cnt(i,5) + clng(rs("err_cnt"))
		'com_cnt(i,6) = com_cnt(i,6) + clng(rs("err_cnt"))

		if rs("as_process") = "입고" then		
			com_in(i,5) = com_in(i,5) + clng(rs("err_cnt"))
			'com_in(i,6) = com_in(i,6) + clng(rs("err_cnt"))
		end if
	end if
	tot_cnt = tot_cnt + clng(rs("err_cnt"))
	rs.movenext()
loop
rs.close()

title_line = "방문처리 지사별 미처리 현황 (요청일 기준)"


dim com_tab2
dim com_cnt2(8,2)

for i = 0 to 8
	for j = 0 to 2
		com_cnt2(i,j) = 0
	next
next
com_tab2 = array("계","본사","부산지사","대구지사","대전지사","광주지사","전주지사","강원지사","제주지사")

sql = "SELECT A.SIDO"
sql = sql + "      , CNT_1"
sql = sql + "      , CNT_2"
sql = sql + "      , CNT_3"
sql = sql + " FROM ("

sql = sql + "   select"
sql = sql + "        IFNULL(A.SIDO,'0') AS SIDO"
sql = sql + "       , SUM(IF(A.REAL_DATE_DIFF = 1 , ERR_CNT ,0))   AS 'CNT_1'"
sql = sql + "       , SUM(IF(A.REAL_DATE_DIFF = 2 , ERR_CNT ,0)) AS 'CNT_2'"
sql = sql + "       , SUM(IF(A.REAL_DATE_DIFF = 3 , ERR_CNT ,0)) AS 'CNT_3'"

sql = sql + "   FROM"
sql = sql + "   ("
sql = sql + "     SELECT case A.SIDO WHEN '서울' then '1'"
sql = sql + "                        WHEN '경기' then '1'"
sql = sql + "                        WHEN '부산' then '2'"
sql = sql + "                        WHEN '대구' then '3'"
sql = sql + "                        WHEN '인천' then '1'"
sql = sql + "                        WHEN '광주' then '5'"
sql = sql + "                        WHEN '대전' then '4'"
sql = sql + "                        WHEN '울산' then '2'"
sql = sql + "                        WHEN '강원' then '7'"
sql = sql + "                        WHEN '경남' then '2'"
sql = sql + "                        WHEN '경북' then '3'"
sql = sql + "                        WHEN '충남' then '4'"
sql = sql + "                        WHEN '충북' then '4'"
sql = sql + "                        WHEN '세종' then '4'"
sql = sql + "                        WHEN '전남' then '5'"
sql = sql + "                        WHEN '전북' then '6'"
sql = sql + "                        WHEN '제주' then '8'"
sql = sql + "                        ELSE 17 END SIDO"
sql = sql + "            ,A.AS_PROCESS"
sql = sql + "            ,A.ACPT_DATE"
sql = sql + "			      ,DATEDIFF(ACPT_DATE,VISIT_DATE)  AS REAL_DATE_DIFF"
sql = sql + "            ,A.ERR_CNT"

sql = sql + "     FROM"
sql = sql + "     ("
sql = sql + "             SELECT   A.SIDO, A.AS_PROCESS"
sql = sql + "                    , CAST(request_date AS DATE) AS ACPT_DATE"
sql = sql + "                    , CAST((A.request_date + INTERVAL 10 DAY_HOUR) AS DATE) AS COM_DATE"
sql = sql + "					 , NOW() AS VISIT_DATE"
sql = sql + "					 , ADDDATE(NOW(),INTERVAL 12 DAY_HOUR) AS VISIT_DAY"
sql = sql + "                    , COUNT(*) AS ERR_CNT"
sql = sql + "               FROM   AS_ACPT A"
sql = sql + "              WHERE   1=1"
sql = sql + "              AND   A.AS_PROCESS IN( '접수' ,'입고' , '연기')"
sql = sql + "              AND   AS_TYPE ='방문처리'"
sql = sql + "			        AND   CAST(request_date AS DATE) > now()"
sql = sql + "             GROUP BY   A.SIDO, A.AS_PROCESS, Cast(request_date as date),CAST((A.request_date + INTERVAL 10 DAY_HOUR) AS DATE)"
sql = sql + "     ) A"
sql = sql + "   Order By A.sido Asc"
sql = sql + "   ) A"

sql = sql + "   group by A.SIDO WITH ROLLUP"
sql = sql + " ) A"
sql = sql + " order by A.SIDO"

Rs.Open Sql, Dbconn, 1


    
do until rs.eof

	i = clng(rs("sido")) 
    com_cnt2(i,0) = clng(rs("cnt_1"))
    com_cnt2(i,1) = clng(rs("cnt_2"))
    com_cnt2(i,2) = clng(rs("cnt_3"))
    
	rs.movenext()
loop
rs.close()
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
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
				<div class="gView" style="position: relative;">
					<h3 class="stit">* 현재시간 : <%=now()%></h3>
					<table cellpadding="0" cellspacing="0" class="tableList2">
						<colgroup>
							<col width="*" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">지사</th>

								<th colspan="2" style=" border-left:1px solid #e3e3e3;border-bottom:1px solid #e3e3e3;" scope="col">당일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">익일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">2일</th>
								<!--
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">3일~6일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">7일이상</th>
								-->
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">3일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">4일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">5일이상</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">소계</th>
								<th rowspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">백분율</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">건수</th>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
						  </tr>
						</thead>
						<tbody>
						<% 	
                    	if tot_cnt > 0 then
                        	k = 0
                      	  else
                        	k = 7
                    	end if
        
		'--------------------------------------여기 확인
                    	for i = k to 7 
                        	if	com_tab(i) <> "" then
        
								for j = 0 to 4
									ok_sum(i) = ok_sum(i) + com_cnt(i,j)
									sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)				
								next
								'for j = 5 to 9
								for j = 5 to 10
									mi_sum(i) = mi_sum(i) + com_cnt(i,j)
									sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)				
									sum_in(j) = sum_in(j) + com_in(i,j)				
								next
								com_sum(i) = ok_sum(i) + mi_sum(i)
				
								sido = com_tab(i)
							end if
						next
		'--------------------------------------여기 확인
                		%>
							<tr>
                              <th>계</th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(5)),0)%></a></th>
                              <th class="right"><%=sum_in(5)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(6)),0)%></a></th>
                              <th class="right"><%=sum_in(6)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(7)),0)%></a></th>
                              <th class="right"><%=sum_in(7)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(8)),0)%></a></th>
                              <th class="right"><%=sum_in(8)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(9)),0)%></a></th>
                              <th class="right"><%=sum_in(9)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(10)),0)%></a></th>
                              <th class="right"><%=sum_in(10)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)+sum_cnt(10)),0)%></th>
                              <th class="right"><%=sum_in(5)+sum_in(6)+sum_in(7)+sum_in(8)+sum_in(9)+sum_in(10)%></th>
                              <th class="right">
                              <% if tot_cnt = 0 then %>
                                    0%
                                <% else %>
                                    <%=formatnumber(((sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)+sum_cnt(10))/tot_cnt * 100),2)%>%
                                <% end if %>
                              &nbsp;
                              </th>
							</tr>
						<% 	
                    	if tot_cnt > 0 then
                        	k = 0
                      	  else
                        	k = 7
                    	end if
        
                    	for i = k to 7 
                        	if	com_tab(i) <> "" then
                		%>
							<tr>
                              <td><%=com_tab(i)%></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=0%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,5)),0)%></td>
                              <td class="right"><%=com_in(i,5)%></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,6)),0)%></td>
                              <td class="right"><%=com_in(i,6)%></td>
                              <td class="right" bgcolor="#FFFF88"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=formatnumber(clng(com_cnt(i,7)),0)%></strong></td>
                              <td class="right"><strong><%=com_in(i,7)%></strong></td>
                              <td class="right" bgcolor="#FFBE7D"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=formatnumber(clng(com_cnt(i,8)),0)%></strong></td>
                              <td class="right"><strong><%=com_in(i,8)%></strong></td>
                              <td class="right" bgcolor="#FF8080"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=4%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=formatnumber(clng(com_cnt(i,9)),0)%></strong></td>
                              <td class="right"><strong><%=com_in(i,9)%></strong></td>

<!-- 추가 4일 -->
                              <td class="right" bgcolor="#FF8080"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=5%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=formatnumber(clng(com_cnt(i,10)),0)%></strong></td>
                              <td class="right"><strong><%=com_in(i,10)%></strong></td>
<!-- 추가 4일 -->

                              <td class="right"><a  href="#" onClick="pop_Window('as_michulri_popup_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(mi_sum(i)),0)%></td>
                              <td class="right"><%=com_in(i,5)+com_in(i,6)+com_in(i,7)+com_in(i,8)+com_in(i,9)+com_in(i,10)%></td>
                              <td class="right">
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
					<h3 class="stit1" >* 요청일 이후 처리 예정 건수</h3>					
					<table cellpadding="0" cellspacing="0" class="tableList1">
						<colgroup>
							<col width="*">
							<col width="25%">
							<col width="25%">
							<col width="25%">
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">지사</th>
								<th colspan="1" style=" border-left:1px solid #e3e3e3;border-bottom:1px solid #e3e3e3;" scope="col"><%=datevalue(mid(dateadd("d",1,now()),1,10))%></th>
								<th colspan="1" style="border-bottom:1px solid #e3e3e3;" scope="col"><%=datevalue(mid(dateadd("d",2,now()),1,10))%></th>
								<th colspan="1" style="border-bottom:1px solid #e3e3e3;" scope="col"><%=datevalue(mid(dateadd("d",3,now()),1,10))%></th>

							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">건수</th>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
						  </tr>
						</thead>
						<tbody>
							<tr>
                              <th>계</th>
                              <th class="right"><%=com_cnt2(0,0)%></th>
                              <th class="right"><%=com_cnt2(0,1)%></th>
                              <th class="right"><%=com_cnt2(0,2)%></th>
							</tr>						
						<% 
							for i = 1 to 8 
                        	if	com_tab2(i) <> "" then
                		%>
							<tr>
                              <td><%=com_tab2(i)%></td>
                              <td class="right"><a  href="#" onClick="pop_Window('as_michulri_popup_request.asp?from_date=<%=datevalue(mid(dateadd("d",1,now()),1,10))%>&to_date=<%=datevalue(mid(dateadd("d",1,now()),1,10))%>&sido=<%=com_tab2(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt2(i,0)),0)%></td>
                              <td class="right"><a  href="#" onClick="pop_Window('as_michulri_popup_request.asp?from_date=<%=datevalue(mid(dateadd("d",2,now()),1,10))%>&to_date=<%=datevalue(mid(dateadd("d",2,now()),1,10))%>&sido=<%=com_tab2(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt2(i,1)),0)%></td>
                              <td class="right"><a  href="#" onClick="pop_Window('as_michulri_popup_request.asp?from_date=<%=datevalue(mid(dateadd("d",3,now()),1,10))%>&to_date=<%=datevalue(mid(dateadd("d",3,now()),1,10))%>&sido=<%=com_tab2(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt2(i,2)),0)%></td>
                              
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

