<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_tab(15)
dim com_sum(15)
dim ok_sum(15)
dim mi_sum(15)
dim com_cnt(15,7)
dim sum_cnt(7)
dim company_tab(150)
dim end_tab(8)
dim mi_tab(8)

from_date = request("from_date")
to_date = request("to_date")
curr_day = datevalue(mid(cstr(now()),1,10))
curr_date = datevalue(mid(dateadd("h",12,now()),1,10))
sido = request("sido")
mg_ce = request("mg_ce")
mg_ce_id = request("mg_ce_id")
mg_group = request("mg_group")
company = request("company")
as_type = request("as_type")
days = int(request("days"))
win_sw = "back"
dis_days = cstr(days) + "일"
if days = 3 then
	dis_days = "3~6일"
end if
if days = 7 then
	dis_days = "7일이상"
end if

if company = "" then
	company = "전체"
	as_type = "전체"
end if

if mg_ce = "" then
	memo01 = "시도"
	memo02 = sido
  else
	memo01 = "담당자"
	memo02 = mg_ce
end if

if company = "전체" then
	com_sql = ""
  else
  	com_sql = " (company ='"+company+"') and "
end if
if as_type = "전체" then
	type_sql = ""
  else
  	type_sql = " (as_type ='"+as_type+"') and "
end if

i = 0
in_cnt = 0
acpt_cnt = 0
yun_cnt = 0

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_mi = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

' 미처리건
if	mg_ce = "" then
	if sido = "총괄" then
		sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
		sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"')"
  elseif sido = "계" then
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"')"
  elseif sido = "본사" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('서울','경기','인천')"
  elseif sido = "부산지사" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('부산','경남','울산')"
  elseif sido = "대구지사" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('대구','경북')"
  elseif sido = "대전지사" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('대전','충남','충북','세종')"
  elseif sido = "광주지사" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('광주','전남')"
  elseif sido = "전주지사" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('전북')"
  elseif sido = "제주지사" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('제주')"
  elseif sido = "강원지사" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('강원')"
	else
		sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
		sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and (sido = '" + sido + "')"
	end if	  
else
	if mg_ce = "총괄" then
		sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
		sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"')"
	  else
		sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
		sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and (mg_ce_id = '" + mg_ce_id + "')"
	end if
end if
Rs.Open Sql, Dbconn, 1

title_line = "기간별 미처리 현황"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>기간별 미처리 현황</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}

        </script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th><%=memo01%></th>
							  <td class="left"><%=memo02%></td>
							  <th>회사</th>
							  <td class="left"><%=company%></td>
							  <th>처리유형</th>
							  <td class="left"><%=as_type%></td>
							</tr>
                            <tr>
							  <th>기간</th>
							  <td class="left"><%=dis_days%></td>
							  <td colspan="4"><a href = "day_michulri_excel_request.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=sido%>&company=<%=company%>&as_type=<%=as_type%>&mg_ce=<%=mg_ce%>&mg_ce_id=<%=mg_ce_id%>&mg_group=<%=mg_group%>&days=<%=days%>" class="btnType04">엑셀다운로드</a>
							  </td>
					      	</tr>
						</tbody>
					</table>
					<br>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="15%" >
							<col width="5%" >
							<col width="18%" >
							<col width="25%" >
							<col width="*" >
							<col width="10%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">요청일자</th>
								<th scope="col">상태</th>
								<th scope="col">회사명</th>
								<th scope="col">부서명</th>
								<th scope="col">지역</th>
								<th scope="col">처리유형</th>
								<th scope="col">조회</th>
							</tr>
						</thead>
						<tbody>
						<%
                        seq = 0
                        do until rs.eof
                            seq = seq + 1
                            com_date = datevalue(mid(dateadd("h",10,rs("request_date")),1,10))
            '				com_date = datevalue(mid(rs("acpt_date"),1,10))
                            dd = datediff("d", com_date, curr_date)
            '				ddd = dd
                        '휴일 계산
                            if dd < 0 then
                                dd = 0 
                            end if
                            
                            if cstr(curr_day) = cstr(mid(rs("request_date"),1,10)) then
                                dd = 0
                            end if
            
                            if dd > 0 then
                                com_date = datevalue(mid(rs("request_date"),1,10))
                                'a = dd
                                a = datediff("d", com_date, curr_day)
                                'b = datepart("w", com_date)
                                'bb = datepart("w", curr_day)
                                'if bb = 1 then
                                '    a = a -1
                                'end if
                                'c = a + b
                                d = a
                                'if a > 1 then
                                '    if c > 7 then
                                '        d = a - 2
                                '    end if					 
                                'end if
                                
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
            '					com_date = datevalue(rs("acpt_date"))
                        '		act_date = com_date
                        
                                com_date = datevalue(mid(rs("request_date"),1,10))
                                
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
                                    acpt_hh = int(datepart("h",rs("request_date")))
                                    if acpt_hh > 13 and curr_hh < 12 then
                                        d = 0
                                    end if
                                end if
            ' 2012-02-06 end
                                dd = d
                                if d > 2 and d < 7 then
                                    dd = 3
                                end if
                                if d > 6 then
                                    dd = 7
                                end if
                              else
                        ' 휴일 계산 끝
                                dd = 0
                            end if
                            int date_len 
                    '		date_len=len(rs("acpt_date"))
                            dim len_date, hangle, bit01, bit02, bit03
                            acpt_date = rs("request_date")
                            len_date = len(acpt_date)
                            bit01 = left(acpt_date, 10)
                        ' 	bit01 = Replace(bit01,"-",".")
                            bit03 = left(right(acpt_date, 5), 2)
                            hangle = mid(acpt_date, 12, 2)
                            if len_date = 22 then
                                bit02 = mid(acpt_date, 15, 2)
                              else
                                bit02 = "0"&mid(acpt_date, 15, 1)
                            end If
                         
                            if hangle = "오후" and bit02 <> 12 then 
                                bit02 = bit02 + 12
                            end if
                            
                            date_to_date = bit01 & " " &bit02 & ":" & bit03
                            acpt_date = mid(date_to_date,3)
                            'acpt_date = replace(acpt_date,"-","/")
                            acpt_date = rs("request_date")
                            
                            if dd = days then
                                if rs("as_process") = "접수" then
                                    acpt_cnt = acpt_cnt + 1
                                end if
                                if rs("as_process") = "연기" then
                                    yun_cnt = yun_cnt + 1
                                end if
                                if rs("as_process") = "입고" then
                                    in_cnt = in_cnt + 1
                                end if
                                i = i + 1
                        %>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=acpt_date%></td>
								<td><%=rs("as_process")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%></td>
								<td><%=rs("as_type")%></td>
								<td><a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=rs("acpt_no")%>&win_sw=<%=win_sw%>','asview_pop','scrollbars=yes,width=800,height=700')">조회</a></td>
							</tr>
							<%
                                end if
                                rs.movenext()
                            loop
                            %>
						</tbody>
					</table>                    
					<br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>접수</th>
							  <td class="left"><%=acpt_cnt%></td>
							  <th>연기</th>
							  <td class="left"><%=yun_cnt%></td>
							  <th>입고</th>
							  <td class="left"><%=in_cnt%></td>
					      	</tr>
						</tbody>
					</table>
					<br>
				</form>
				</div>
			</div>
	</body>
</html>

