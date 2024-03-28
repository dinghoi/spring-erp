<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim company_tab(150)
dim com_tab(300)
dim ce_id(300)
dim com_sum(300)
dim ok_sum(300)
dim mi_sum(300)
dim dang_acpt(300)
dim dang_cnt(300)
dim com_cnt(300,9)
dim sum_cnt(9)
dim mob_end(300)
dim mob_mi(300)
dim team_tab(100)

from_date=Request.form("from_date")
to_date=Request.form("to_date")
as_type=Request.form("as_type")
company=Request.form("company")
team=Request.form("team")
If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	as_type = "방문처리"
	company = "전체"
	team = "전체"
End If


for i = 0 to 300
	com_tab(i) = ""
	com_sum(i) = 0
	ok_sum(i) = 0
	mi_sum(i) = 0
	dang_cnt(i) = 0
	dang_acpt(i) = 0
	mob_end(i) = 0
	mob_mi(i) = 0
	for j = 0 to 9
		com_cnt(i,j) = 0
	next
next
for j = 0 to 9
	sum_cnt(j) = 0
next
for i = 0 to 20
	team_tab(i) = ""
next

curr_day = datevalue(mid(cstr(now()),1,10))
curr_date = datevalue(mid(dateadd("h",12,now()),1,10))

sql = "select team from as_acpt Where (Cast(acpt_date as date) >= '" + from_date + "' and Cast(acpt_date as date) <= '"+to_date+"') GROUP BY team Order By team Asc"
i = 0
Rs.Open Sql, Dbconn, 1
do until rs.eof
	if rs("team") <> "" then
		i = i + 1
		team_tab(i) = rs("team")
	  else
		i = i + 1
		team_tab(i) = "미지정"	  
	end if
	rs.movenext()
loop
rs.close()


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
if team = "전체" then
	team_sql = ""
  else
  	team_sql = " (team ='"+team+"') and "
end if

' 완료건
sql = "select mg_ce, mg_ce_id, as_process, CAST(acpt_date as date) as acpt_day, CAST((acpt_date + interval 10 DAY_HOUR) as date) as com_date, visit_date, substring(visit_time,1,2) as visit_hh, count(*) as err_cnt from as_acpt"
sql = sql + " where "+com_sql+type_sql+team_sql+" (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
sql = sql + " GROUP BY mg_ce, mg_ce_id, as_process, CAST(acpt_date as date), CAST((acpt_date + interval 10 DAY_HOUR) as date) , visit_date, substring(visit_time,1,2) Order By mg_ce, mg_ce_id Asc"
Rs.Open Sql, Dbconn, 1
first_sw = "y"
'bi_mg_ce = ""
i = 0
do until rs.eof

	if isnull(rs("mg_ce_id")) or rs("mg_ce_id") = "" then
		mg_ce_id = "오류"
	  else
		mg_ce_id = rs("mg_ce_id")
		mg_ce = rs("mg_ce")
	end if
	if firsr_sw = "y" then
		bi_mg_ce = mg_ce_id
		first_sw = "n"
		mg_ce = rs("mg_ce")
	end if
	
	if bi_mg_ce <> mg_ce_id then
		bi_mg_ce = mg_ce_id
		mg_ce = rs("mg_ce")
		i = i + 1	
	end if

	com_tab(i) = mg_ce
	ce_id(i) = mg_ce_id
	if	rs("as_process") = "완료" or rs("as_process") = "취소" or rs("as_process") = "대체" then
		as_process = "완료"		

	  	visit_date = datevalue(rs("visit_date"))
' 1/19 추가
	  	visit_day = datevalue(rs("visit_date"))
' 1/19 추가 end

		if cstr(rs("visit_hh")) > "12" then
			visit_date = dateadd("d",1,visit_date)
		end if
		
		dd = datediff("d", rs("com_date"), visit_date)

		if	visit_day = datevalue(to_date) then
			dang_cnt(i) = dang_cnt(i) + cint(rs("err_cnt"))
		end if

		if cstr(visit_day) = cstr(rs("acpt_day")) then
			dd = 0
		end if
	  else
		as_process = "미처리"
		dd = datediff("d", rs("com_date"), curr_date)
		if cstr(curr_day) = cstr(rs("acpt_day")) then
			dd = 0
		end if
	end if

	if cstr(rs("acpt_day")) = cstr(to_date) then
		dang_acpt(i) = dang_acpt(i) + cint(rs("err_cnt"))
	end if

	if dd < 0 then
		dd = 0 
	end if

'휴일 계산
	if dd > 0 then
		if	as_process = "미처리" then	
'			visit_date = rs("visit_date")
'		  else
		  	visit_day = curr_day
		end if
		a = dd
		a = datediff("d", rs("acpt_day"), visit_day)
		b = datepart("w",rs("acpt_day"))		
		if as_process = "미처리" then
			bb = datepart("w", curr_day)
			if bb = 1 then
				a = a -1
			end if
		end if
		c = a + b
		d = a
		if a > 1 then
			if c > 7 then
				d = a - 2
			end if
		end if
			
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
'		if d > 2 then
'			d = 3
'		end if
		if	as_process = "완료" then
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
			com_cnt(i,d) = com_cnt(i,d) + cint(rs("err_cnt"))	
		  else
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
			com_cnt(i,j) = com_cnt(i,j) + cint(rs("err_cnt"))	
		end if			  	
	  else
	
' 휴일 계산 끝
		if	as_process = "완료" then
			com_cnt(i,0) = com_cnt(i,0) + cint(rs("err_cnt"))
		  else
			com_cnt(i,5) = com_cnt(i,5) + cint(rs("err_cnt"))
		end if				
	end if
	rs.movenext()
loop
rs.close()

title_line = "CE별 기간별 처리현황"
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
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=ce_term_pro.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<strong>회사</strong>
							  	<%
								sql="select * from trade where (use_sw = 'Y') and (mg_group = '1' or mg_group = '2') and (trade_id = '매출' or trade_id = '공통') order by trade_name asc"
                                rs_trade.Open Sql, Dbconn, 1
                                %>
								<label>
        						<select name="company" id="company">
									<option value="전체">전체</option>
          					<% 
								do until rs_trade.eof 
							%>
          							<option value='<%=rs_trade("trade_name")%>' <%If rs_trade("trade_name") = company  then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
          					<%
									rs_trade.movenext()  
								loop 
								rs_trade.Close()
							%>
        						</select>
                                </label>
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
								<strong>소속</strong>
								<% 
                                    Sql="select * from etc_code where etc_type = '62' and used_sw = 'Y' order by etc_name asc"
                                    Rs_etc.Open Sql, Dbconn, 1
                                %>
                                <select name="team" id="team" style="width:150px">
                                  <option value="전체" <%If team = "전체" then %>selected<% end if %>>전체</option>
                                <%
                                    for i = 1 to 20
 										if team_tab(i) = "" then
											exit for
										end if
                                %>
                                  <option value='<%=team_tab(i)%>' <%If team_tab(i) = team then %>selected<% end if %>><%=team_tab(i)%></option>
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
							<col width="6%" >
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">담당CE</th>
								<th scope="col" rowspan="2">소속</th>
								<th scope="col" rowspan="2">금일처리</th>
								<th scope="col" rowspan="2">당일접수</th>
								<th scope="col" colspan="6" style=" border-bottom:1px solid #e3e3e3;">처리완료</th>
								<th scope="col" colspan="6" style=" border-bottom:1px solid #e3e3e3;">미처리</th>
								<th scope="col" rowspan="2">CE 계</th>
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
                            dang_sum = 0
                            dang_acpt_sum = 0
                            sum_mob_end = 0
                            sum_mob_mi = 0
                            for i = 0 to 300 
                                if	com_tab(i) <> "" then
                
                                    for j = 0 to 4
                                        ok_sum(i) = ok_sum(i) + com_cnt(i,j)
                                        sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)				
                                    next
                                    for j = 5 to 9
                                        mi_sum(i) = mi_sum(i) + com_cnt(i,j)
                                        sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)				
                                    next
                                    com_sum(i) = ok_sum(i) + mi_sum(i)
                                    dang_sum = dang_sum + dang_cnt(i)	
                                    dang_acpt_sum = dang_acpt_sum + dang_acpt(i)
                                    sido = com_tab(i)
								end if
							next                						
                        %>
							<tr>
                              <th colspan="2">총계</th>
                              <th class="right"><%=formatnumber(clng(dang_sum),0)%></th>
                              <th class="right"><%=formatnumber(clng(dang_acpt_sum),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(0)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(1)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(2)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(3)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(4)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(5)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(6)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(7)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(8)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(9)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)),0)%></th>
                              <th class="right"><%=formatnumber(clng(sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)),0)%></th>
							</tr>
						 <% 	
                            for i = 0 to 300 
                                if	com_tab(i) <> "" then                
									sql = "select * from memb where user_id = '" + ce_id(i) + "'"
									Set rs_memb=DbConn.Execute(SQL)
									if rs_memb.eof or rs_memb.bof then
										team_view = "퇴직자"
									  else
										team_view = rs_memb("team")
									end if
                        %>
							<tr>
                              <td><%=com_tab(i)%></td>
                              <td><%=team_view%></td>
                              <td class="right"><%=formatnumber(clng(dang_cnt(i)),0)%></td>
                              <td class="right"><%=formatnumber(clng(dang_acpt(i)),0)%></td>
                              <td class="right"><%=formatnumber(clng(com_cnt(i,0)),0)%></td>
                              <td class="right"><%=formatnumber(clng(com_cnt(i,1)),0)%></td>
                              <td class="right"><%=formatnumber(clng(com_cnt(i,2)),0)%></td>
                              <td class="right"><%=formatnumber(clng(com_cnt(i,3)),0)%></td>
                              <td class="right"><%=formatnumber(clng(com_cnt(i,4)),0)%></td>
                              <td class="right"><%=formatnumber(clng(ok_sum(i)),0)%></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&mg_ce=<%=com_tab(i)%>&mg_ce_id=<%=ce_id(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=0%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,5)),0)%></a></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&mg_ce=<%=com_tab(i)%>&mg_ce_id=<%=ce_id(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,6)),0)%></a></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&mg_ce=<%=com_tab(i)%>&mg_ce_id=<%=ce_id(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,7)),0)%></a></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&mg_ce=<%=com_tab(i)%>&mg_ce_id=<%=ce_id(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,8)),0)%></a></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&mg_ce=<%=com_tab(i)%>&mg_ce_id=<%=ce_id(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=7%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,9)),0)%></a></td>
                              <td class="right"><a  href="#" onClick="pop_Window('as_michulri_popup.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&mg_ce=<%=com_tab(i)%>&mg_ce_id=<%=ce_id(i)%>&company=<%=company%>&as_type=<%=as_type%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(mi_sum(i)),0)%></a></td>
                              <td class="right"><%=formatnumber(clng(com_sum(i)),0)%></td>
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

