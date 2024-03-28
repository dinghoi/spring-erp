<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
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

if mg_ce = "" then
	title_memo = sido + "_지역별_"
  else
    title_memo = mg_ce + "_담당자_"
end if
savefilename = title_memo + cstr(days) + "일 요청일자 기준 미처리 내역.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
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

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title></title>
<style type="text/css">
<!--
.style14 {color: #FFCCFF}
-->
</style>
</head>
<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0">
  <tr bgcolor="#CCCCCC" class="style11">
    <td height="25" bgcolor="#FFFFFF"><%=memo01%></td>
    <td height="25" bgcolor="#FFFFFF">&nbsp;<%=memo02%></td>
    <td height="25" bgcolor="#FFFFFF">회사</td>
    <td height="25" bgcolor="#FFFFFF">&nbsp;<%=company%></td>
    <td height="25" bgcolor="#FFFFFF">처리유형</td>
    <td height="25" bgcolor="#FFFFFF">&nbsp;<%=as_type%></td>
    <td height="25" bgcolor="#FFFFFF">기간</td>
    <td bgcolor="#FFFFFF"><%=days%>일 미처리</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF">접수일자 기준</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF" class="style11">
    <td width="88"><div align="center"><strong>접수일자</strong></div></td>
    <td width="57" height="20"><div align="center"><strong><span class="style25">접수자</span></strong></div></td>
    <td width="56" height="20"><div align="center"><strong>사용자</strong></div></td>
    <td width="101" height="20" class="style11B"><div align="center"><strong>전화번호</strong></div></td>
    <td width="102" height="20" class="style11B"><div align="center"><strong>회사</strong></div></td>
    <td width="101" height="20" class="style11B"><div align="center"><strong>조직명</strong></div></td>
    <td width="165" height="20"><div align="center"><strong>주소</strong></div></td>
    <td width="63"><div align="center"><strong>CE명</strong></div></td>
    <td width="110"><div align="center"><strong>장애내역</strong></div></td>
    <td width="64"><div align="center"><strong>요청일</strong></div></td>
    <td width="57"><div align="center"><strong>요청시간</strong></div></td>
    <td width="56"><div align="center"><strong>처리방법</strong></div></td>
    <td width="38"><div align="center"><strong>진행</strong></div></td>
    <td width="55"><div align="center"><strong>입고사유</strong></div></td>
  </tr>
  <%
			seq = 0
			do until rs.eof
				seq = seq + 1
				com_date = datevalue(mid(dateadd("h",10,rs("request_date")),1,10))
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
					'bb = datepart("w", curr_date)
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
' 1/19 추가
					acpt_day = datevalue(mid(rs("request_date"),1,10))
					ddd = datediff("d", acpt_day, curr_day)
					if d > ddd then
						d = ddd
					end if
' 1/19 추가 end
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
					'if d > 2 and d < 7 then
					'	dd = 3
					'end if
					'if d > 6 then
					'	dd = 7
					if d > 4 then
						dd = 5
					end if
				  else
			' 휴일 계산 끝
					dd = 0
				end if
				
				if dd = days then
			%>
			  <tr valign="middle" class="style11">
				<td width="88" height="20" class="style11"><div align="center"><%=rs("acpt_date")%></div></td>
				<td width="57" height="20" class="style11"><div align="center" class="style25"><%=rs("acpt_man")%></div></td>
				<td width="56" height="20" class="style11"><div align="center" class="style25"><%=rs("acpt_user")%></div></td>
				<td width="101" height="20" class="style11"><div align="center"><%=replace(rs("tel_ddd")," ","")%>-<%=replace(rs("tel_no1")," ","")%>-<%=replace(rs("tel_no2")," ","")%></div></td>
				<td width="102" height="20" class="style11"><div align="center"><%=rs("company")%></div></td>
				<td width="101" height="20" class="style11"><div align="center"><%=rs("dept")%></div></td>
				<td width="165" height="20"><div align="center"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>&nbsp;<%=rs("addr")%></div></td>
			    <td width="63"><div align="center"><%=rs("mg_ce")%></div></td>
			    <td width="110"><div align="center"><%=rs("as_memo")%></div></td>
			    <td width="64"><div align="center"><%=rs("request_date")%></div></td>
			    <td width="57"><div align="center"><%=rs("request_time")%></div></td>
			    <td width="56"><div align="center"><%=rs("as_type")%></div></td>
			    <td width="38"><div align="center"><%=rs("as_process")%></div></td>
			    <td width="55"><div align="center"><%=rs("into_reason")%></div></td>
			  </tr>
  			<%
				end if
				rs.movenext()
			loop
			%>
</table>
</body>
</html>
<%
rs.close()
dbconn.Close()
Set dbconn = Nothing
%>
