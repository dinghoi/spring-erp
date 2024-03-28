<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim company_tab(50)
dim page_cnt
dim pg_cnt
Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "large_list_ce.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

if page_cnt > 0 then 
	pg_cnt = page_cnt
end if
if pg_cnt > 0 then
	page_cnt = pg_cnt
end if

if page_cnt < 10 or page_cnt > 20 then
	page_cnt = 10
end if

pgsize = page_cnt ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
'  else
'  	page = cint(page)
'	start_page = int(page/setsize)
'	if start_page = (page/setsize) then
'		strat_page = page - setsize + 1
'	  else
'	  	start_page = int(page/setsize)*setsize + 1
'	end if
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

if c_grade = "7" then
	k = 0
	Sql="select * from etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' and group_name = '"+user_name+"' order by etc_name asc"
	rs_etc.Open Sql, Dbconn, 1
	while not rs_etc.eof
		k = k + 1
		company_tab(k) = rs_etc("etc_name")
		rs_etc.movenext()
	Wend
	rs_etc.close()						
end if

view_sort = request("view_sort")

if view_sort = "" then
	view_sort = "DESC"
end if
order_Sql = " ORDER BY acpt_date " + view_sort

where_sql = " WHERE (mg_group = '" + mg_group + "') and "
base_sql = " (as_process = '접수' or as_process = '입고' or as_process = '연기' or as_process = '대체입고') "
condi_sql = " and (mg_ce_id = '" + c_id + "') "
if c_grade = "0" or ( c_grade = "1" and c_belong = "수도권지사" ) then
	condi_Sql = " "
end if	
if ( c_grade = "1" and c_belong <> "수도권지사" ) then
	condi_Sql = " and (belong = '"+c_belong+"' or mg_ce_id = '"+c_id+"') "
end if	
if c_grade = "2" then
	condi_Sql = " and (reside_place = '"+reside_place+"' or mg_ce_id = '"+c_id+"') "
end if
if c_grade = "3"  and c_belong <> "수도권지사" then
	condi_Sql = " and (belong = '"+c_belong+"' or mg_ce_id = '"+c_id+"') "
end if
if c_grade = "3"  and c_belong = "수도권지사" then
	condi_Sql = "and (mg_ce_id = '"+c_id+"') "
end if

if c_grade = "7" then
	com_sql = "company = '" + company_tab(1) + "'"	
	for kk = 2 to k
		com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
	next
	where_sql = "WHERE "
	condi_Sql = " and (" + com_sql + ") "
end if

if c_grade = "8" then
	where_sql = "WHERE "
	condi_Sql = " and (company = '" + user_name + "') "
end if

Sql = "SELECT count(*) FROM large_acpt " + where_sql + base_sql + condi_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from large_acpt " + where_sql + base_sql + condi_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "담당자별 대량건"

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
		<script src="/java/ui.js" type="text/javascript"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
//				if (formcheck(document.frm) && chkfrm()) {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				alert("aaaa");
//				if (document.frm.condi.value == "") {
//					alert ("소속을 선택하시기 바랍니다");
//					return false;
//				}	
//				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/large_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="large_list_ce_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="3%" >
							<col width="3%" >
							<col width="7%" >
							<col width="10%" >
							<col width="12%" >
							<col width="12%" >
							<col width="8%" >
							<col width="*" >
							<col width="8%" >
							<col width="5%" >
							<col width="6%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">접수일자</th>
								<th scope="col">상태</th>
								<th scope="col">경과</th>
								<th scope="col">접수자</th>
								<th scope="col">사용자</th>
								<th scope="col">회사</th>
								<th scope="col">조직명</th>
								<th scope="col">전화번호</th>
								<th scope="col">지역</th>
								<th scope="col">요청일자</th>
								<th scope="col">담당CE</th>
								<th scope="col">처리유형</th>
								<th scope="col">완료</th>
							</tr>
						</thead>
						<tbody>
						<%
						j = 0
						do until rs.eof
							j = j +1
							dim len_date, hangle, bit01, bit02, bit03
							acpt_date = rs("acpt_date")
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
						
						'휴일 계산
							hol_d = 0
							acpt_date = datevalue(mid(rs("acpt_date"),1,10))
							dd = datediff("d", acpt_date, curr_date)
							if dd > 0 then
								a = datediff("d", acpt_date, curr_date)
								b = datepart("w",acpt_date)
								c = a + b
								d = a
								if a > 1 then
									if c > 7 then
										d = a - 2
									end if
								end if
									
						'		visit_date = rs("visit_date")
								com_date = acpt_date
						'		act_date = com_date
								
								do until com_date > curr_date
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
								if d > 6 then
									hol_d = int(d/7) * 2
								end if
					'			if d > 2 then
					'				d = 3
					'			end if
					'			if d = 1 then
					'				j = 5
					'			  elseif d = 2 then
					'				j = 6
					'			  else
					'				j = 7
					'			end if
								d_day = d - hol_d
							  else
								d_day = 0
							end if
						' 휴일 계산 끝						
							as_memo = replace(rs("as_memo"),chr(34),chr(39))
							view_memo = as_memo
							if len(as_memo) > 15 then
								view_memo = mid(as_memo,1,15) + ".."
							end if
						%>
							<tr>
								<td class="first"><%=acpt_date%></td>
								<td><%=rs("as_process")%></td>
								<td><%=d_day%></td>
								<td><%=rs("acpt_man")%>&nbsp;<%=rs("acpt_grade")%></td>
								<td><%=rs("acpt_user")%>&nbsp;<%=rs("user_grade")%></a></td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("tel_ddd")%>)<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%></td>
								<td><%=mid(cstr(rs("request_date")),3)%>&nbsp;<%=rs("request_time")%></td>
								<td><%=rs("mg_ce")%></td>
								<td><%=rs("as_type")%></td>
								<td><a href="#" onClick="pop_Window('large_result_reg.asp?acpt_no=<%=rs("acpt_no")%>','lage_result_reg_popup','scrollbars=yes,width=750,height=350')">등록</a></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
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
				    <td width="15%">
					<div class="btnCenter">
                    <a href="excel_down_ce.asp" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="as_list_ce.asp?page=<%=first_page%>&view_sort=<%=view_sort%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="as_list_ce.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="as_list_ce.asp?page=<%=i%>&view_sort=<%=view_sort%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="as_list_ce.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>">[다음]</a> <a href="as_list_ce.asp?page=<%=total_page%>&view_sort=<%=view_sort%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
					<div class="btnCenter">
						<a href="#" class="btnType04" onclick="javascript:frmcheck();">등록</a>
                    </div>                  
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

