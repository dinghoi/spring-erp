<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim company_tab(150)

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	replace_sw=Request("replace_sw")
	company=Request("company")
Else
	replace_sw=Request.form("replace_sw")
	company=Request.form("company")
End if

if replace_sw = "" then
	replace_sw = "전체"
	company = "전체"
end if

be_pg = "into_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

replace_sql = ""
if replace_sw <> "전체" then
	if replace_sw = "대체" then
		replace_sql = " and (in_replace = '대체')"
	  else
	  	replace_sql = " and (in_replace <> '대체')"
	end if
end if

company_sql = ""
if company <> "전체" then
	company_sql = " and (company = '"+company+"')"
end if
'where_sql = " WHERE (mg_group = '" + mg_group + "') and (as_process = '입고') "
where_sql = " WHERE (as_process = '입고') "
order_sql = " ORDER BY acpt_date ASC"
condi_Sql = " and (mg_ce_id = '" + user_id + "')"

'c_grade=1
'Response.write "c_grade:" & c_grade & "<br>"
'Response.write "team:" & team & "<br>"

if c_grade = "0" or ( c_grade = "1" and team = "본사팀" ) then
	condi_Sql = " "
End If

If c_grade = "1" And team <> "본사팀" Then
	condi_Sql = " and (team = '"&team&"' or mg_ce_id = '"&user_id&"')"
	'Select Case emp_no
		'Case "100064", "102419" '김영진A, 양동평
			'condi_sql = "AND (team='' OR team = '"&team&"' OR mg_ce_id = '"&emp_no&"') "
	''		condi_sql = "AND (saupbu='호남사업부' OR team = '"&team&"' OR mg_ce_id = '"&emp_no&"') "
	''	Case Else
	''		condi_Sql = "AND (team = '"&team&"' OR mg_ce_id = '"&emp_no&"') "
	'End Select
End If

if c_grade = "2" then
	condi_Sql = " and (company = '"+reside_company+"' or mg_ce_id = '"+user_id+"') "
end if
if c_grade = "3"  and team <> "본사팀" then
	condi_Sql = " and (team = '"+team+"' or mg_ce_id = '"+user_id+"') "
end if
if c_grade = "3"  and team = "본사팀" then
	Sql = " and (mg_ce_id = '"+user_id+"') "
end if

Sql = "SELECT count(*) FROM as_acpt " + where_sql + condi_sql + replace_sql + company_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from as_acpt " + where_sql + condi_sql + replace_sql + company_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1
' 김승호 101247 grade 1 -> 0 으로 수정 2019.06.20
'Response.write sql & "<br>"

title_line = "입고 진행 현황"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}

		</script>

	</head>
	<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">-->
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/as_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="into_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>대체유무 : </strong>
                                <select name="replace_sw" id="replace_sw">
                                  <option value="전체" <% if replace_sw = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="대체" <% if replace_sw = "대체" then %>selected<% end if %>>대체</option>
                                  <option value="미대체" <% if replace_sw = "미대체" then %>selected<% end if %>>미대체</option>
                                </select>
								</label>
								<label>
								<strong>회사 : </strong>
                                    <%
                                    rs_trade.Open sql_trade, Dbconn, 1
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
 								</label>
                               <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="15%" >
							<col width="*" >
							<col width="6%" >
							<col width="9%" >
							<col width="8%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">경과</th>
								<th scope="col">접수일자</th>
								<th scope="col">입고일자</th>
								<th scope="col">고객명</th>
								<th scope="col">회사</th>
								<th scope="col">조직명</th>
								<th scope="col">담당CE</th>
								<th scope="col">제조사</th>
								<th scope="col">입고장비</th>
								<th scope="col">입고처</th>
								<th scope="col">대체</th>
								<th scope="col">최종처리</th>
								<th scope="col">진행</th>
							</tr>
						</thead>
						<tbody>
						<%
                    	do until rs.eof

					'휴일 계산
							hol_d = 0
							com_date = datevalue(mid(rs("acpt_date"),1,10))
							dd = datediff("d", com_date, curr_date)
							if dd > 0 then
								a = datediff("d", com_date, curr_date)
								b = datepart("w",com_date)
								c = a + b
								d = a
								if a > 1 then
									if c > 7 then
										d = a - 2
									end if
								end if

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
								d_day = d - hol_d
							  else
						' 휴일 계산 끝
								d_day = 0
							end if

							sql = "select into_date,in_process,in_place from as_into where acpt_no="&rs("acpt_no")&" and in_seq="&"(select max(in_seq) from as_into where acpt_no="&rs("acpt_no")&")"
							Set rs_in=dbconn.execute(sql)
							if	rs_in.eof then
									into_date = "없음"
									in_place = "없음"
									in_process = "없음"
								else
									into_date = rs_in("into_date")
									in_place = rs_in("in_place")
									in_process = rs_in("in_process")
							end if

							if rs("in_replace") = "" or isnull(rs("in_replace")) then
								in_replace = "."
							  else
								in_replace = rs("in_replace")
							end if
                    		%>
							<tr>
								<td class="first"><span style="color:#F60; font-weight:bold"><%=d_day%>일</span></td>
								<td><%=mid(rs("acpt_date"),1,10)%></td>
								<td><%=rs("in_date")%></td>
								<td><a href="as_result_reg.asp?acpt_no=<%=rs("acpt_no")%>&be_pg=<%=be_pg%>"><%=rs("acpt_user")%></a></td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("mg_ce")%></td>
								<td><%=rs("maker")%></td>
								<td><%=rs("as_device")%></td>
								<td><%=in_place%></td>
								<td>
                                <% if in_replace = "대체" then %>
              						<span style="color:#090; font-weight:bold"><%=in_replace%></span>
              					<% else %>
              						<%=in_replace%>
              					<% end if %>
                                </td>
								<td>
								<% if in_process = "수리완료" then %>
                                	<span style="color:#006; font-weight:bold"><%=in_process%></span>
                                <% else %>
                                	<%=in_process%>
                                <% end if %>
                                </td>
							  	<td><a href="#" onClick="pop_Window('into_mg.asp?acpt_no=<%=rs("acpt_no")%>','into_pop','scrollbars=yes,width=900,height=600')">입력</a></td>
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
                    <a href="excel_down_into.asp?replace_sw=<%=replace_sw%>&company=<%=company%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="into_list.asp?page=<%=first_page%>&ck_sw=<%="y"%>&end_sw=<%=end_sw%>&replace_sw=<%=replace_sw%>&company=<%=company%>">[처음]</a>
                        <% if intstart > 1 then %>
                            <a href="into_list.asp?page=<%=intstart -1%>&ck_sw=<%="y"%>&end_sw=<%=end_sw%>&replace_sw=<%=replace_sw%>&company=<%=company%>">[이전]</a>
                        <% end if %>
                        <% for i = intstart to intend %>
                            <% if i = int(page) then %>
                                <b>[<%=i%>]</b>
                            <% else %>
                                <a href="into_list.asp?page=<%=i%>&ck_sw=<%="y"%>&end_sw=<%=end_sw%>&replace_sw=<%=replace_sw%>&company=<%=company%>">[<%=i%>]</a>
                            <% end if %>
                        <% next %>
                        <% if intend < total_page then %>
                            <a href="into_list.asp?page=<%=intend+1%>&ck_sw=<%="y"%>&end_sw=<%=end_sw%>&replace_sw=<%=replace_sw%>&company=<%=company%>">[다음]</a> <a href="into_list.asp?page=<%=total_page%>&ck_sw=<%="y"%>&end_sw=<%=end_sw%>&replace_sw=<%=replace_sw%>&company=<%=company%>">[마지막]</a>
                        <%	else %>
                            [다음]&nbsp;[마지막]
                        <% end if %>
                    </div>
                    </td>
				    <td width="15%">
                    </td>
			      </tr>
				</table>
			</form>
		</div>
	</div>
	</body>
</html>

