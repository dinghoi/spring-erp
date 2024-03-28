<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim ck_sw
dim company_tab(50)
Dim cowork_yn
dim page_cnt
dim pg_cnt

ck_sw=Request("ck_sw")
Page=Request("page")
dong = request("dong")
view_sort = request("view_sort")
cowork_yn = Request("cowork_yn")

If ck_sw = "y" Then
	view_c=Request("view_c")
	'cowork_yn = Request("cowork_yn")
  else
	view_c=Request.form("view_c")
	'cowork_yn = Request("cowork_yn")
End if

If view_c = "" Then
	view_c = "total"
	dong = ""
End If

If cowork_yn = "" Then
	cowork_yn = "A"
End If

'검색 조건 추가/검색 노출 권한 사용자[허정호_20210702]
Dim field_check, field_view, field_sql, search_level

Select Case user_id
	'허정호, 전창곤
	Case "102592", "101778"
		search_level = "Y"
	Case Else
		search_level  = "N"
End Select

field_check = Request("field_check")

If field_check = "total" Then
	field_check = ""
End If

If field_check <> "" Then
	field_view = Request("field_view")

	If field_check = "acpt_no" Then
		field_sql = " AND ( " & field_check & " = '" & field_view & "' ) "
	Else
		field_sql = " AND ( " & field_check & " LIKE '%" & field_view & "%' ) "
	End If
Else
	field_sql = ""
End If

be_pg = "as_list_ce.asp"
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

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

if c_grade = "7" then
	k = 0
'	Sql="select * from trade where use_sw = 'Y' and mg_group = '"+mg_group+"' and group_name = '"+user_name+"' order by trade_name asc"
	Sql = "SELECT * "&_
	      "  FROM trade "&_
	      " WHERE use_sw = 'Y' "&_
	      "   AND group_name = '"+user_name+"' "&_
	      " ORDER BY trade_name ASC"
	rs_trade.Open Sql, Dbconn, 1

	do until rs_trade.eof
		k = k + 1
		company_tab(k) = rs_trade("trade_name")
		rs_trade.movenext()
	loop
	rs_trade.close()
end if

if view_sort = "" then
	view_sort = "DESC"
end if

view_sql = " "
if view_c = "as" then
	view_sql = " AND (as_type = '방문처리' OR as_type = '원격처리') "
end if
if view_c = "inst" then
	view_sql = " AND (as_type <> '방문처리' AND as_type <> '원격처리') "
end if
order_Sql = " ORDER BY acpt_date " + view_sort

if view_c = "dong" then
	view_sql = " AND (dong like '%" + dong + "%' )"
	order_Sql = " ORDER BY sido, gugun, dong " + view_sort
end if

if view_c = "large" then
	view_sql = " and (large_paper_no <> '') "
	order_Sql = " ORDER BY large_paper_no, sido, gugun, dong " + view_sort
end if

if cowork_yn = "Y" then
	cowork_sql = " AND ( cowork_yn = 'Y' ) "
elseif  cowork_yn = "N" then
	cowork_sql = " AND ( cowork_yn = 'N' OR cowork_yn = ''  OR cowork_yn is null  ) "
else
	cowork_sql = ""
end if

'where_sql = " WHERE (mg_group = '" + mg_group + "') and "
base_sql = " WHERE (as_process = '접수' OR as_process = '입고' OR as_process = '연기' OR as_process = '대체입고') "
condi_sql = "  AND (mg_ce_id = '" + user_id + "') "
'if c_grade = "0" or c_grade = "1" then
'	condi_Sql = " "
'end if
if c_grade = "0" or ( c_grade = "1" and team = "본사팀" ) then
	condi_Sql = " "
end if

if ( c_grade = "1" and team <> "본사팀" ) Then
	If user_id = "100032" Then	'100032 정구일 차장 본인 내역만 확인 요청[20210630_허정호]
		condi_Sql = " AND mg_ce_id = '"&user_id&"'  "
	Else
		condi_Sql = " AND (team = '"&team&"' OR mg_ce_id = '"&user_id&"')  "
	End If
end if

if c_grade = "2" then
	if user_id = "100780" then ' 100780 장시철 이사람만 company 조건을 뺀다... 2018-12-06
		condi_Sql = " AND (mg_ce_id = '"+user_id+"') "
	else
		condi_Sql = " AND (company = '"+reside_company+"' OR mg_ce_id = '"+user_id+"') "
	end if
end If

if c_grade = "3"  and team <> "본사팀" then
	condi_Sql = " AND (team = '"+team+"' OR mg_ce_id = '"+user_id+"') "
end if
if c_grade = "3"  and team = "본사팀" then
	condi_Sql = "AND (mg_ce_id = '"+user_id+"') "
end if

if c_grade = "7" then
	com_sql = "company = '" + company_tab(1) + "'"
	for kk = 2 to k
		com_sql = com_sql + " OR company = '" + company_tab(kk) + "'"
	next
	where_sql = "WHERE "
	condi_Sql = " AND (" + com_sql + ") "
end if

if c_grade = "8" then
	where_sql = "WHERE "
	condi_Sql = " AND (company = '" + user_name + "') "
end if

if user_id = "102305" or user_id = "102306" then
sql = "SELECT count(*) FROM as_acpt A  where 1=1 AND A.acpt_man = '최성민'"
else
sql = "SELECT count(*) FROM as_acpt A  " + base_sql + cowork_sql + view_sql + condi_sql
end if

Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if user_id = "102305" or user_id = "102306" then
	sql = "SELECT * FROM as_acpt where 1=1 AND acpt_man = '최성민'"
else
	sql = "SELECT * FROM as_acpt "
	sql = sql & base_sql & cowork_sql & view_sql & condi_sql & field_sql & order_sql
	sql = sql & " LIMIT " & stpage & "," &pgsize
end If

Rs.Open Sql, Dbconn, 1

title_line = "나의 A/S 현황"
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
		<script src="/java/ui.js" type="text/javascript"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if (document.frm.view_c.value == ""){
					alert ("조회조건을 선택하시기 바랍니다");
					return false;
				}

				return true;
			}

			function condi_view(){
				if (eval("document.frm.view_c[0].checked || document.frm.view_c[1].checked || document.frm.view_c[2].checked || document.frm.view_c[3].checked"  )){
					document.getElementById('dong_view').style.display = 'none';
				}

				if (eval("document.frm.view_c[4].checked")){
					document.getElementById('dong_view').style.display = '';
				}
			}
		</script>
	</head>
	<body onLoad="condi_view();">
		<div id="wrap">
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/as_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_list_ce.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
						<dd>
							<p>
								<label>&nbsp; &nbsp; &nbsp;
									<strong>조회조건 : </strong>
                                    <input type="radio" name="view_c" value="total"	<% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">전체
                                    <input type="radio" name="view_c" value="as" 	<% if view_c = "as" then 	%>checked<% end if %> style="width:25px" onClick="condi_view()">A/S건
                                    <input type="radio" name="view_c" value="inst"	<% if view_c = "inst" then 	%>checked<% end if %> style="width:25px" onClick="condi_view()">설치공사이전외 기타
                                    <input type="radio" name="view_c" value="large"	<% if view_c = "large" then %>checked<% end if %> style="width:25px" onClick="condi_view()">대량건
                                    <input type="radio" name="view_c" value="dong"	<% if view_c = "dong" then 	%>checked<% end if %> style="width:25px" onClick="condi_view()">동별
								</label>
								<label>
									<input name="dong" type="text" value="<%=dong%>" style="width:70px; display:none" id="dong_view">
								</label>
								<br>
								<label>&nbsp; &nbsp;
                                </label>
								<label>
									<strong>업무유형 : </strong>
                                    <input type="radio" name="cowork_yn" value="A"  <% if cowork_yn = "A" then %>checked<% end if %> style="width:25px">전체
                                    <input type="radio" name="cowork_yn" value="N"  <% if cowork_yn = "N" then %>checked<% end if %> style="width:25px">일반
                                    <input type="radio" name="cowork_yn" value="Y"  <% if cowork_yn = "Y" then %>checked<% end if %> style="width:25px">협업
                                </label>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<%If search_level = "Y" Then %>
								<label>
                                    <strong>조건 : </strong>
                                    <select name="field_check" id="field_check" style="width:80px">
                                        <option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                        <option value="acpt_no" <% if field_check = "acpt_no" then %>selected<% end if %>>접수번호</option>
                                        <option value="mg_ce_id" <% if field_check = "mg_ce_id" then %>selected<% end if %>>담당CE ID</option>
                                        <option value="mg_ce" <% if field_check = "mg_ce" then %>selected<% end if %>>담당CE</option>
                                        <option value="acpt_man" <% if field_check = "acpt_man" then %>selected<% end if %>>접수자</option>
                                        <option value="acpt_user" <% if field_check = "acpt_user" then %>selected<% end if %>>사용자</option>
                                    </select>
									<input name="field_view" type="text" value="<%=field_view%>" style="width:80px;" id="field_view" >
                                </label>
								<%End If %>
								<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" />
							<col width="5%" />
							<col width="3%" />
							<col width="3%" />
							<col width="3%" />
							<col width="7%" />
							<col width="7%" />
							<col width="10%" />
							<col width="10%" />
							<col width="7%" />
							<col width="12%" />
							<col width="7%" />
							<col width="5%" />
							<col width="5%" />
							<col width="*" />
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">접수일자</th>
								<th scope="col">접수NO</th>
								<th scope="col">협업<br>여부</th>
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
								<th scope="col">장애내용</th>
							</tr>
						</thead>
						<tbody>
						<%
							do until rs.eof

								Dim len_date, hangle, bit01, bit02, bit03
								acpt_date = rs("acpt_date")
								len_date = len(acpt_date)
								bit01 = left(acpt_date, 10)								'bit01 = Replace(bit01,"-",".")
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

									'visit_date = rs("visit_date")
									com_date = acpt_date
									'act_date = com_date

									'휴일 계산 쿼리 조회 및 로직 변경[허정호_20210120]
									'do until (com_date > curr_date)
									'	sql_hol = "SELECT * FROM holiday WHERE holiday = '" + cstr(com_date) + "'"
									'	Set rs_hol=DbConn.Execute(SQL_hol)

									'	if rs_hol.eof or rs_hol.bof then
									'		d = d
									 ' else
									'		d = d -1
									'	end if

									'	com_date = dateadd("d",1,com_date)
									'	rs_hol.close()
									'Loop
									sql_hol = "SELECT COUNT(*) AS h_cnt FROM holiday WHERE holiday >= '"&com_date&"' AND holiday < '"&curr_date&"';"
									Set rs_hol = DBConn.Execute(sql_hol)
									If rs_hol.BOF Or rs_hol.EOF Then
										d = d
									Else
										d = CInt(d) - CInt(rs_hol("h_cnt"))
									End If

									rs_hol.Close()

									if d > 6 then
										hol_d = int(d/7) * 2
									end if
									'if d > 2 then
									'	d = 3
									'end if
									'if d = 1 then
									'	j = 5
									'elseif d = 2 then
									'	j = 6
									'else
									'	j = 7
									'end if
									d_day = d - hol_d
							  else
									d_day = 0
								end if
						' 휴일 계산 끝

								as_memo = replace(rs("as_memo"),chr(34),chr(39))
								view_memo = as_memo
								if len(as_memo) > 10 then
									view_memo = mid(as_memo,1,10) + ".."
								end if
						%>
							<tr>
								<td class="first"><%=mid(acpt_date,3)%></td>
								<td><%=rs("acpt_no")%></td>
								<td>
                                <% if (rs("cowork_yn")="Y") then %>
                                    <%="협업"%>
                                <% else %>
                                    <%="일반"%>
                                <% end if 	%>
                                </td>
								<td><%=rs("as_process")%></td>
								<td><%=d_day%></td>
								<td><%=rs("acpt_man")%>&nbsp;<%=rs("acpt_grade")%></td>
								<td>
                                <%
                                if rs("large_paper_no") = "" or isnull(rs("large_paper_no")) then
                                    %>
                                    <a href="as_result_reg.asp?acpt_no=<%=rs("acpt_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_c=<%=view_c%>&dong=<%=dong%>&view_sort=<%=view_sort%>"><%=rs("acpt_user")%>&nbsp;<%=rs("user_grade")%></a>
                                    <%
                                else
                                    %>
                                    <a href="#" onClick="pop_Window('large_result_reg.asp?acpt_no=<%=rs("acpt_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_c=<%=view_c%>&dong=<%=dong%>&view_sort=<%=view_sort%>','lage_result_reg_popup','scrollbars=yes,width=750,height=450')"><%=rs("acpt_user")%>&nbsp;<%=rs("user_grade")%></a>
                                    <%
                                end if
                                %>
                                </td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("tel_ddd")%>)<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%></td>
								<td><%=mid(cstr(rs("request_date")),3)%>&nbsp;<%=rs("request_time")%></td>
								<td><%=rs("mg_ce")%></td>
								<td><%=rs("as_type")%></td>
							  <td class="left"><p style="cursor:pointer"><span title="<%=as_memo%>"><%=view_memo%></span></p></td>
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
				    <td width="25%">
				    	<div class="btnCenter"><a href="excel_down_ce.asp?view_c=<%=view_c%>&cowork_yn=<%=cowork_yn%>&dong=<%=dong%>" class="btnType04">엑셀다운로드</a></div>
				   	</td>
				   	<td>
                        <div id="paging">
                            <a href="as_list_ce.asp?page=<%=first_page%>&view_c=<%=view_c%>&cowork_yn=<%=cowork_yn%>&dong=<%=dong%>&view_sort=<%=view_sort%>&ck_sw=<%="y"%>">[처음]</a>
                            <%
                            if intstart > 1 then
                                Response.write "<a href='as_list_ce.asp?page="&intstart-1&"&view_c="&view_c&"&cowork_yn="&cowork_yn&"&dong="&dong&"&view_sort="&view_sort&"&ck_sw=y'>[이전]</a>"
                            end if

                            for i = intstart to intend
                            %>
                                <% if i = int(page) then %>
                                    <b>[<%=i%>]</b>
                                <% else %>
                                    <a href="as_list_ce.asp?page=<%=i%>&view_c=<%=view_c%>&cowork_yn=<%=cowork_yn%>&dong=<%=dong%>&view_sort=<%=view_sort%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                                <% end if %>
                            <% next %>
                            <% if intend < total_page then %>
                                <a href="as_list_ce.asp?page=<%=intend+1%>&view_c=<%=view_c%>&cowork_yn=<%=cowork_yn%>&dong=<%=dong%>&view_sort=<%=view_sort%>&ck_sw=<%="y"%>">[다음]</a> <a href="as_list_ce.asp?page=<%=total_page%>&view_c=<%=view_c%>&cowork_yn=<%=cowork_yn%>&dong=<%=dong%>&view_sort=<%=view_sort%>&ck_sw=<%="y"%>">[마지막]</a>
                            <% else %>
                                [다음]&nbsp;[마지막]
                            <% end if %>
                    </div>
                    </td>
				    <td width="25%">
				    	<div class="btnCenter">
                        <% if view_sort = "DESC" then	%>
                            <a href="as_list_ce.asp?page=<%=page%>&view_c=<%=view_c%>&cowork_yn=<%=cowork_yn%>&dong=<%=dong%>&view_sort=<%="ASC"%>&ck_sw=<%="y"%>" class="btnType04">정순조회</a>
                        <% else %>
                            <a href="as_list_ce.asp?page=<%=page%>&view_c=<%=view_c%>&cowork_yn=<%=cowork_yn%>&dong=<%=dong%>&view_sort=<%="DESC"%>&ck_sw=<%="y"%>" class="btnType04">역순조회</a>
                        <% end if %>
					    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>
