<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'on Error resume next
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim company_tab(50)
Dim Rs, Repeat_Rows, page_cnt, pg_cnt
Dim page, be_pg, curr_date, pgsize, start_page, stpage
Dim view_sort, condi_sql, order_sql, where_sql
Dim rsCount, total_record, total_page, title_line
Dim base_sql, str_param

page = f_Request("page")
page_cnt = f_Request("page_cnt")
pg_cnt = CInt(f_Request("pg_cnt"))
view_sort = f_Request("view_sort")

title_line = "나의 A/S 현황"
be_pg = "/service/as_list_ce_user.asp"
curr_date = DateValue(Mid(CStr(Now()), 1, 10))

If page_cnt > 0 Then
	pg_cnt = page_cnt
End If

If pg_cnt > 0 Then
	page_cnt = pg_cnt
End If

If page_cnt < 10 Or page_cnt > 20 Then
	page_cnt = 10
End If

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

stpage = Int((page - 1) * pgsize)

str_param = "&view_sort="&view_sort

If reside = "9" Then
	k = 0

	'Sql="select * from trade where use_sw = 'Y' and group_name = '"+user_name+"' order by trade_name asc"
	objBuilder.Append "SELECT trade_name FROM trade "
	objBuilder.Append "WHERE use_sw = 'Y' AND group_name = '' "
	objBuilder.Append "ORDER BY trade_name ASC "

	Set rs_trade = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Do Until rs_trade.EOF
		k = k + 1
		company_tab(k) = rs_trade("trade_name")
		rs_trade.MoveNext()
	Loop
	rs_trade.Close() : Set rs_trade = Nothing
End If

If reside = "9" Then
	com_sql = "company = '" & company_tab(1) & "' "

	For kk = 2 To k
		com_sql = com_sql & " OR company = '" & company_tab(kk) & "' "
	Next

	condi_sql = " OR " & com_sql & ") "
Else
	condi_sql = " OR company = '" & reside_company & "' OR company = '" & user_name & "') "
End If

'//2017-06-07 아이티퓨처(사번:900002) 로그인시 웅진관련 기업 검색하게 수정
If  user_id = "900002" Then
	condi_sql = " OR company IN ('웅진식품', '웅진씽크빅', '코웨이') " & condi_sql
End IF

If view_sort = "" Then
	view_sort = "DESC"
End If

order_Sql = " ORDER BY acpt_date " & view_sort

where_sql = " WHERE (acpt_man = '" + user_name + "'" + condi_sql
base_sql = " and (as_process = '접수' OR as_process = '입고' OR as_process = '연기' OR as_process = '대체입고') "

'Sql = "SELECT count(*) FROM as_acpt " + where_sql + base_sql
objBuilder.Append "SELECT COUNT(*) FROM as_acpt "&where_sql&base_sql

Set RsCount = Dbconn.Execute (objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(RsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

'objBuilder.Append "select a.* "
objBuilder.Append "SELECT acpt_date, as_memo, as_process, acpt_man, acpt_grade, "
objBuilder.Append "	acpt_no, acpt_user, user_grade, company, dept, tel_ddd, "
objBuilder.Append "	tel_no1, tel_no2, mg_ce, as_type, sido, gugun, "
objBuilder.Append "	request_date, request_time, "
objBuilder.Append " (SELECT CONCAT(emp_hp_ddd ,'-', emp_hp_no1, '-', emp_hp_no2) "
objBuilder.Append "	FROM emp_master WHERE emp_no = a.mg_ce_id) AS mg_ce_tel "
objBuilder.Append "FROM as_acpt a "
objBuilder.Append where_sql & base_sql & order_sql & " limit "& stpage & "," &pgsize

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.condi.value == ""){
					alert("소속을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/user_header.asp" -->
			<!--#include virtual = "/include/as_sub_menu_user.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_list_ce_user.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="3%" >
							<col width="3%" >
							<col width="7%" >
							<col width="4%" >
							<col width="5%" >
							<col width="8%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="5%" >
                             <col width="9%" >
							<col width="5%" >
							<col width="*" >
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
								<th scope="col">담당전화번호</th>
								<th scope="col">처리유형</th>
								<th scope="col">장애내용</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim len_date, hangle, bit01, bit02, bit03
						Dim acpt_date, date_to_date, hol_d, dd, com_date
						Dim rs_hol, d_day, as_memo, view_memo

						Dim a, b, c, d

						Do Until rs.EOF
							acpt_date = rs("acpt_date")
							len_date = Len(acpt_date)
							bit01 = Left(acpt_date, 10)
						' 	bit01 = Replace(bit01,"-",".")
							bit03 = Left(Right(acpt_date, 5), 2)
							hangle = Mid(acpt_date, 12, 2)

							If len_date = 22 Then
								bit02 = Mid(acpt_date, 15, 2)
							Else
								bit02 = "0"&Mid(acpt_date, 15, 1)
							End If

							If hangle = "오후" And bit02 <> 12 Then
								bit02 = bit02 + 12
							End If

							date_to_date = bit01 & " " &bit02 & ":" & bit03
							acpt_date = Mid(date_to_date, 3)

						'휴일 계산
							hol_d = 0
							acpt_date = DateValue(Mid(rs("acpt_date"), 1, 10))
							dd = DateDiff("d", acpt_date, curr_date)
							If dd > 0 Then
								a = DateDiff("d", acpt_date, curr_date)
								b = DatePart("w", acpt_date)
								c = a + b
								d = a

								If a > 1 Then
									If c > 7 Then
										d = a - 2
									End If
								End If

						'		visit_date = rs("visit_date")
								com_date = acpt_date
						'		act_date = com_date

								Do Until com_date > curr_date
									'sql_hol = "select * from holiday where holiday = '" + cstr(com_date) + "'"
									objBuilder.Append "SELECT holiday FROM holiday WHERE holiday = '"&CStr(com_date)&"' "

									Set rs_hol = DbConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rs_hol.eof Or rs_hol.bof Then
										d = d
									Else
										d = d -1
									End If

									com_date = DateAdd("d", 1, com_date)
									rs_hol.close()
								Loop
								Set rs_hol = Nothing

								If d > 6 Then
									hol_d = Int(d/7) * 2
								End If
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
							Else
								d_day = 0
							End If
						' 휴일 계산 끝
							as_memo = Replace(rs("as_memo"), Chr(34), Chr(39))
							view_memo = as_memo

							If Len(as_memo) > 15 Then
								view_memo = Mid(as_memo, 1, 15) & ".."
							End If
						%>
							<tr>
								<td class="first"><%=acpt_date%></td>
								<td><%=rs("as_process")%></td>
								<td><%=d_day%></td>
								<td><%=rs("acpt_man")%>&nbsp;<%=rs("acpt_grade")%></td>
								<td>
						<%'아래 조건 확인 시 결과가 True는 나오지 않음(로그인 시 grade가 5가 아니고 Team이 외주관리인 회원이 없음)[허정호_20211221]
						If c_grade <> "5" Then %>
                                <a href="as_result_reg_user.asp?acpt_no=<%=rs("acpt_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>"><%=rs("acpt_user")%>&nbsp;<%=rs("user_grade")%></a>
						<%Else%>
								<%=rs("acpt_user")%>&nbsp;<%=rs("user_grade")%>
                        <%End If %>
                                </td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("tel_ddd")%>)<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%></td>
								<td><%=Mid(CStr(rs("request_date")), 3)%>&nbsp;<%=rs("request_time")%></td>
								<td><%=rs("mg_ce")%></td>
						        <td><%=rs("mg_ce_tel")%></td>
								<td><%=rs("as_type")%></td>
							  	<td class="left"><p style="cursor:pointer"><span title="<%=as_memo%>"><%=view_memo%></span></p></td>
							</tr>
						<%
							rs.MoveNext()
						Loop
						rs.close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page, i

                intstart = (Int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                If intend > total_page Then
                    intend = total_page
                End If
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
					<div class="btnCenter">
                    <a href="/service/excel/excel_down_ce_user.asp" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <%
					'page navi
					Call Page_Navi(page, be_pg, str_param, total_page)
					%>
                    </td>
				    <td width="15%">
					<div class="btnCenter">
					<% If view_sort = "DESC" Then %>
                          <a href="/service/as_list_ce_user.asp?view_sort=asc" class="btnType04">정순조회</a>
                    <% Else %>
                          <a href="/service/as_list_ce_user.asp?view_sort=desc" class="btnType04">역순조회</a>
                    <% End If %>
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