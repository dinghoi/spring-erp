<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw
dim company_tab(150)

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	company=Request("company")
	as_type=Request("as_type")
	field_check=Request("field_check")
	field_view=Request("field_view")

Else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	company=Request.form("company")
	as_type=Request.form("as_type")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If to_date = "" or from_date = "" Then
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-7),1,10)
	field_check = "total"
	company = "전체"
	as_type = "전체"
End If

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

' 조건별 조회.........
' 날짜별 조회(1)
base_sql = "select *  from att_file "
date_sql = "where (visit_date >= '" + from_date  + "' and visit_date <= '" + to_date  + "')"
if company = "전체" then
	company_sql = ""
  else
	company_sql = " and ( company = '" + company + "') "
end if
if as_type = "전체" then
	type_sql = ""
  else
	type_sql = " and ( as_type = '" + as_type + "') "
end if

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if
order_sql = " ORDER BY visit_date DESC"

Sql = "SELECT count(*) FROM att_file " + date_sql + company_sql + type_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + date_sql + company_sql + type_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1


title_line = "설치/공사 첨부관리"
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
				return "0 1";
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
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.field_check.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/as_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="att_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>회사</strong>
								<%
                                if c_grade = "7" or (c_grade = "5" and c_reside = "1") then
                                    sql_trade="select * from trade where use_sw = 'Y' and mg_group = '"+mg_group+"' and trade_name = '"+user_name+"' order by etc_name asc"
                                end if
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
								<label>
								<strong>등록일&nbsp;&nbsp;시작 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<label>
								<strong>처리유형</strong>
                                <select name="as_type" id="as_type" style="width:100px">
                                  <option value="전체" <%If as_type = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="신규설치" <%If as_type = "신규설치" then %>selected<% end if %>>신규설치</option>
                                  <option value="신규설치공사" <%If as_type = "신규설치공사" then %>selected<% end if %>>신규설치공사</option>
                                  <option value="이전설치" <%If as_type = "이전설치" then %>selected<% end if %>>이전설치</option>
                                  <option value="이전설치공사" <%If as_type = "이전설치공사" then %>selected<% end if %>>이전설치공사</option>
                                  <option value="랜공사" <%If as_type = "랜공사" then %>selected<% end if %>>랜공사</option>
                                  <option value="이전랜공사" <%If as_type = "이전랜공사" then %>selected<% end if %>>이전랜공사</option>
                                  <option value="장비회수" <%If as_type = "장비회수" then %>selected<% end if %>>장비회수</option>
                                  <option value="예방점검" <%If as_type = "예방점검" then %>selected<% end if %>>예방점검</option>
                                </select>
								</label>
                                <label>
								<strong>조건검색</strong>
                                <select name="field_check" id="field_check" style="width:80px">
                                    <option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                    <option value="acpt_no" <% if field_check = "acpt_no" then %>selected<% end if %>>접수번호</option>
                                    <option value="mg_ce" <% if field_check = "mg_ce" then %>selected<% end if %>>담당CE</option>
                                    <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>시도</option>
                                    <option value="gugun" <% if field_check = "gugun" then %>selected<% end if %>>구군</option>
                                    <option value="dept" <% if field_check = "dept" then %>selected<% end if %>>조직명</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:80px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
							<col width="7%" >
							<col width="12%" >
							<col width="18%" >
							<col width="13%" >
							<col width="6%" >
							<col width="6%" >
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">처리유형</th>
								<th scope="col">처리일자</th>
								<th scope="col">회사</th>
								<th scope="col">부서</th>
								<th scope="col">지역</th>
								<th scope="col">담당CE</th>
								<th scope="col">접수번호</th>
								<th scope="col">첨부파일</th>
								<th scope="col">첨부변경</th>
								<th scope="col">세부내역</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							path = "/att_file/" + rs("company")
						%>
							<tr>
								<td class="first"><%=rs("as_type")%></td>
								<td><%=rs("visit_date")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%></td>
								<td><%=rs("mg_ce")%></td>
								<td><%=rs("acpt_no")%></td>
								<td>&nbsp;
								<%
                                    if rs("att_file1") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file1")%>">첨부1</a>&nbsp;
                                <%
                                    end if
                                    if rs("att_file2") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file2")%>">첨부2</a>&nbsp;
                                <%
                                    end if
                                    if rs("att_file3") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file3")%>">첨부3</a>&nbsp;
                                <%
                                    end if
                                    if rs("att_file4") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file4")%>">첨부4</a>&nbsp;
                                <%
                                    end if
                                    if rs("att_file5") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file5")%>">첨부5</a>&nbsp;
                                <%
                                    end if
                                %>
                                </td>
								<td><a href="#" onClick="pop_Window('att_file_mod.asp?acpt_no=<%=rs("acpt_no")%>','att_file_mod_pop','scrollbars=yes,width=800,height=410')">변경</a></td>
								<td><a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=rs("acpt_no")%>','asview_pop','scrollbars=yes,width=800,height=700')">조회</a></td>
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
				    <td width="15%"></td>
				    <td>
                    <div id="paging">
                        <a href = "att_list.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="att_list.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="att_list.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="att_list.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[다음]</a> <a href="att_list.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%"></td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

