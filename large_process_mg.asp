<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim field_check
Dim field_view
Dim win_sw
dim paper_tab(10)
dim com_tab(10)
dim type_tab(10)
dim cnt_tab(10,6)

for i = 1 to 10
	paper_tab(i) = ""
	com_tab(i) = ""
	type_tab(i) = ""
	for j = 1 to 6
		cnt_tab(i,j) = 0
	next
next

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	field_check=Request("field_check")
	field_view=Request("field_view")
	company=Request("company")
  else
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	company=Request.form("company")
End if

If company = "" Then
	field_check = "total"
	company = "전체"
End If

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' 화면 한 페이지 

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
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 조건별 조회.........
base_sql = "select * from large_acpt "

where_sql = "where upload_ok = 'N' "

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if

if company = "전체" then
	com_sql = " "
  else
  	com_sql = " and (company = '" + company + "') "
end if

order_sql = " ORDER BY paper_no, sido, gugun, dong, addr ASC"

Sql = "SELECT large_paper_no FROM as_acpt where large_paper_no <> ''  group by large_paper_no"
'Set RsCount = Dbconn.Execute (sql)
rs.Open Sql, Dbconn, 1
tottal_record = 0
do until rs.eof
	tottal_record = tottal_record + 1
	rs.movenext()
loop
rs.close()
'tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + where_sql + com_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
sql = "select large_paper_no,company,as_type,start_date,end_date,count(*) as acpt_cnt from "& _
"as_acpt where large_paper_no <> '' group by large_paper_no,company,as_type order by end_date asc limit "& stpage & "," &pgsize 
Response.write "<!-- " & sql & " -->"
rs.Open Sql, Dbconn, 1

title_line = "대량건 진행 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				return "1 1";
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
			<!--#include virtual = "/include/large_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="large_process_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>회사</strong>
								<%
                                Sql="select * from trade where use_sw = 'Y' and mg_group = '"+mg_group+"' order by trade_name asc"
                                rs_etc.Open Sql, Dbconn, 1
                                %>
                                <select name="company" id="company" >
 									<option value="전체">전체</option> 
          					<% 
								While not rs_etc.eof 
							%>
          							<option value='<%=rs_etc("trade_name")%>' <%If rs_etc("trade_name") = company  then %>selected<% end if %>><%=rs_etc("trade_name")%></option>
          					<%
									rs_etc.movenext()  
								Wend 
								rs_etc.Close()
							%>
                                </select>
								</label>
								<strong>필드조건</strong>
                                <select name="field_check" id="field_check" style="width:70px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                    <option value="paper_no" <% if field_check = "paper_no" then %>selected<% end if %>>문서번호</option>
                                    <option value="mg_ce" <% if field_check = "mg_ce" then %>selected<% end if %>>담당CE</option>
                                    <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>시도</option>
                                    <option value="gugun" <% if field_check = "gugun" then %>selected<% end if %>>구군</option>
                                    <option value="dong" <% if field_check = "dong" then %>selected<% end if %>>동</option>
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
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
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
								<th rowspan="2" class="first" scope="col">회사</th>
								<th rowspan="2" scope="col">문서번호</th>
								<th rowspan="2" scope="col">처리유형</th>
								<th rowspan="2" scope="col">시작일</th>
								<th rowspan="2" scope="col">마감일</th>
								<th rowspan="2" scope="col">총건수</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">접 수 건 수</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">설 치 수 량</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">랜 공 사 수 량</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">완료</th>
							  <th scope="col">미처리</th>
							  <th scope="col">진척율(%)</th>
							  <th scope="col">완료수량</th>
							  <th scope="col">미처리수량</th>
							  <th scope="col">완료수량</th>
							  <th scope="col">미처리수량</th>
                          </tr>
						</thead>
						<tbody>
						<%
 						do until rs.eof
							sql = "select count(*) as acpt_cnt, sum(dev_inst_cnt) as inst_cnt, sum(ran_cnt) as ran_cnt from as_acpt "& _
							" where large_paper_no ='"&rs("large_paper_no")&"' and (as_process = '완료' or as_process = '취소') "& _
							" group by large_paper_no"
							set rs_sum=dbconn.execute(sql)
							if rs_sum.eof or rs_sum.bof then
								end_acpt_cnt = 0
								end_inst_cnt = 0
								end_ran_cnt = 0
							  else
								end_acpt_cnt = cint(rs_sum("acpt_cnt"))
								end_inst_cnt = cint(rs_sum("inst_cnt"))
								end_ran_cnt = cint(rs_sum("ran_cnt"))
							end if
							rs_sum.close()

							sql = "select count(*) as acpt_cnt, sum(dev_inst_cnt) as inst_cnt, sum(ran_cnt) as ran_cnt from as_acpt "& _
							" where large_paper_no ='"&rs("large_paper_no")&"' and (as_process = '접수' or as_process = '입고' or "& _
							" as_process = '연기') group by large_paper_no"
							set rs_sum=dbconn.execute(sql)
							if rs_sum.eof or rs_sum.bof then
								mi_acpt_cnt = 0
								mi_inst_cnt = 0
								mi_ran_cnt = 0
							  else
								mi_acpt_cnt = cint(rs_sum("acpt_cnt"))
								mi_inst_cnt = cint(rs_sum("inst_cnt"))
								mi_ran_cnt = cint(rs_sum("ran_cnt"))
							end if
							rs_sum.close()

							if end_acpt_cnt = 0 then
								ing_pro = 0
							  else
								ing_pro = end_acpt_cnt/cint(rs("acpt_cnt")) * 100
							end if
						%>
							<tr>
								<td class="first"><%=rs("company")%></td>
								<td><a href="#" onClick="pop_Window('area_large_process.asp?large_paper_no=<%=rs("large_paper_no")%>&company=<%=rs("company")%>&as_type=<%=rs("as_type")%>&acpt_cnt=<%=rs("acpt_cnt")%>','area_large_process_popup','scrollbars=yes,width=750,height=730')"><%=rs("large_paper_no")%></a></td>
								<td><%=rs("as_type")%></td>
								<td><%=rs("start_date")%></td>
								<td><%=rs("end_date")%></td>
								<td><%=formatnumber(rs("acpt_cnt"),0)%></td>
								<td><%=formatnumber(end_acpt_cnt,0)%></td>
								<td><%=formatnumber(mi_acpt_cnt,0)%></td>
								<td><%=formatnumber(ing_pro,2)%>%</td>
								<td><%=formatnumber(end_inst_cnt,0)%></td>
								<td><%=formatnumber(mi_inst_cnt,0)%></td>
								<td><%=formatnumber(end_ran_cnt,0)%></td>
								<td><%=formatnumber(mi_ran_cnt,0)%></td>
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
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="large_process_mg.asp?page=<%=first_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="large_process_mg.asp?page=<%=intstart -1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="large_process_mg.asp?page=<%=i%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="large_process_mg.asp?page=<%=intend+1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[다음]</a> <a href="large_process_mg.asp?page=<%=total_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
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

