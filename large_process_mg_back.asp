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
Set rs_into = Server.CreateObject("ADODB.Recordset")
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
sql = "select large_paper_no,company,as_type,as_process,count(*) as acpt_cnt,sum(dev_inst_cnt) as inst_cnt,sum(ran_cnt) as ran_cnt from "& _
"as_acpt where large_paper_no <> '' group by large_paper_no,company,as_type,as_process order by large_paper_no desc limit "& stpage & "," &pgsize 

rs.Open Sql, Dbconn, 1
first_sw = "y"
i = 1
do until rs.eof
	if firsr_sw = "y" then
		bi_paper_no = rs("large_paper_no")
		bi_company = rs("company")
		bi_type = rs("as_type")
		first_sw = "n"
	end if
	
	if bi_paper_no <> rs("large_paper_no") then
		bi_mg_ce = mg_ce_id
		mg_ce = rs("mg_ce")
		i = i + 1	
	end if
	if rs("as_process") = "완료" or rs("as_process") = "취소" then
		cnt_tab(i,1) = cnt_tab(i,1) + 1
		cnt_tab(i,3) = cnt_tab(i,3) + int(rs("dev_inst_cnt"))
		cnt_tab(i,5) = cnt_tab(i,5) + int(rs("ran_cnt"))
	  else
		cnt_tab(i,2) = cnt_tab(i,2) + 1
		cnt_tab(i,4) = cnt_tab(i,4) + int(rs("dev_inst_cnt"))
		cnt_tab(i,6) = cnt_tab(i,6) + int(rs("ran_cnt"))
	end if
	rs.movenext()
loop

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
                                <select name="company" id="company">
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
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="*" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">문서번호</th>
								<th scope="col">회사</th>
								<th scope="col">처리유형</th>
								<th scope="col">진행현황</th>
								<th scope="col">설치수량</th>
								<th scope="col">공사수량</th>
							</tr>
						</thead>
						<tbody>
						<%
 						do until rs.eof
						%>
							<tr>
								<td class="first"><%=rs("large_paper_no")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("as_type")%></td>
								<td><%=rs("as_process")%></td>
								<td><%=rs("inst_cnt")%></td>
								<td><%=rs("ran_cnt")%></td>
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

