<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_condi = "전체"
	condi_sql = " "
	condi = ""
end if

if view_condi = "사용자" then
	condi_sql = " and user_name like '%" + condi + "%'"
end if
if view_condi = "직급별" then
	condi_sql = " and user_grade like '%" + condi + "%'"
end if
if view_condi = "직위별" then
	condi_sql = " and position like '%" + condi + "%'"
end if
if view_condi = "팀별" then
	condi_sql = "and team like '%" + condi + "%'"
end if
if view_condi = "상주처별" then
	condi_sql = "and reside_place like '%" + condi + "%'"
end if

use_sql = " and grade < '5'"
emp_sql = "(emp_no < '200000') "

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "SELECT count(*) FROM memb where "+emp_sql+condi_sql+use_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM memb where "+emp_sql+condi_sql+use_sql+" ORDER BY user_name ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "사용자별 비용 권한 관리"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
		</script>
	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="cost_grade_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<strong>항목조건 : </strong>
                                <select name="view_condi" id="select3" style="width:150px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="사용자" <%If view_condi = "사용자" then %>selected<% end if %>>사용자</option>
                                  <option value="직급별" <%If view_condi = "직급별" then %>selected<% end if %>>직급별</option>
                                  <option value="직위별" <%If view_condi = "직위별" then %>selected<% end if %>>직위별</option>
                                  <option value="팀별" <%If view_condi = "팀별" then %>selected<% end if %>>팀별</option>
                                  <option value="상주처별" <%If view_condi = "상주처별" then %>selected<% end if %>>상주처별</option>
                                </select>
								<strong>조건 : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="8%" >
							<col width="*" >
							<col width="10%" >
							<col width="8%" >
							<col width="8%" >
							<col width="12%" >
							<col width="8%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">이름</th>
								<th scope="col">아이디</th>
								<th scope="col">소속</th>
								<th scope="col">핸드폰</th>
								<th scope="col">서비스권한</th>
								<th scope="col">관리그룹</th>
								<th scope="col">상주처</th>
								<th scope="col">비용권한</th>
								<th scope="col">변경</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							if rs("grade") = 0 then
								grade_view = "마스터"
							end if
							if rs("grade") = 1 then
								grade_view = "관리자"
							end if
							if rs("grade") = 2 then
								grade_view = "상주관리자"
							end if
							if rs("grade") = 3 then
								grade_view = "상주CE"
							end if
							if rs("grade") = 4 then
								grade_view = "CE"
							end if
							if rs("grade") = 5 then
								grade_view = "사용자"
							end if
							if rs("grade") > 5 or rs("grade") < 0 then
								grade_view = "권한없음"
							end if

							if rs("cost_grade") = 0 then
								cost_grade_view = "마스터"
							end if
							if rs("cost_grade") = 1 then
								cost_grade_view = "본부장권한"
							end if
							if rs("cost_grade") = 2 then
								cost_grade_view = "사업부장권한"
							end if
							if rs("cost_grade") = 3 then
								cost_grade_view = "비용대행"
							end if
							if rs("cost_grade") = 4 then
								cost_grade_view = "영업및관리"
							end if
							if rs("cost_grade") = 5 then
								cost_grade_view = "일반CE/관리"
							end if
							if rs("cost_grade") = 6 then
								cost_grade_view = "일반CE"
							end if
							if rs("cost_grade") = 7 then
								cost_grade_view = "권한없음"
							end if

							if rs("mg_group") = "2" then
								mg_group = "한진그룹"
							  elseif rs("mg_group") = "1" then
							  	mg_group = "일반그룹"
							  else
							  	mg_group = "Error"
							end if
							i = i + 1
	           			%>
							<tr>
								<td class="first"><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%></td>
								<td><a href="#" onClick="pop_Window('pass_init.asp?user_id=<%=rs("user_id")%>','pass_init_pop','scrollbars=no,width=400,height=200')"><%=rs("user_id")%></a></td>
								<td class="left"><%=rs("bonbu")%>&nbsp;<%=rs("saupbu")%>&nbsp;<%=rs("team")%></td>
								<td><%=rs("hp")%></td>
								<td><%=grade_view%></td>
								<td><%=mg_group%></td>
								<td><%=rs("reside_place")%>&nbsp;</td>
								<td><%=cost_grade_view%></td>
								<td><a href="#" onClick="pop_Window('cost_grade_mod.asp?user_id=<%=rs("user_id")%>&u_type=<%="U"%>','cost_grade_pop','scrollbars=no,width=800,height=170')">변경</a></td>
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
                        <a href = "cost_grade_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="cost_grade_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
       	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                    <a href="cost_grade_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
       	<% if 	intend < total_page then %>
                        <a href="cost_grade_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[다음]</a> <a href="cost_grade_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%" align="center"></td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

