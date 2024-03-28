<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
On Error Resume Next

Dim Rs
Dim Repeat_Rows
dim view_tab(4,50)
dim page_cnt
dim pg_cnt
Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

ck_sw=Request("ck_sw")
Page=Request("page")

be_pg = "insa_org_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

reg_user = request.cookies("nkpmg_user")("coo_user_name")
mod_user = request.cookies("nkpmg_user")("coo_user_name")

view_condi = Request.form("view_condi")
sel_company = Request.form("sel_company")
sel_bonbu = Request.form("sel_bonbu")
sel_saupbu = Request.form("sel_saupbu")
sel_team = Request.form("sel_team")

if ck_sw = "y" then
    view_condi = request("view_condi")
	sel_company = Request("sel_company")
    sel_bonbu = Request("sel_bonbu")
    sel_saupbu = Request("sel_saupbu") 
	sel_team = Request("sel_team") 
  else
	view_condi = request.form("view_condi")
	sel_company = Request.form("sel_company")
    sel_bonbu = Request.form("sel_bonbu")
    sel_saupbu = Request.form("sel_saupbu")
	sel_team = Request.form("sel_team")
end if

if view_condi = "" then
	view_condi = "1"
	condi_sql = " "
	sel_company = "케이원정보통신"
    sel_bonbu = ""
	sel_saupbu = ""
	sel_team = ""
'	for i = 0 to 4
'	    for j = 0 to 50
'		    view_tab(i,j) = ""
'	    next
'   next
end if

for i = 0 to 4
    for j = 0 to 50
    view_tab(i,j) = ""
    next
next

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_tab = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "1" then

   '회사
	k = 0
    Sql="select org_name from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_level = '회사') ORDER BY org_code ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(1,k) = rs_tab("org_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	
	
	'본부
	k = 0
    Sql="select org_name from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_company = '"+sel_company+"') and  (org_level = '본부') ORDER BY org_code ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(2,k) = rs_tab("org_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	
	
	'사업부
	k = 0
    Sql="select org_name from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_company = '"+sel_company+"') and  (org_level = '사업부') ORDER BY org_code ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(3,k) = rs_tab("org_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	
	
	'팀
	k = 0
   Sql="select org_name from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_company = '"+sel_company+"') and  (org_level = '팀') ORDER BY org_code ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(4,k) = rs_tab("org_name")
		rs_tab.movenext()
	Wend
   rs_tab.close()	
	
  else	


'회사
	k = 0
    Sql="select org_name from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_level = '회사') ORDER BY org_code ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(1,k) = rs_tab("org_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	

'본부
	k = 0
    Sql="select org_name from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_company = '"+sel_company+"') and  (org_level = '본부') ORDER BY org_code ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(2,k) = rs_tab("org_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	

'사업부
	k = 0
    Sql="select org_name from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_company = '"+sel_company+"') and (org_bonbu = '"+sel_bonbu+"') and  (org_level = '사업부') ORDER BY org_code ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(3,k) = rs_tab("org_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	

'팀
	k = 0
   Sql="select org_name from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_company = '"+sel_company+"') and (org_bonbu = '"+sel_bonbu+"') and (org_saupbu = '"+sel_saupbu+"') and  (org_level = '팀') ORDER BY org_code ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(4,k) = rs_tab("org_name")
		rs_tab.movenext()
	Wend
   rs_tab.close()	
end if


if view_condi = "1" then
   condi_Sql = " and (org_company = '" + sel_company + "')"
end if

if view_condi = "2" then
   condi_Sql = " and (org_company = '"+sel_company+"') and (org_bonbu = '" + sel_bonbu + "')"
end if

if view_condi = "3" then
   condi_Sql = " and (org_company = '"+sel_company+"') and (org_bonbu = '" + sel_bonbu + "') and (org_saupbu = '" + sel_saupbu + "')"
end if

if view_condi = "4" then
   condi_Sql = " and (org_company = '"+sel_company+"') and (org_bonbu = '" + sel_bonbu + "') and (org_saupbu = '" + sel_saupbu + "') and (org_team = '" + sel_team + "')"
end if

view_sort = request("view_sort")
if view_sort = "" then
	view_sort = "ASC"
end if

order_Sql = " ORDER BY org_company,org_bonbu,org_saupbu,org_team,org_code " + view_sort
where_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00')"

Sql = "SELECT count(*) FROM emp_org_mst " + where_sql + condi_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_org_mst " + where_sql + condi_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " 조직별 현황 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			
			function chkfrm() {
				k = 0;
				for (j=0;j<4;j++) {
					if (eval("document.frm.view_condi[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("조건검색 기준을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
			function condi_view() {
				k = 0;
				for (j=0;j<4;j++) {
					if (eval("document.frm.view_condi[" + j + "].checked")) {
						k = j + 1
					}
				}
				if (k==1){
					document.frm.sel_company.style.display = '';				
					document.frm.sel_bonbu.style.display = 'none';				
					document.frm.sel_saupbu.style.display = 'none';
					document.frm.sel_team.style.display = 'none';
				}
				if (k==2){
					document.frm.sel_company.style.display = 'none';				
					document.frm.sel_bonbu.style.display = '';				
					document.frm.sel_saupbu.style.display = 'none';	
					document.frm.sel_team.style.display = 'none';
				}
				if (k==3){
					document.frm.sel_company.style.display = 'none';				
					document.frm.sel_bonbu.style.display = 'none';				
					document.frm.sel_saupbu.style.display = '';	
					document.frm.sel_team.style.display = 'none';
				}
				if (k==4){
					document.frm.sel_company.style.display = 'none';				
					document.frm.sel_bonbu.style.display = 'none';				
					document.frm.sel_saupbu.style.display = 'none';	
					document.frm.sel_team.style.display = '';
				}
			}			
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_org_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                              <input type="radio" name="view_condi" value="1" <% if view_condi = "1" then %>checked<% end if %> title="회사별" style="width:30px" onClick="condi_view()">회사별
                                  <select name="sel_company" id="sel_company" type="text" style="display:none; width:150px">
                                    <%	for i = 1 to 50 
									        if view_tab(1,i) <> "" then %>
                                    <option value="<%=view_tab(1,i)%>" <%If sel_company = view_tab(1,i) then %>selected<% end if %>><%=view_tab(1,i)%></option>
                                    <%	     end if
									    next	%>
                                  </select>
                              <input type="radio" name="view_condi" value="2" <% if view_condi = "2" then %>checked<% end if %> title="본부별" style="width:30px" onClick="condi_view()">본부별
                                  <select name="sel_bonbu" id="sel_bonbu" style="display:none; width:150px">
                                    <%	for i = 1 to 50 
									        if view_tab(2,i) <> "" then %>
                                    <option value="<%=view_tab(2,i)%>" <%If sel_bonbu = view_tab(2,i) then %>selected<% end if %>><%=view_tab(2,i)%></option>
                                    <%	     end if
									    next %>
                                  </select>
                              <input type="radio" name="view_condi" value="3" <% if view_condi = "3" then %>checked<% end if %> title="사업부별" style="width:30px" onClick="condi_view()">사업부별
                                  <select name="sel_saupbu" id="sel_saupbu" style="display:none; width:150px">
                                    <%	for i = 1 to 50 
									        if view_tab(3,i) <> "" then %>
                                    <option value="<%=view_tab(3,i)%>" <%If sel_saupbu = view_tab(3,i) then %>selected<% end if %>><%=view_tab(3,i)%></option>
                                    <%	     end if
									    next	%>
                                  </select>
                              <input type="radio" name="view_condi" value="4" <% if view_condi = "4" then %>checked<% end if %> title="팀별" style="width:30px" onClick="condi_view()">팀별
                                  <select name="sel_team" id="sel_team" style="display:none; width:150px">
                                    <%	for i = 1 to 50 
									        if view_tab(4,i) <> "" then %>
                                    <option value="<%=view_tab(4,i)%>" <%If sel_team = view_tab(4,i) then %>selected<% end if %>><%=view_tab(4,i)%></option>
                                    <%	     end if
									    next	%>
                                  </select>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>				
                <form name="frm_del" method="post" action="org_del_ok.asp?page=<%=page%>&ck_sw=<%="n"%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
					 <colgroup>
				         <col width="8%" >
				         <col width="8%" >
				         <col width="8%" >
				         <col width="8%" >
                         <col width="4%" >
				         <col width="8%" >
                         <col width="6%" >
                         <col width="4%" >
				         <col width="6%" >
				         <col width="8%" >
                         <col width="8%" >
				         <col width="6%" >
                         <col width="6%" >
				         <col width="6%" >
                         <col width="3%" >
					 </colgroup>
				 		<thead>
				      <tr>
				        <th colspan="4" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
                        <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;직&nbsp;&nbsp;장</th>
				        <th rowspan="2" scope="col">상주회사</th>
                        <th rowspan="2" scope="col">조직생성일</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">상위&nbsp;조직장</th>
                        <th rowspan="2" scope="col">비고</th>
			          </tr>
                      <tr>
				        <th class="first" scope="col">회&nbsp;&nbsp;사</th>
				        <th scope="col">본&nbsp;&nbsp;부</th>
				        <th scope="col">사업부</th>
				        <th scope="col">팀</th>
                        <th scope="col">코드</th>
				        <th scope="col">조직명</th>
                        <th scope="col">조직<br>Level</th>
                        <th scope="col">T.O</th>
				        <th scope="col">사번</th>
				        <th scope="col">성명</th>
				        <th scope="col">사번</th>
                        <th scope="col">성명</th>
                      </tr>
						</thead>
						<tbody>
                      <%
						do until rs.eof
					  %>
				      <tr>
				        <td class="first"><%=rs("org_company")%>&nbsp;</td>
				        <td><%=rs("org_bonbu")%>&nbsp;</td>
                        <td><%=rs("org_saupbu")%>&nbsp;</td>
                        <td><%=rs("org_team")%>&nbsp;</td>
                        <td><%=rs("org_code")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_org_view.asp?org_code=<%=rs("org_code")%>&org_name=<%=org_name%>&u_type=<%="U"%>','insa_org_view_pop','scrollbars=yes,width=750,height=350')"><%=rs("org_name")%></a>&nbsp;</td>
                        <td><%=rs("org_level")%>&nbsp;</td>
                        <td><%=rs("org_table_org")%>&nbsp;</td>
                        <td><%=rs("org_empno")%>&nbsp;</td>
                        <td><%=rs("org_emp_name")%>&nbsp;</td>
                        <td><%=rs("org_reside_company")%>&nbsp;</td>
                        <td><%=rs("org_date")%>&nbsp;</td>
                        <td><%=rs("org_owner_empno")%>&nbsp;</td>
                        <td><%=rs("org_owner_empname")%>&nbsp;</td>
                        <td>&nbsp;</td>
                      <% 
                        '<td><a href="#" onClick="pop_Window('insa_org_modify.asp?org_code=<%=rs("org_code")%> <% '&u_type=<%="U"%>  <% '','insa_org_reg_pop','scrollbars=yes,width=1400,height=600')">수정</a>&nbsp;</td> %>
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
                    <a href="insa_excel_orglist.asp?view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_org_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_org_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_org_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_org_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_org_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
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

