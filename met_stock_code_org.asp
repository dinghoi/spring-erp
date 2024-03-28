<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
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

be_pg = "met_stock_code_org.asp"
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
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'회사
	k = 0
    Sql="select * from met_stock_code where (isNull(stock_end_date) or stock_end_date = '1900-01-01') and (stock_level = '본사') ORDER BY stock_name ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(1,k) = rs_tab("stock_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	

'본부
	k = 0
    Sql="select * from met_stock_code where (isNull(stock_end_date) or stock_end_date = '1900-01-01') and (stock_company = '"+sel_company+"') and  (stock_level = '본부') ORDER BY stock_name ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(2,k) = rs_tab("stock_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	

'사업부
	k = 0
    Sql="select * from met_stock_code where (isNull(stock_end_date) or stock_end_date = '1900-01-01') and (stock_company = '"+sel_company+"') and (stock_bonbu = '"+sel_bonbu+"') and  (stock_level = '사업부') ORDER BY stock_name ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(3,k) = rs_tab("stock_name")
		rs_tab.movenext()
	Wend
    rs_tab.close()	

'팀
	k = 0
   Sql="select * from met_stock_code where (isNull(stock_end_date) or stock_end_date = '1900-01-01') and (stock_company = '"+sel_company+"') and (stock_bonbu = '"+sel_bonbu+"') and (stock_saupbu = '"+sel_saupbu+"') and  (stock_level = '팀') ORDER BY stock_name ASC"
	rs_tab.Open Sql, Dbconn, 1	
	while not rs_tab.eof
		k = k + 1
		view_tab(4,k) = rs_tab("stock_name")
		rs_tab.movenext()
	Wend
   rs_tab.close()	

if view_condi = "1" then
   condi_Sql = " and (stock_company = '" + sel_company + "')"
end if

if view_condi = "2" then
   condi_Sql = " and (stock_company = '"+sel_company+"') and (stock_bonbu = '" + sel_bonbu + "')"
end if

if view_condi = "3" then
   condi_Sql = " and (stock_company = '"+sel_company+"') and (stock_bonbu = '" + sel_bonbu + "') and (stock_saupbu = '" + sel_saupbu + "')"
end if

if view_condi = "4" then
   condi_Sql = " and (stock_company = '"+sel_company+"') and (stock_bonbu = '" + sel_bonbu + "') and (stock_saupbu = '" + sel_saupbu + "') and (stock_team = '" + sel_team + "')"
end if

view_sort = request("view_sort")
if view_sort = "" then
	view_sort = "DESC"
end if

order_Sql = " ORDER BY stock_level,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_name DESC" 
'order_Sql = " ORDER BY stock_code " + view_sort
where_sql = " WHERE (isNull(stock_end_date) or stock_end_date = '1900-01-01')"

Sql = "SELECT count(*) FROM met_stock_code " + where_sql + condi_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_stock_code " + where_sql + condi_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1


title_line = " 조직별 창고 현황 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>상품자재관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "6 1";
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/meterials_control_header01.asp" -->
            <!--#include virtual = "/include/meterials_basic_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_code_org.asp?ck_sw=<%="n"%>" method="post" name="frm">
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

				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
					 <colgroup>
				         <col width="6%" >
				         <col width="10%" >
                         <col width="6%" >
				         <col width="10%" >
				         <col width="10%" >
                         <col width="6%" >
				         <col width="6%" >
                         <col width="6%" >
                         <col width="6%" >
				         <col width="*" >
                         <col width="3%" >
					 </colgroup>
				 	<thead>
				      <tr>
				        <th class="first" scope="col">창고코드</th>
				        <th scope="col">창고명</th>
                        <th scope="col">창고유형</th>
                        <th scope="col">창고장</th>
                        <th scope="col">회사</th>
                        <th scope="col">생성일</th>
                        <th scope="col">폐쇄일</th>
                        <th scope="col">출고담당</th>
                        <th scope="col">입고담당</th>
                        <th scope="col">소속조직</th>
                        <th scope="col">비고</th>
			          </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof 
						   stock_end_date = rs("stock_end_date")
						   if stock_end_date = "1900-01-01" then
	                            stock_end_date = ""
	                       end if
					  %>
				      <tr>
				        <td class="first"><%=rs("stock_code")%>&nbsp;</td>
                        <td><%=rs("stock_name")%>&nbsp;</td>
                        <td><%=rs("stock_level")%>&nbsp;</td>
                        <td><%=rs("stock_manager_name")%>(<%=rs("stock_manager_code")%>)&nbsp;</td>
                        <td><%=rs("stock_company")%>&nbsp;</td>
                        <td><%=rs("stock_open_date")%>&nbsp;</td>
                        <td><%=stock_end_date%>&nbsp;</td>
                        <td><%=rs("stock_go_name")%>&nbsp;</td>
                        <td><%=rs("stock_in_name")%>&nbsp;</td>
                        <td class="left"><%=rs("stock_bonbu")%>-<%=rs("stock_saupbu")%>-<%=rs("stock_team")%>&nbsp;</td>
                    <% if stock_level <> "개인" then %>
                        <td><a href="#" onClick="pop_Window('met_stock_code_add.asp?stock_code=<%=rs("stock_code")%>&stock_name=<%=rs("stock_name")%>&stock_level=<%=rs("stock_level")%>&u_type=<%="U"%>','met_stock_code_pop','scrollbars=yes,width=750,height=300')">수정</a>&nbsp;</td>
                    <% else %>
                        <td>&nbsp;</td>
                    <% end if %>
			          </tr>
				      <%
							rs.movenext()
						loop
						rs.close()
						%>
			        </tbody
				  ></table>
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
                    <a href="met_stock_code_org_excel.asp?view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "met_stock_code_org.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_stock_code_org.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_stock_code_org.asp?page=<%=i%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="met_stock_code_org.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_stock_code_org.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&sel_company=<%=sel_company%>&sel_bonbu=<%=sel_bonbu%>&sel_saupbu=<%=sel_saupbu%>&sel_team=<%=sel_team%>&ck_sw=<%="y"%>">[마지막]</a>
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

