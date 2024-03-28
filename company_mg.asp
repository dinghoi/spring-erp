<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

Page=Request("page")
use_sw = request("use_sw")  
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
	use_sw = request.form("use_sw")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
	use_sw = request("use_sw")  
end if

if use_sw = "" then
	view_condi = "전체"
	use_sw = "T"
	condi_sql = ""
	condi = ""
	use_sql = ""
end if

where_sql = " where (trade_id = '매출' or trade_id = '공통') "

if view_condi = "전체" then
	condi_sql = " "
  else
	if condi = "" then
		condi_sql = "and " + view_condi + " = '" + condi + "'"
	  else
		condi_sql = "and " + view_condi + " like '%" + condi + "%'"
	end if
end if

if use_sw = "T" then
	use_sql = " "
  else
 	use_sql = " and use_sw = '" + use_sw + "'"
end if

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

Sql = "SELECT count(*) FROM trade "&where_sql&condi_sql&use_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM trade "&where_sql&condi_sql&use_sql&" ORDER BY trade_name ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "서비스 관련 회사코드 관리"
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
				return "5 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="company_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>사용구분</strong>
                                <input name="use_sw" type="radio" value="T"  <% if use_sw = "T" then %>checked<% end if %> style="width:25px">총괄
                                <input name="use_sw" type="radio" value="Y"  <% if use_sw = "Y" then %>checked<% end if %> style="width:25px">사용
                                <input name="use_sw" type="radio" value="N"  <% if use_sw = "N" then %>checked<% end if %> style="width:25px">미사용
								</label>
                                <label>
								<strong>조회조건</strong>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="trade_name" <%If view_condi = "trade_name" then %>selected<% end if %>>거래처명</option>
                                  <option value="group_name" <%If view_condi = "group_name" then %>selected<% end if %>>그룹명관리</option>
                                  <option value="support_company" <%If view_condi = "support_company" then %>selected<% end if %>>지원회사</option>
                                </select>
								</label>
                                <label>
								<strong>조건 : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
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
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="10%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">거래처(회사명)</th>
								<th scope="col">거래처유형</th>
								<th scope="col">그룹</th>
								<th scope="col">관리그룹</th>
								<th scope="col">지원회사</th>
								<th scope="col">사용유무</th>
								<th scope="col">변경</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
							if rs("mg_group") = "1" then
								mg_group = "일반그룹" 
							  elseif rs("mg_group") = "2" then
							  	mg_group = "한진그룹"
							  else
							 	mg_group = "기타그룹"
							end if
							
							if rs("use_sw") = "Y" then
								view_use = "사용"
							  else
							  	view_use = "미사용"
							end if
	           			%>
							<tr>
								<td class="first"><%=rs("trade_name")%></td>
								<td><%=rs("trade_id")%></td>
								<td><%=rs("group_name")%>&nbsp;</td>
								<td><%=mg_group%></td>
								<td><%=rs("support_company")%></td>
								<td><%=view_use%></td>
								<td><a href="#" onClick="pop_Window('company_mod.asp?trade_code=<%=rs("trade_code")%>&u_type=<%="U"%>','company_mod_pop','scrollbars=yes,width=750,height=250')">변경</a></td>
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
				    <td width="20%"></td>
				    <td>
                  <div id="paging">
                        <a href = "company_mg.asp?page=<%=first_page%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="company_mg.asp?page=<%=intstart -1%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="company_mg.asp?page=<%=i%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="company_mg.asp?page=<%=intend+1%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="company_mg.asp?page=<%=total_page%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

