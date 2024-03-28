<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim field_check
Dim field_view
Dim win_sw
dim org_code_last
dim company_tab(50,2)

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	company=Request("company")
	org_gubun=Request("org_gubun")
 else
	company=Request.form("company")
	org_gubun=Request.form("org_gubun")
End if

If company = "" Then
	company = "00"
	org_gubun = "1"
End If

if asset_company <> "00" then
	company = asset_company
end if

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")					
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if company = "00" then
	com_sql = ""
  else
	com_sql = " and asset.company = '" + company + "' "
end if

Sql = "select count(*) from org_code where org_company = '" + company + "' and org_gubun = '" + org_gubun + "'"
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

Sql = "select * from org_code where org_company = '" + company + "' and org_gubun = '" + org_gubun + "' order by org_code asc limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1

title_line = "자산조직 구분코드 관리"
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
			<!--#include virtual = "/include/asset_header.asp" -->
			<!--#include virtual = "/include/asset_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="org_code_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>회사</strong>
								<%
                                    if asset_company = "00" then
                                        k = 0
                                        Sql="select * from etc_code where etc_type = '75' and used_sw = 'Y' order by etc_name asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                        do until rs_etc.eof
                                            k = k + 1
                                            company_tab(k,1) = rs_etc("etc_name")
                                            company_tab(k,2) = mid(rs_etc("etc_code"),3,2)
                                            rs_etc.movenext()
                                        loop
                                        rs_etc.close()						
                                    %>
                                <select name="company" id="company">
                                  <% 
                                            for kk = 1 to k
                                        %>
                                  <option value='<%=company_tab(kk,2)%>' <%If company_tab(kk,2) = company then %>selected<% end if %>><%=company_tab(kk,1)%></option>
                                  <%
                                            next
                                        %>
                                </select>
                                <%		else %>
                                &nbsp;<%=user_name%>
                                <input name="company" type="hidden" id="company" value="<%=company%>">
                                <%	end if %>
								</label>
                                <label>
								<strong>구분코드</strong>
                                <select name="org_gubun" id="org_gubun">
                                    <option value="1" <%If org_gubun = "1" then %>selected<% end if %>>관리조직</option>
                                <% if company = "01" then %>
                                    <option value="2" <%If org_gubun = "2" then %>selected<% end if %>>법인명</option>
                                <% else %>
                                    <option value="2" <%If org_gubun = "2" then %>selected<% end if %>>상위조직</option>
                                <%  end if %>
                                </select>                                </select>
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
							<col width="*" >
							<col width="10%" >
							<col width="20%" >
							<col width="20%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">코드</th>
								<th scope="col">코드명</th>
								<th scope="col">사용유무</th>
								<th scope="col">등록인</th>
								<th scope="col">등록일자</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							if rs("used_sw") = "Y" then
								used_sw = "사용"
							  else
							  	used_sw = "미사용"
							end if
							sql="select * from memb where user_id = '" + rs("reg_id") + "'"
							set rs_memb=dbconn.execute(sql)
						
							if	rs_memb.eof or rs_memb.bof then
								reg_name = "ERROR"
							  else
								reg_name = rs_memb("user_name")
							end if
						%>
							<tr>
								<td class="first"><%=rs("org_code")%></td>
								<td><a href="#" onClick="pop_Window('org_code_add.asp?company=<%=rs("org_company")%>&org_gubun=<%=rs("org_gubun")%>&org_code=<%=rs("org_code")%>&u_type=<%="U"%>','org_code_reg_popup','scrollbars=yes,width=500,height=280')"><%=rs("org_name")%></a></td>
								<td><%=used_sw%></td>
								<td><%=reg_name%>(<%=rs("reg_id")%>)</td>
								<td><%=rs("reg_date")%></td>
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
                        <a href="org_code_mg.asp?page=<%=first_page%>&org_gubun=<%=org_gubun%>&ck_sw=<%="y"%>&company=<%=company%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="org_code_mg.asp?page=<%=intstart -1%>&org_gubun=<%=org_gubun%>&ck_sw=<%="y"%>&company=<%=company%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="org_code_mg.asp?page=<%=i%>&org_gubun=<%=org_gubun%>&ck_sw=<%="y"%>&company=<%=company%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="org_code_mg.asp?page=<%=intend+1%>&org_gubun=<%=org_gubun%>&ck_sw=<%="y"%>&company=<%=company%>">[다음]</a> <a href="org_code_mg.asp?page=<%=total_page%>&org_gubun=<%=org_gubun%>&ck_sw=<%="y"%>&company=<%=company%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('org_code_add.asp?company=<%=company%>&org_gubun=<%=org_gubun%>','org_code_add_popup','scrollbars=yes,width=500,height=250')" class="btnType04">구분코드등록</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

