<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Repeat_Rows
Dim field_check
Dim field_view
Dim win_sw
dim company_tab(50,2)

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	company=Request("company")
	field_view=Request("field_view")
	field_check=Request("field_check")

 else
	company=Request.form("company")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if


If company = "" Then
	company = "01"
	field_check = "total"
End If

if asset_company <> "00" then
	company = asset_company
end if

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if company = "00" then
	com_sql = ""
  else
	com_sql = " and asset.company = '" + company + "' "
end if
if field_check = "total" then
	condi_sql = ""
  else
	condi_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

Sql = "SELECT count(*) FROM asset inner join asset_dept on (asset.company = asset_dept.company) and (asset.dept_code = asset_dept.dept_code) where asset.dept_code > '0' and (inst_process = 'Y') " + com_sql + condi_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF
order_sql = " order by asset_dept.org_first , asset_dept.org_second , asset_dept.dept_name , asset.gubun , asset.code_seq "

Sql = "SELECT * FROM asset inner join asset_dept on (asset.company = asset_dept.company) and (asset.dept_code = asset_dept.dept_code) where asset.dept_code > '0' and (inst_process = 'Y') " + com_sql + condi_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1

if company = "01" then
	title_01 = "법인명 / 지사명 / 지점명"
  else
	title_01 = "조직명1 / 조직명2 / 조직명3"
end if

title_line = "자산 세부 내역"
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
				<form action="asset_detail.asp" method="post" name="frm">
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
								<strong>필드조건</strong>
								  <%  if company = "01" then %>
                                        <select name="field_check" id="select5">
                                            <option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                            <option value="high_org" <% if field_check = "high_org" then %>selected<% end if %>>관리조직</option>
                                            <option value="org_first" <% if field_check = "org_first" then %>selected<% end if %>>법인명</option>
                                            <option value="org_second" <% if field_check = "org_second" then %>selected<% end if %>>지사명</option>
                                            <option value="dept_name" <% if field_check = "dept_name" then %>selected<% end if %>>지점명</option>
                                            <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>시도</option>
                                            <option value="tel_no" <% if field_check = "tel_no" then %>selected<% end if %>>전화번호</option>
                                            <option value="serial_no" <% if field_check = "serial_no" then %>selected<% end if %>>시리얼NO</option>
                                        </select>              
                                  <%	else  %>
                                        <select name="field_check" id="select5">
                                            <option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                            <option value="high_org" <% if field_check = "high_org" then %>selected<% end if %>>관리조직</option>
                                            <option value="org_first" <% if field_check = "org_first" then %>selected<% end if %>>조직명1</option>
                                            <option value="org_second" <% if field_check = "org_second" then %>selected<% end if %>>조직명1</option>
                                            <option value="dept_name" <% if field_check = "dept_name" then %>selected<% end if %>>조직명1</option>
                                            <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>시도</option>
                                            <option value="tel_no" <% if field_check = "tel_no" then %>selected<% end if %>>전화번호</option>
                                            <option value="serial_no" <% if field_check = "serial_no" then %>selected<% end if %>>시리얼NO</option>
                                        </select>              
                                  <%  end if  %>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:150px; text-align:left" >
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
							<col width="8%" >
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="15%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">소속회사</th>
								<th scope="col">관리조직</th>
								<th scope="col"><%=title_01%></th>
								<th scope="col">자산코드</th>
								<th scope="col">자산구분</th>
								<th scope="col">자산명</th>
								<th scope="col">자산번호</th>
								<th scope="col">시리얼NO</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

							etc_code = "75" + rs("company")
							Sql="select * from etc_code where etc_code = '" + etc_code + "'"
							Set rs_etc=DbConn.Execute(SQL)
							if rs_etc.eof or rs_etc.bof then
								company_name = "없음"
							  else
								company_name = rs_etc("etc_name")
							end if
							rs_etc.close()						
					
							gubun = "ERROR"
							if rs("gubun") = "01" then
							   gubun = "데스크탑"
							end if
							if rs("gubun") = "02" then
							   gubun = "모니터"
							end if
							if rs("gubun") = "03" then
							   gubun = "노트북"
							end if
							if rs("gubun") = "04" then
							   gubun = "프린터"
							end if
						%>
							<tr>
								<td class="first"><%=company_name%></td>
								<td><%=rs("high_org")%></td>
								<td><%=rs("org_first")%>&nbsp;/&nbsp;<%=rs("org_second")%>&nbsp;/&nbsp;<%=rs("dept_name")%></td>
								<td><%=rs("company")%>-<%=rs("gubun")%>-<%=rs("code_seq")%></td>
								<td><%=gubun%></td>
								<td><%=rs("asset_name")%></td>
								<td><%=mid(rs("asset_no"),1,2)%>-<%=mid(rs("asset_no"),3,6)%>-<%=right(rs("asset_no"),4)%></td>
								<td><%=rs("serial_no")%></td>
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
                    <a href = "asset_detail_excel.asp?company=<%=company%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                    </td>
				    <td>
                    <div id="paging">
                        <a href="asset_detail.asp?page=<%=first_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="asset_detail.asp?page=<%=intstart -1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="asset_detail.asp?page=<%=i%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="asset_detail.asp?page=<%=intend+1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[다음]</a> <a href="asset_detail.asp?page=<%=total_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

