<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
gubun = request("gubun")
org_company = Request("org_company")

if org_company = "" then
	org_company = Request.form("org_company")
end if
org_name = Request.form("org_name")

if org_company = "" then
	'org_company = "케이원정보통신"
	org_company = "케이원"
end If

' 2019.02.22 박정신 요구 'N/W 1사업부','N/W 2사업부'는 나오지않도록 조건으로 처리..
if org_name = "" then
	SQL = "  SELECT *                                               " & chr(13) & _
		  "    FROM emp_org_mst                                     " & chr(13) & _
		  "   WHERE org_company = '"&org_company&"'                 " & chr(13) & _
		  "     AND org_name = '"&org_name&"'                       " & chr(13) & _
		  "     AND org_name not in ('N/W 1사업부','N/W 2사업부')   " & chr(13) & _
		  "     AND org_saupbu not in ('N/W 1사업부','N/W 2사업부') " & chr(13) & _
		  "ORDER BY org_name ASC"
 else
	SQL = "  SELECT *                                               " & chr(13) & _
	      "    FROM emp_org_mst                                     " & chr(13) & _
	      "   WHERE org_company = '"&org_company&"'                 " & chr(13) & _
		  "     AND org_name like '%" + org_name + "%'              " & chr(13) & _
		  "     AND org_name not in ('N/W 1사업부','N/W 2사업부')   " & chr(13) & _
		  "     AND org_saupbu not in ('N/W 1사업부','N/W 2사업부') " & chr(13) & _
	      "ORDER BY org_name ASC"
end if
'Response.write "<pre>"&Sql&"</pre>"
Rs.open SQL, Dbconn, 1

title_line = "조직 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>조직 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function org_list(gubun,emp_company,bonbu,saupbu,team,org_name,reside_place,reside_company)
			{
				if(gubun =="영업")
					{
					opener.document.frm.sales_company.value = emp_company;
					opener.document.frm.sales_bonbu.value = bonbu;
					opener.document.frm.sales_saupbu.value = saupbu;
					opener.document.frm.sales_team.value = team;
					opener.document.frm.sales_org_name.value = org_name;
					window.close();
					}
				else
					{
					opener.document.frm.emp_company.value = emp_company;
					opener.document.frm.bonbu.value = bonbu;
					opener.document.frm.saupbu.value = saupbu;
					opener.document.frm.team.value = team;
					opener.document.frm.org_name.value = org_name;
					opener.document.frm.reside_place.value = reside_place;
					opener.document.frm.reside_company.value = reside_company;
					window.close();
//					opener.document.frm.as_memo.focus();
					}
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if(document.frm.org_name.value =="") {
					alert('조직명을 입력하세요');
					frm.org_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/org_search.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>회사 : </strong>
						<% if gubun = "계산서" then	%>
							<%=org_company%>
							<input type="hidden" name="org_company" value="<%=org_company%>" ID="Hidden1">
                        <%   else	%>
							<label>
                            <select name="org_company" id="org_company" style="width:120px">
                              <%
                                'Sql="select * from emp_org_mst where org_level = '회사' order by org_name asc"
								sql = ""
								sql = "select org_name from emp_org_mst where (org_level = '회사') "
								sql = sql & "AND (org_end_date IS NULL OR org_end_date = '0000-00-00') "
								'sql = sql & "AND org_code > '6504' "
								sql = sql & "ORDER BY FIELD(org_company, '케이원') DESC, org_code DESC "
                                rs_org.Open Sql, Dbconn, 1
                                do until rs_org.eof
                                %>
                              <option value='<%=rs_org("org_name")%>' <%If org_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                              <%
                                    rs_org.movenext()
                                loop
                                rs_org.close()
                                %>
                            </select>
							</label>
                        <% end if	%>
							<strong>조직명 : </strong>
							<label>
        						<input name="org_name" type="text" id="org_name" value="<%=org_name%>" style="width:120px;text-align:left;ime-mode:active">
							</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">조직명</th>
								<th scope="col">소속</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
							org_company = rs("org_company")
							org_bonbu = rs("org_bonbu")
							org_saupbu = rs("org_saupbu")
							org_team = rs("org_team")
							org_name = rs("org_name")
						%>
							<tr>
								<td class="first">
                                <a href="#" onClick="org_list('<%=gubun%>','<%=org_company%>','<%=org_bonbu%>','<%=org_saupbu%>','<%=org_team%>','<%=rs("org_name")%>','<%=rs("org_reside_place")%>','<%=rs("org_reside_company")%>');"><%=rs("org_name")%></a>
                                </td>
								<td class="left">
									<%=org_company%>&nbsp;>&nbsp;<%=org_bonbu%>&nbsp;>&nbsp;<%'=org_saupbu%><%=org_team%>&nbsp;>&nbsp;(<%=org_name%>)
								</td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						if i = 0 then
						%>
							<tr>
								<td class="first" colspan="2">내역이 없습니다</td>
							</tr>
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</body>
</html>

