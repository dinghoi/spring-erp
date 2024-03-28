<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim dept_name
Dim rs
Dim rs_numRows

dept_name = ""
If (request.form("dept_name")  <> "") Then 
  dept_name = request.form("dept_name") 
End If
company = request("company")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if dept_name = "" then
	sql = "select * from asset_dept where company = '" + company + "' and dept_name  = 'none'"
  else
	Sql = "select company, dept_code, sido, gugun, dong, org_first, concat(ifnull(org_second,' '),' ',ifnull(dept_name,' ')) as org_name from asset_dept where company = '" + company + "' and concat(ifnull(org_second,' '),' ',ifnull(dept_name,' ')) like '%" + dept_name + "%' ORDER BY dept_name ASC"
end if
rs.open SQL, DbConn, 1

title_line = "조직코드 조회"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>조직코드 조회</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function deptcode(dept_code,org_name)
			{
				opener.document.frm.dept_code.value = dept_code;
				opener.document.frm.dept_name.value = org_name;
				window.close();
				opener.document.frm.user_name.focus();
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.dept_name.value =="") {
					alert('부서를 입력하세요');
					frm.dept_name.focus();
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
				<form action="dept_search.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>조직명 : </strong>	
                            <label>
                              <input name="dept_name" type="text" id="dept_name" value="<%=dept_name%>" style="width:150px">
                              <input name="company" type="hidden" id="company" value="<%=company%>">
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
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">코드</th>
								<th scope="col">조직명</th>
								<th scope="col">주소</th>
							</tr>
						</thead>
						<tbody>
						<% 
                        i = 0
                        do until rs.eof
                            i = i + 1
                        %>
							<tr>
								<td class="first"><%=rs("dept_code")%></td>
								<td><a href="#" onClick="deptcode('<%=rs("dept_code")%>','<%=rs("org_name")%>');"><%=rs("org_name")%></a></td>
								<td class="left"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						if  i = 0 and dept_name <> "" then
							msg = "내역이 없습니다 !!!"
						  else
							msg = ""
						end if
						%>
							<tr>
								<td class="first" colspan="3"><%=msg%></td>
							</tr>
						</tbody>
					</table>
				</div>
			</div>				
	</div>        				
	</form>
	</body>
</html>

