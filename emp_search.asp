<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim emp_name

gubun = request("gubun")
first_view = request("first_view")
emp_name = Request.Form("emp_name")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

SQL = "SELECT emtt.emp_name, emtt.emp_no, emtt.emp_job, emtt.emp_org_name, emtt.emp_saupbu, emtt.emp_grade, "
SQL = SQL & "	eomt.org_name, eomt.org_bonbu, eomt.org_saupbu "
SQL = SQL & "FROM emp_master AS emtt "
SQL = SQL & "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

if emp_name = "" then
	first_view = "N"
	'sql = "select * from emp_master where emp_name = '"&emp_name&"'"
	SQL = SQL & "WHERE emtt.emp_name = '"&emp_name&"' "
else
	first_view = "Y"
	'sql = "select * from emp_master where emp_name like '%"&emp_name&"%' ORDER BY emp_name ASC"
	SQL = SQL & "WHERE emtt.emp_name LIKE '%"&emp_name&"%' ORDER BY emtt.emp_name ASC "
end If

rs.Open SQL, Dbconn, 1

title_line = "사원 검색"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>직원 검색</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if(document.frm.emp_name.value =="") {
					alert('직원 이름을 입력하세요');
					frm.emp_name.focus();
					return false;}
				{
					return true;
				}
			}
			function emp_code(gubun,emp_name,emp_no,emp_job,org_name,saupbu)
			{
				if(gubun =="1")
				{
					opener.document.frm.emp_name.value = emp_name;
					opener.document.frm.emp_no.value = emp_no;
					opener.document.frm.emp_grade.value = emp_job;
					window.close();
				}
				else if(gubun =="2")
				{
					opener.document.frm.emp_name.value = emp_name;
					opener.document.frm.emp_no.value = emp_no;
					opener.document.frm.saupbu.value = saupbu;
					window.close();
				}
				else
				{
					opener.document.frm.emp_name.value = emp_name;
					opener.document.frm.emp_no.value = emp_no;
					opener.document.frm.emp_grade.value = emp_job;
					opener.document.frm.org_name.value = org_name;
					window.close();
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/emp_search.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>직원 이름을 입력하세요 </strong>
								<label>
        						<input name="emp_name" type="text" id="emp_name" value="<%=emp_name%>" style="width:150px;text-align:left; ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">이 름</th>
								<th scope="col">사원번호</th>
								<th scope="col">직 급</th>
								<th scope="col">사업부/부 서</th>
							</tr>
						</thead>
						<tbody>
						<%
						if first_view = "Y" then
							ii = 0
							do until rs.eof or rs.bof
								ii = ii + 1
							%>
							<tr>
								<td class="first"><a href="#" onClick="emp_code('<%=gubun%>','<%=rs("emp_name")%>','<%=rs("emp_no")%>','<%=rs("emp_job")%>','<%=rs("org_name")%>','<%=rs("org_bonbu")%>');"><%=rs("emp_name")%></a>
                                </td>
								<td><%=rs("emp_no")%></td>
								<td><%=rs("emp_grade")%></td>
								<td><%=rs("org_bonbu")%>/<%=rs("org_name")%></td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
						<%
						  else
						%>
							<tr>
								<td class="first" colspan="4">내역이 없습니다</td>
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

