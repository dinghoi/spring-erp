<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!DOCTYPE HTML>
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim gubun, first_view, emp_name
Dim rsEmp
Dim title_line

gubun = Request("gubun")
first_view = request("first_view")
emp_name = Request.Form("emp_name")

objBuilder.Append "SELECT emtt.emp_name, emtt.emp_no, emtt.emp_job, emtt.emp_grade, "
objBuilder.Append "	eomt.org_name, eomt.org_bonbu "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emtt.emp_pay_id <> '2' "

If emp_name = "" Then
	first_view = "N"

	'sql = "select * from emp_master where emp_name = '"&emp_name&"'"
	objBuilder.Append "AND emtt.emp_name = '"&emp_name&"'"
Else
	first_view = "Y"

	'sql = "select * from emp_master where emp_name like '%"&emp_name&"%' ORDER BY emp_name ASC"
	objBuilder.Append "AND emtt.emp_name LIKE '%"&emp_name&"%' ORDER BY emtt.emp_name ASC"
End If

Set rsEmp = Server.CreateObject("ADODB.RecordSet")
rsEmp.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

title_line = "사원 검색"
%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">-->
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
			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.emp_name.value ==""){
					alert('직원 이름을 입력하세요');
					frm.emp_name.focus();
					return false;
				}
				{
					return true;
				}
			}

			function emp_code(gubun,emp_name,emp_no,emp_job,org_name,saupbu){
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
				<form action="/insa/emp_search.asp?gubun=<%=gubun%>" method="post" name="frm">
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
						Dim ii

						If first_view = "Y" Then
							ii = 0

							Do Until rsEmp.EOF Or  rsEmp.BOF
								ii = ii + 1
							%>
							<tr>
								<td class="first"><a href="#" onClick="emp_code('<%=gubun%>','<%=rsEmp("emp_name")%>','<%=rsEmp("emp_no")%>','<%=rsEmp("emp_job")%>','<%=rsEmp("org_name")%>','<%=rsEmp("org_bonbu")%>');"><%=rsEmp("emp_name")%></a>
                                </td>
								<td><%=rsEmp("emp_no")%></td>
								<td><%=rsEmp("emp_grade")%></td>
								<td><%=rsEmp("org_bonbu")%>/<%=rsEmp("org_name")%></td>
							</tr>
							<%
								rsEmp.movenext()
							Loop
							rsEmp.close() : Set rsEmp = Nothing
							DBConn.Close() : Set DBConn = Nothing
							%>
						<%
						  Else
						%>
							<tr>
								<td class="first" colspan="4">내역이 없습니다</td>
							</tr>
                        <%
						End If
						%>
						</tbody>
					</table>
				</div>
				</form>
		</div>
	</body>
</html>

