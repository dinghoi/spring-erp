<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim gubun, view_condi, mg_level, in_name
Dim emp_rs, title_line, first_view, rsEmp

gubun = Request("gubun")
view_condi = Request("view_condi")

If gubun = "" Then
   gubun = Request.Form("gubun")
   mg_level = Request.Form("mg_level")
   view_condi = Request.Form("view_condi")
End If

title_line = " 조직장/직원 검색 "

in_name = ""

If Request.Form("in_name")  <> "" Then
  in_name = Request.Form("in_name")
End If

objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_job, "
objBuilder.Append "	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "
objBuilder.Append "	emtt.emp_reside_place, emtt.emp_reside_company, "
objBuilder.Append "	emtt.emp_org_code, emtt.emp_org_name, emtt.emp_in_date, emtt.emp_position, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "	eomt.org_reside_place, eomt.org_reside_company "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "

If view_condi = "" And in_name = "" Then
	first_view = "N"
	objBuilder.Append "AND emtt.emp_name = '" & in_name & "' "
End If

If view_condi = "" And in_name <> "" Then
	first_view = "Y"
	objBuilder.Append "AND emtt.emp_name LIKE '%" & in_name & "%' ORDER BY emtt.emp_name ASC "
End If

If view_condi <> "" And in_name = "" Then
	first_view = "N"
	objBuilder.Append "AND emtt.emp_name LIKE '%" & in_name & "%' "
End If

If view_condi <> "" And in_name <> "" Then
	first_view = "Y"
	objBuilder.Append "AND emtt.emp_name LIKE '%" & in_name & "%' ORDER BY emtt.emp_name ASC "
End If

Set rsEmp = Server.CreateObject("ADODB.RecordSet")
rsEmp.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>조직장 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function empsel(empno,empname,empjob,empcompany,empbonbu,empsaupbu,empteam,empresideplace,empresidecompany,emporgcode,emporgname,emp_in_date,emp_position,gubun){
				if(gubun =="orgemp")
					{
					opener.document.frm.org_empno.value = empno;
					opener.document.frm.org_empname.value = empname;
					window.close();
					opener.document.frm.owner_org.focus();
					}
				if(gubun =="car")
					{
					opener.document.frm.owner_emp_no.value = empno;
					opener.document.frm.emp_name.value = empname;
					opener.document.frm.emp_grade.value = empjob;
					opener.document.frm.emp_company.value = empcompany;
					opener.document.frm.emp_org_code.value = emporgcode;
					opener.document.frm.emp_org_name.value = emporgname;
					window.close();
					opener.document.frm.emp_name.focus();
					}
				if(gubun =="payexp")
					{
					opener.document.frm.ex_emp_no.value = empno;
					opener.document.frm.ex_emp_name.value = empname;
					opener.document.frm.ex_company.value = empcompany;
					opener.document.frm.ex_bonbu.value = empbonbu;
					opener.document.frm.ex_saupbu.value = empsaupbu;
					opener.document.frm.ex_team.value = empteam;
					opener.document.frm.ex_reside_place.value = empresideplace;
					opener.document.frm.ex_reside_company.value = empresidecompany;
					opener.document.frm.ex_org_code.value = emporgcode;
					opener.document.frm.ex_org_name.value = emporgname;
					window.close();
					opener.document.frm.rever_yymm.focus();
					}
                if(gubun =="stock")
					{
					opener.document.frm.stock_name.value = empname;
					opener.document.frm.stock_code.value = empno;
					opener.document.frm.stock_company.value = empcompany;
					opener.document.frm.stock_bonbu.value = empbonbu;
					opener.document.frm.stock_saupbu.value = empsaupbu;
					opener.document.frm.stock_team.value = empteam;
					opener.document.frm.stock_manager_code.value = empno;
					opener.document.frm.stock_manager_name.value = empname;
					opener.document.frm.stock_open_date.value = emp_in_date;
					window.close();
					opener.document.frm.stock_team.focus();
					}
				if(gubun =="st_emp1")
					{
					opener.document.frm.stock_go_name.value = empname;
					opener.document.frm.stock_go_man.value = empno;
					window.close();
					opener.document.frm.stock_go_man.focus();
					}
				if(gubun =="st_emp2")
					{
					opener.document.frm.stock_in_name.value = empname;
					opener.document.frm.stock_in_man.value = empno;
					window.close();
					opener.document.frm.stock_in_man.focus();
					}
				if(gubun =="chulgo01")
					{
					opener.document.frm.rele_emp_name.value = empname;
					opener.document.frm.rele_emp_no.value = empno;
					window.close();
					opener.document.frm.rele_stock_company.focus();
					}
				if(gubun =="holi")
					{
					opener.document.frm.holi_sing_empname.value = empname;
					opener.document.frm.holi_sign_empno.value = empno;
					opener.document.frm.holi_sign_org_name.value = emporgname;
					opener.document.frm.holi_sign_saupbu.value = empsaupbu;
					opener.document.frm.holi_sign_grade.value = empjob;
					opener.document.frm.holi_sign_position.value = emp_position;
					window.close();
					opener.document.frm.holi_type.focus();
					}
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if(document.frm.in_name.value =="") {
					alert('성명을 입력하세요');
					frm.in_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/popup/insa_emp_select.asp?gubun=<%=gubun%>" method="post" name="frm">
					<input type="hidden" name="emp_no" id="emp_no" value="<%=rsEmp("emp_no")%>">
					<input type="hidden" name="emp_name" id="emp_name" value="<%=rsEmp("emp_name")%>">
					<input type="hidden" name="emp_job" id="emp_job" value="<%=rsEmp("emp_job")%>">
					<input type="hidden" name="emp_company" id="emp_company" value="<%=rsEmp("emp_company")%>">
					<input type="hidden" name="emp_company" id="emp_company" value="<%=rsEmp("emp_company")%>">

				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>성명을 입력하세요 </strong>
								<label>
        						<input name="in_name" type="text" id="in_name" value="<%=in_name%>" style="width:150px; text-align:left; ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">성명</th>
								<th scope="col">사번</th>
								<th scope="col">현소속</th>
 							</tr>
						</thead>
						<tbody>
						<%
						if first_view = "Y" then
							do until rsEmp.EOF or rsEmp.BOF
						%>
							<tr>
								<td class="first">
									<a href="#" onClick="javascript:empsel('<%=rsEmp("emp_no")%>','<%=rsEmp("emp_name")%>','<%=rsEmp("emp_job")%>','<%=rsEmp("emp_company")%>','<%=rsEmp("emp_bonbu")%>','<%=rsEmp("emp_saupbu")%>','<%=rsEmp("emp_team")%>','<%=rsEmp("emp_reside_place")%>','<%=rsEmp("emp_reside_company")%>','<%=rsEmp("emp_org_code")%>','<%=rsEmp("emp_org_name")%>','<%=rsEmp("emp_in_date")%>','<%=rsEmp("emp_position")%>','<%=gubun%>');"><%=rsEmp("emp_name")%></a>
                                </td>
								<td><%=rsEmp("emp_no")%></td>
								<td class="left"><%=rsEmp("org_company")%> - <%=rsEmp("org_bonbu")%> - <%=rsEmp("org_saupbu")%> - <%=rsEmp("org_team")%> - <%=rsEmp("emp_position")%></td>



							</tr>
							<%
								rsEmp.movenext()
							Loop
							rsEmp.close() : Set rsEmp = Nothing
							DBConn.Close() : Set DBConn = Nothing
							%>
						<%
						  else
						%>
							<tr>
								<td class="first" colspan="3">내역이 없습니다</td>
							</tr>
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>
			</div>
		</div>
		<input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
		<input type="hidden" name="mg_level" value="<%=mg_level%>" ID="Hidden1">
		<input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">
	</form>
	</body>
</html>

