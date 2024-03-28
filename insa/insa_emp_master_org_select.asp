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
Dim gubun, view_condi
Dim org_id, in_name, first_view
Dim rs, title_line

gubun = Request("gubun")
view_condi = Request("view_condi")
org_id = Request("org_id")

'mg_level = Request("mg_level")
'stock_level = Request("stock_level")
'target = Request("target")		'//2017-06-12 부모창이 리스트인 경우 선택 항목에 전달할때 필요

If gubun = "" Then
	gubun = Request.Form("gubun")
	'mg_level = Request.Form("mg_level")
	view_condi = Request.Form("view_condi")
End If

in_name = ""

If Request.Form("in_name")  <> "" Then
  in_name = Request.Form("in_name")
End If

title_line = "조직 검색"

objBuilder.Append "SELECT org_code, org_company, org_emp_name, org_empno, org_bonbu, org_saupbu, "
objBuilder.Append "	org_team, org_name, org_reside_place, org_reside_company, "
objBuilder.Append "	org_owner_empno, org_owner_empname, org_date, org_level, "
objBuilder.Append "	org_cost_group, org_cost_center "
objBuilder.Append "FROM emp_org_mst "
objBuilder.Append "WHERE (org_end_date IS NULL OR org_end_date = '0000-00-00' OR org_end_date = '1900-01-01') "

If view_condi = "" And in_name = "" Then
	first_view = "N"

	objBuilder.Append "AND org_name = '" & in_name & "' "
End If

If view_condi = "" And in_name <> "" Then
	first_view = "Y"

	objBuilder.Append "AND org_name LIKE '%" & in_name & "%' "
	objBuilder.Append "ORDER BY org_company, org_bonbu, org_saupbu, org_team, org_name ASC "
End If

If view_condi <> "" And in_name = "" Then
	first_view = "N"

	objBuilder.Append "AND org_company = '"&view_condi&"' AND org_name = '" & in_name & "' "
End If

If view_condi <> "" And in_name <> "" Then
	first_view = "Y"

	objBuilder.Append "AND org_company = '"&view_condi&"' AND org_name LIKE '%" & in_name & "%' "
	objBuilder.Append "ORDER BY org_company, org_bonbu, org_saupbu, org_team, org_name ASC "
End If

Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open objBuilder.toString(), DBConn, 1
objBuilder.Clear()

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
		function orgsel(gubun, org_code){
			var org_name, org_company, org_bonbu, org_saupbu
			var org_team, org_reside_place, org_reside_company
			var org_owner_empno, org_owner_empname, org_empno, org_emp_name
			var org_date, mg_level, org_cost_group, org_cost_center

			org_name = $('#o_name_' + org_code).val();
			org_company = $('#o_company_' + org_code).val();
			org_bonbu = $('#o_bonbu_' + org_code).val();
			org_saupbu = $('#o_saupbu_' + org_code).val();
			org_team = $('#o_team_' + org_code).val();
			org_reside_place = $('#o_reside_place_' + org_code).val();
			org_reside_company = $('#o_reside_company_' + org_code).val();
			org_owner_empno = $('#o_owner_empno_' + org_code).val();
			org_owner_empname = $('#o_owner_empname_' + org_code).val();
			org_empno = $('#o_empno_' + org_code).val();
			org_emp_name = $('#o_emp_name_' + org_code).val();
			org_date = $('#o_date_' + org_code).val();
			mg_level = $('#o_level_' + org_code).val();
			org_cost_group = $('#o_cost_group_' + org_code).val();
			org_cost_center = $('#o_cost_center_' + org_code).val();

			console.log(gubun);

			if(gubun =="org"){
				opener.document.frm.emp_org_name.value = org_name;
				opener.document.frm.emp_org_code.value = org_code;
				opener.document.frm.emp_company.value = org_company;
				opener.document.frm.emp_bonbu.value = org_bonbu;
				opener.document.frm.emp_saupbu.value = org_saupbu;
				opener.document.frm.emp_team.value = org_team;
				opener.document.frm.emp_reside_place.value = org_reside_place;
				opener.document.frm.emp_reside_company.value = org_reside_company;
				opener.document.frm.emp_org_level.value = mg_level;
				opener.document.frm.cost_center.value = org_cost_center;
				//opener.document.frm.cost_group.value = org_cost_group;

				opener.document.frm.emp_org_name<%=org_id%>.value = org_name;
				opener.document.frm.emp_org_code<%=org_id%>.value = org_code;
				opener.document.frm.emp_company<%=org_id%>.value = org_company;
				opener.document.frm.emp_bonbu<%=org_id%>.value = org_bonbu;
				opener.document.frm.emp_saupbu<%=org_id%>.value = org_saupbu;
				opener.document.frm.emp_team<%=org_id%>.value = org_team;
				opener.document.frm.emp_reside_place<%=org_id%>.value = org_reside_place;
				opener.document.frm.emp_reside_company<%=org_id%>.value = org_reside_company;
				opener.document.frm.emp_org_level<%=org_id%>.value = mg_level;
				opener.document.frm.cost_center<%=org_id%>.value = org_cost_center;
				//opener.document.frm.cost_group<%=org_id%>.value = org_cost_group;

				if(org_company =="" && mg_level === "회사"){
					 opener.document.frm.emp_company.value = org_name;
					 opener.document.frm.emp_company<%=org_id%>.value = org_name;
				}

				if(org_bonbu =="" && mg_level === "본부"){
					 opener.document.frm.emp_bonbu.value = org_name;
					 opener.document.frm.emp_bonbu<%=org_id%>.value = org_name;
				}

				if(org_bonbu =="" && mg_level === "사업부"){
					 opener.document.frm.emp_saupbu.value = org_name;
					 opener.document.frm.emp_saupbu<%=org_id%>.value = org_name;
				}

				if(org_team =="" && mg_level === "팀"){
					opener.document.frm.emp_team.value = org_name;
					opener.document.frm.emp_team<%=org_id%>.value = org_name;
				}

				window.close();
				opener.document.frm.emp_type.focus();
			}
		}

		function frmcheck(){
			if(formcheck(document.frm) && chkfrm()){
				document.frm.submit();
			}
		}

		function chkfrm(){
			if(document.frm.in_name.value ==""){
				alert('조직명을 입력하세요');
				frm.in_name.focus();
				return false;
			}
			return true;
		}
	</script>

</head>
<body oncontextmenu="return false" ondragstart="return false">
	<div id="container">
			<h3 class="insa"><%=title_line%></h3>
			<form action="/insa/popup/insa_emp_master_org_select.asp?gubun=<%=gubun%>&org_id=<%=org_id%>" method="post" name="frm">
			<fieldset class="srch">
				<legend>조회영역</legend>
				<dl>
					<dd>
						<p>
						<strong>조직명을 입력하세요 </strong>
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
						<col width="20%" >
						<col width="10%" >
						<col width="10%" >
						<col width="10%" >
						<col width="*" >
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">조직명</th>
							<th scope="col">조직코드</th>
							<th scope="col">조직장명</th>
							<th scope="col">조직장사번</th>
							<th scope="col">소속</th>
						</tr>
					</thead>
					<tbody>
					<%
					If first_view = "Y" Then
						Dim org_code

						Do Until rs.EOF Or rs.BOF
							org_code = rs("org_code")
						%>
						<tr>
							<td class="first">
								<a href="#" onclick="orgsel('<%=gubun%>', '<%=org_code%>');" ><%=rs("org_company")%></a>
							</td>
							<td><%=org_code%></td>
							<td><%=rs("org_emp_name")%></td>
							<td><%=rs("org_empno")%></td>
							<td class="left">
							<%
							Call EmpOrgInSaupbuText(rs("org_company"), rs("org_bonbu"), rs("org_saupbu"), rs("org_team"))
							%>
							(<%=rs("org_name")%>)
							</td>
						</tr>

						<input type = "hidden" id = "o_name_<%=org_code%>" value = "<%=rs("org_name")%>">
						<input type = "hidden" id = "o_code_<%=org_code%>" value = "<%=rs("org_code")%>">
						<input type = "hidden" id = "o_company_<%=org_code%>" value = "<%=rs("org_company")%>">
						<input type = "hidden" id = "o_bonbu_<%=org_code%>" value = "<%=rs("org_bonbu")%>">
						<input type = "hidden" id = "o_saupbu_<%=org_code%>" value = "<%=rs("org_saupbu")%>">
						<input type = "hidden" id = "o_team_<%=org_code%>" value = "<%=rs("org_team")%>">
						<input type = "hidden" id = "o_reside_place_<%=org_code%>" value = "<%=rs("org_reside_place")%>">
						<input type = "hidden" id = "o_reside_company_<%=org_code%>" value = "<%=rs("org_reside_company")%>">
						<input type = "hidden" id = "o_owner_empno_<%=org_code%>" value = "<%=rs("org_owner_empno")%>">
						<input type = "hidden" id = "o_owner_empname_<%=org_code%>" value = "<%=rs("org_owner_empname")%>">
						<input type = "hidden" id = "o_empno_<%=org_code%>" value = "<%=rs("org_empno")%>">
						<input type = "hidden" id = "o_emp_name_<%=org_code%>" value = "<%=rs("org_emp_name")%>">
						<input type = "hidden" id = "o_date_<%=org_code%>" value = "<%=rs("org_date")%>">
						<input type = "hidden" id = "o_level_<%=org_code%>" value = "<%=rs("org_level")%>">
						<input type = "hidden" id = "o_cost_group_<%=org_code%>" value = "<%=rs("org_cost_group")%>">
						<input type = "hidden" id = "o_cost_center_<%=org_code%>" value = "<%=rs("org_cost_center")%>">
						<%
							rs.MoveNext()
						Loop

						rs.Close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
					<%
					End If
					%>
					</tbody>
				</table>
			</div>
			<input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
			<!--<input type="hidden" name="mg_level" value="<%=mg_level%>" ID="Hidden1">-->
			<input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">
			</form>
	</div>
</body>
</html>

