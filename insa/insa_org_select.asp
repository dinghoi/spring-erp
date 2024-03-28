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
Dim gubun, mg_level, stock_level, view_condi
Dim target, org_id, in_name, first_view
Dim rs, title_line

gubun = f_Request("gubun")
mg_level = f_Request("mg_level")
stock_level = f_Request("stock_level")
view_condi = f_Request("view_condi")
target = f_Request("target")		'//2017-06-12 부모창이 리스트인 경우 선택 항목에 전달할때 필요
org_id = f_Request("org_id")
in_name = f_Request("in_name")

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

	If view_condi = "전체" Then
		objBuilder.Append "AND org_name = '" & in_name & "' "
	Else
		objBuilder.Append "AND org_company = '"&view_condi&"' AND org_name = '" & in_name & "' "
	End If
End If

If view_condi <> "" And in_name <> "" Then
	first_view = "Y"

	'//2017-09-25 폐쇄조직 제외
	'sql = sql & " AND (org_end_date is null or org_end_date = '0000-00-00' or org_end_date = '') "
	'sql = sql & " ORDER BY org_company,org_bonbu,org_saupbu,org_team,org_name ASC"

	If view_condi = "전체" Then
		objBuilder.Append "AND org_name LIKE '%" & in_name & "%' "
	Else
		objBuilder.Append "AND org_company = '"&view_condi&"' AND org_name LIKE '%" & in_name & "%' "
	End If

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
	<title>인사 관리 시스템</title>
	<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>
	<script type="text/javascript" src="/java/js_window.js"></script>
	<script type="text/javascript">
		$(document).ready(function(){
			$('#in_name').focus();
		});

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

			//console.log(gubun);

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
				opener.document.frm.cost_group.value = org_cost_group;

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
				opener.document.frm.cost_group<%=org_id%>.value = org_cost_group;

				if(org_company =="" && mg_level === "회사"){
					 opener.document.frm.emp_company.value = org_name;
					 opener.document.frm.emp_company<%=org_id%>.value = org_name;
				}

				if(org_bonbu =="" && mg_level === "본부"){
					 opener.document.frm.emp_bonbu.value = org_name;
					 opener.document.frm.emp_bonbu<%=org_id%>.value = org_name;
				}

				if(org_team =="" && mg_level === "팀"){
					opener.document.frm.emp_team.value = org_name;
					opener.document.frm.emp_team<%=org_id%>.value = org_name;
				}

				window.close();
				opener.document.frm.emp_type.focus();
			}

			if(gubun =="owner"){
				opener.document.frm.owner_orgname.value = org_name;
				opener.document.frm.owner_org.value = org_code;
				opener.document.frm.org_company.value = org_company;
				opener.document.frm.org_bonbu.value = org_bonbu;
				opener.document.frm.org_saupbu.value = org_saupbu;
				opener.document.frm.org_team.value = org_team;
				//opener.document.frm.org_reside_place.value = org_reside_place;
				//opener.document.frm.org_reside_company.value = org_reside_company;
				opener.document.frm.owner_empno.value = org_empno;
				opener.document.frm.owner_empname.value = org_emp_name;

				opener.document.frm.owner_orgname<%=org_id%>.value = org_name;
				opener.document.frm.owner_org<%=org_id%>.value = org_code;
				opener.document.frm.org_company<%=org_id%>.value = org_company;
				opener.document.frm.org_bonbu<%=org_id%>.value = org_bonbu;
				opener.document.frm.org_saupbu<%=org_id%>.value = org_saupbu;
				opener.document.frm.org_team<%=org_id%>.value = org_team;
				opener.document.frm.owner_empno<%=org_id%>.value = org_empno;
				opener.document.frm.owner_empname<%=org_id%>.value = org_emp_name;

				if(org_company =="") {
				  if(mg_level =="회사")  {
					 opener.document.frm.org_company.value = org_name;
					 opener.document.frm.org_company<%=org_id%>.value = org_name;
					}
				}
				if(org_bonbu =="") {
				  if(mg_level =="본부")  {
					 opener.document.frm.org_bonbu.value = org_name;
					 opener.document.frm.org_bonbu<%=org_id%>.value = org_name;
					}
				}

				if(org_saupbu =="") {
				  if(mg_level =="사업부") {
					 opener.document.frm.org_saupbu.value = org_name;
					 opener.document.frm.org_saupbu<%=org_id%>.value = org_name;
					}
				}

				if(org_team =="") {
				  if(mg_level =="팀") {
					 opener.document.frm.org_team.value = org_name;
					 opener.document.frm.org_team<%=org_id%>.value = org_name;
					}
				}
				window.close();
				opener.document.frm.tel_ddd.focus();
			}

			if(gubun =="owner2"){
				opener.document.frm.owner_orgname.value = org_name;
				opener.document.frm.owner_org.value = org_code;
				opener.document.frm.org_company.value = org_company;
				opener.document.frm.org_bonbu.value = org_bonbu;
				opener.document.frm.org_saupbu.value = org_saupbu;
				opener.document.frm.org_team.value = org_team;
				//opener.document.frm.org_reside_place.value = org_reside_place;
				//opener.document.frm.org_reside_company.value = org_reside_company;
				opener.document.frm.owner_empno.value = org_empno;
				opener.document.frm.owner_empname.value = org_emp_name;

				opener.document.frm.owner_orgname<%=org_id%>.value = org_name;
				opener.document.frm.owner_org<%=org_id%>.value = org_code;
				opener.document.frm.org_company<%=org_id%>.value = org_company;
				opener.document.frm.org_bonbu<%=org_id%>.value = org_bonbu;
				opener.document.frm.org_saupbu<%=org_id%>.value = org_saupbu;
				opener.document.frm.org_team<%=org_id%>.value = org_team;
				opener.document.frm.owner_empno<%=org_id%>.value = org_empno;
				opener.document.frm.owner_empname<%=org_id%>.value = org_emp_name;

				console.log(org_empno);

				if(org_company =="")
				  if(mg_level =="회사")  {
					 opener.document.frm.org_company.value = org_name;
					 opener.document.frm.org_company<%=org_id%>.value = org_name;
				}
				if(org_bonbu =="")
				  if(mg_level =="본부")  {
					 opener.document.frm.org_bonbu.value = org_name;
					 opener.document.frm.org_bonbu<%=org_id%>.value = org_name;
				}


				if(org_saupbu ==""){
					if(mg_level =="사업부"){
						opener.document.frm.org_saupbu.value = org_name;
						opener.document.frm.org_saupbu<%=org_id%>.value = org_name;
					}
				}

				if(org_team =="")
				  if(mg_level =="팀") {
					 opener.document.frm.org_team.value = org_name;
					 opener.document.frm.org_team<%=org_id%>.value = org_name;
				}
				window.close();
				opener.document.frm.org_owner_date.focus();
			}

			if(gubun =="apporg"){
				opener.document.frm.app_be_org.value = org_name;
				opener.document.frm.app_be_orgcode.value = org_code;
				opener.document.frm.app_company.value = org_company;
				opener.document.frm.app_bonbu.value = org_bonbu;
				//opener.document.frm.app_saupbu.value = org_saupbu;
				opener.document.frm.app_team.value = org_team;
				opener.document.frm.app_reside_place.value = org_reside_place;
				opener.document.frm.app_reside_company.value = org_reside_company;
				opener.document.frm.app_org_level.value = mg_level;
				opener.document.frm.cost_center.value = org_cost_center;
				opener.document.frm.app_cost_group.value = org_cost_group;

				opener.document.frm.app_be_org<%=org_id%>.value = org_name;
				opener.document.frm.app_be_orgcode<%=org_id%>.value = org_code;
				opener.document.frm.app_company<%=org_id%>.value = org_company;
				opener.document.frm.app_bonbu<%=org_id%>.value = org_bonbu;
				//opener.document.frm.app_saupbu<%=org_id%>.value = org_saupbu;
				opener.document.frm.app_team<%=org_id%>.value = org_team;
				opener.document.frm.app_reside_place<%=org_id%>.value = org_reside_place;
				opener.document.frm.app_reside_company<%=org_id%>.value = org_reside_company;
				opener.document.frm.app_org_level<%=org_id%>.value = mg_level;
				opener.document.frm.cost_center<%=org_id%>.value = org_cost_center;
				opener.document.frm.app_cost_group<%=org_id%>.value = org_cost_group;

				if(org_company =="")
				  if(mg_level =="회사")  {
					 opener.document.frm.app_company.value = org_name;
					 opener.document.frm.app_company<%=org_id%>.value = org_name;
				}
				if(org_bonbu =="")
				  if(mg_level =="본부")  {
					 opener.document.frm.app_bonbu.value = org_name;
					 opener.document.frm.app_bonbu<%=org_id%>.value = org_name;
				}
				/*
				if(org_saupbu =="")
				  if(mg_level =="사업부") {
					opener.document.frm.app_saupbu.value = org_name;
					opener.document.frm.app_saupbu<%=org_id%>.value = org_name;
				}*/
				if(org_team =="")
				  if(mg_level =="팀") {
					opener.document.frm.app_team.value = org_name;
					opener.document.frm.app_team<%=org_id%>.value = org_name;
				}
				window.close();
				opener.document.frm.app_comment.focus();
			}

			if(gubun =="appbmorg"){
				opener.document.frm.app_bm_org.value = org_name;
				opener.document.frm.app_bm_orgcode.value = org_code;
				opener.document.frm.app_bm_company.value = org_company;
				opener.document.frm.app_bm_bonbu.value = org_bonbu;
				//opener.document.frm.app_bm_saupbu.value = org_saupbu;
				opener.document.frm.app_bm_team.value = org_team;
				opener.document.frm.app_bm_reside_place.value = org_reside_place;
				opener.document.frm.app_bm_reside_company.value = org_reside_company;
				opener.document.frm.app_bm_org_level.value = mg_level;

				opener.document.frm.app_bm_org<%=org_id%>.value = org_name;
				opener.document.frm.app_bm_orgcode<%=org_id%>.value = org_code;
				opener.document.frm.app_bm_company<%=org_id%>.value = org_company;
				opener.document.frm.app_bm_bonbu<%=org_id%>.value = org_bonbu;
				//opener.document.frm.app_bm_saupbu<%=org_id%>.value = org_saupbu;
				opener.document.frm.app_bm_team<%=org_id%>.value = org_team;
				opener.document.frm.app_bm_reside_place<%=org_id%>.value = org_reside_place;
				opener.document.frm.app_bm_reside_company<%=org_id%>.value = org_reside_company;
				opener.document.frm.app_bm_org_level<%=org_id%>.value = mg_level;

				if(org_company =="")
				  if(mg_level =="회사")  {
					 opener.document.frm.app_bm_company.value = org_name;
					 opener.document.frm.app_bm_company<%=org_id%>.value = org_name;
				}
				if(org_bonbu =="")
				  if(mg_level =="본부")  {
					 opener.document.frm.app_bm_bonbu.value = org_name;
					 opener.document.frm.app_bm_bonbu<%=org_id%>.value = org_name;
				}
				/*
				if(org_saupbu =="")
				  if(mg_level =="사업부") {
					opener.document.frm.app_bm_saupbu.value = org_name;
					opener.document.frm.app_bm_saupbu<%=org_id%>.value = org_name;
				}*/
				if(org_team =="")
				  if(mg_level =="팀") {
					opener.document.frm.app_bm_team.value = org_name;
					opener.document.frm.app_bm_team<%=org_id%>.value = org_name;
				}					window.close();
				opener.document.frm.app_bm_comment.focus();
			}

			if(gubun =="stock"){
				opener.document.frm.stock_name.value = org_name;
				opener.document.frm.stock_code.value = org_code;
				opener.document.frm.stock_company.value = org_company;
				opener.document.frm.stock_bonbu.value = org_bonbu;
				//opener.document.frm.stock_saupbu.value = org_saupbu;
				opener.document.frm.stock_team.value = org_team;
				opener.document.frm.stock_manager_code.value = org_empno;
				opener.document.frm.stock_manager_name.value = org_emp_name;
				opener.document.frm.stock_open_date.value = org_date;

				opener.document.frm.stock_name<%=org_id%>.value = org_name;
				opener.document.frm.stock_code<%=org_id%>.value = org_code;
				opener.document.frm.stock_company<%=org_id%>.value = org_company;
				opener.document.frm.stock_bonbu<%=org_id%>.value = org_bonbu;
				//opener.document.frm.stock_saupbu<%=org_id%>.value = org_saupbu;
				opener.document.frm.stock_team<%=org_id%>.value = org_team;
				opener.document.frm.stock_manager_code<%=org_id%>.value = org_empno;
				opener.document.frm.stock_manager_name<%=org_id%>.value = org_emp_name;
				opener.document.frm.stock_open_date<%=org_id%>.value = org_date;

				if(org_company =="")
				  if(mg_level =="회사")  {
					 opener.document.frm.stock_company.value = org_name;
					 opener.document.frm.stock_company<%=org_id%>.value = org_name;
				}
				if(org_bonbu =="")
				  if(mg_level =="본부")  {
					 opener.document.frm.stock_bonbu.value = org_name;
					 opener.document.frm.stock_bonbu<%=org_id%>.value = org_name;
				}
				/*
				if(org_saupbu =="")
				  if(mg_level =="사업부") {
					opener.document.frm.stock_saupbu.value = org_name;
					opener.document.frm.stock_saupbu<%=org_id%>.value = org_name;
				}*/
				if(org_team =="")
				  if(mg_level =="팀") {
					opener.document.frm.stock_team.value = org_name;
					opener.document.frm.stock_team<%=org_id%>.value = org_name;
				}
				window.close();
				opener.document.frm.stock_team.focus();
			}

			if(gubun =="stay"){
				opener.document.frm.stay_org_name.value = org_name;
				opener.document.frm.stay_org_code.value = org_code;
				opener.document.frm.stay_reside_company.value = org_reside_company;

				opener.document.frm.stay_org_name<%=org_id%>.value = org_name;
				opener.document.frm.stay_org_code<%=org_id%>.value = org_code;
				opener.document.frm.stay_reside_company<%=org_id%>.value = org_reside_company;

				window.close();
				opener.document.frm.stay_sido.focus();
			}

			if(gubun =="car"){
				opener.document.frm.car_use_dept.value = org_name;
				opener.document.frm.car_use_dept<%=org_id%>.value = org_name;
				window.close();
				opener.document.frm.emp_name.focus();
			}

			if(gubun =="alba"){
				opener.document.frm.org_name.value = org_name;
				opener.document.frm.company.value = org_company;
				opener.document.frm.bonbu.value = org_bonbu;
				opener.document.frm.saupbu.value = org_saupbu;
				opener.document.frm.team.value = org_team;

				opener.document.frm.org_name<%=org_id%>.value = org_name;
				opener.document.frm.company<%=org_id%>.value = org_company;
				opener.document.frm.bonbu<%=org_id%>.value = org_bonbu;
				opener.document.frm.saupbu<%=org_id%>.value = org_saupbu;
				opener.document.frm.team<%=org_id%>.value = org_team;

				window.close();
				opener.document.frm.person_no1.focus();
			}

			if(gubun =="costEmp"){
				//alert(org_code);
				opener.callbackOrgSelect('<%=target%>',org_cost_group);
				opener.callbackOrgCodeSelect('<%=target%>',org_code);
				window.close();
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
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/insa/insa_org_select.asp?gubun=<%=gubun%>&mg_level=<%=mg_level%>&target=<%=target%>&org_id=<%=org_id%>" method="post" name="frm">
			<fieldset class="srch">
				<legend>조회영역</legend>
				<dl>
					<dd>
						<p>
						<strong>조직명을 입력하세요 </strong>
							<label>
								<input type="text" name="in_name" id="in_name" value="<%=in_name%>" style="width:150px; text-align:left; ime-mode:active;">
							</label>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
						</p>
					</dd>
				</dl>
			</fieldset>
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableList">
					<colgroup>
						<col width="10%" >
						<col width="10%" >
						<col width="10%" >
						<col width="20%" >
						<col width="*" >
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">조직코드</th>
							<th scope="col">소속회사</th>
							<th scope="col">조직구분</th>
							<th scope="col">조직명</th>
							<th scope="col">소속[상주처](상주회사)</th>
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
							<td class="first"><%=org_code%></td>
							<td><%=rs("org_company")%></td>
							<td><%=rs("org_level")%></td>
							<td><a href="#" onclick="orgsel('<%=gubun%>', '<%=org_code%>');" ><%=rs("org_name")%></a></td>
							<td class="left">
							<%
							Call EmpOrgCodeSelect(rs("org_code"))
							%>
								[<%=rs("org_reside_place")%>]
								(<%=rs("org_reside_company")%>)
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
					Else
					%>
						<tr><td colspan="5" style="height:30px;">조회된 내역이 없습니다.</td></tr>
					<%
					End If
					%>
					</tbody>
				</table>
			</div>
			<input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
			<input type="hidden" name="mg_level" value="<%=mg_level%>" ID="Hidden1">
			<input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">
			</form>
	</div>
</body>
</html>