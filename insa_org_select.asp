<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

gubun = request("gubun")
mg_level = request("mg_level")
stock_level = request("stock_level")
view_condi=Request("view_condi")
target = Request("target")		'//2017-06-12 부모창이 리스트인 경우 선택 항목에 전달할때 필요
org_id = Request("org_id")

if gubun = "" then
   gubun = Request.Form("gubun")
   mg_level = Request.Form("mg_level")
   view_condi = Request.Form("view_condi")
end if

in_name = ""
If Request.Form("in_name")  <> "" Then
  in_name = Request.Form("in_name")
End If

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if view_condi = "" and in_name = "" then
	first_view = "N"
	sql = "select * from emp_org_mst where (org_name = '" + in_name + "')"
end if
if view_condi = "" and in_name <> "" then
	first_view = "Y"
	Sql = "select * from emp_org_mst where (org_name like '%" + in_name + "%') ORDER BY org_company,org_bonbu,org_saupbu,org_team,org_name ASC"
end if

if view_condi <> "" and in_name = "" then
	first_view = "N"
	sql = "select * from emp_org_mst where (org_company = '"&view_condi&"') and (org_name = '" + in_name + "')"
end if
if view_condi <> "" and in_name <> "" then
	first_view = "Y"
	Sql = "select * from emp_org_mst where (org_company = '"&view_condi&"') and (org_name like '%" + in_name + "%') "
	'//2017-09-25 폐쇄조직 제외
	sql = sql & " AND (org_end_date is null or org_end_date = '0000-00-00' or org_end_date = '') "
	sql = sql & " ORDER BY org_company,org_bonbu,org_saupbu,org_team,org_name ASC"
end if

rs.open SQL, DbConn, 1

'Response.write SQL
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
	    function orgsel(org_name,org_code,org_company,org_bonbu,org_saupbu,org_team,org_reside_place,org_reside_company,owner_empno,owner_empname,org_empno,org_emp_name,org_date,mg_level,org_cost_group,org_cost_center,gubun,org_id)
			{

				if(gubun =="owner"){
					opener.document.frm.owner_orgname.value = org_name;
					opener.document.frm.owner_org.value = org_code;
					opener.document.frm.org_company.value = org_company;
					opener.document.frm.org_bonbu.value = org_bonbu;
					opener.document.frm.org_saupbu.value = org_saupbu;
					opener.document.frm.org_team.value = org_team;
//					opener.document.frm.org_reside_place.value = org_reside_place;
//					opener.document.frm.org_reside_company.value = org_reside_company;
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
//					opener.document.frm.org_reside_place.value = org_reside_place;
//					opener.document.frm.org_reside_company.value = org_reside_company;
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
					if(org_saupbu =="")
					  if(mg_level =="사업부") {
					     opener.document.frm.org_saupbu.value = org_name;
					     opener.document.frm.org_saupbu<%=org_id%>.value = org_name;
					}
					if(org_team =="")
					  if(mg_level =="팀") {
					     opener.document.frm.org_team.value = org_name;
					     opener.document.frm.org_team<%=org_id%>.value = org_name;
					}
					window.close();
					opener.document.frm.org_owner_date.focus();
				}
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

					if(org_company =="")
					  if(mg_level =="회사")  {
					     opener.document.frm.emp_company.value = org_name;
					     opener.document.frm.emp_company<%=org_id%>.value = org_name;
					}
					if(org_bonbu =="")
					  if(mg_level =="본부")  {
					     opener.document.frm.emp_bonbu.value = org_name;
					     opener.document.frm.emp_bonbu<%=org_id%>.value = org_name;
					}
					if(org_saupbu =="")
					  if(mg_level =="사업부") {
					    opener.document.frm.emp_saupbu.value = org_name;
					    opener.document.frm.emp_saupbu<%=org_id%>.value = org_name;
					}
					if(org_team =="")
					  if(mg_level =="팀") {
					    opener.document.frm.emp_team.value = org_name;
					    opener.document.frm.emp_team<%=org_id%>.value = org_name;
					}
					window.close();
					opener.document.frm.emp_type.focus();
					opener.document.frm.emp_type<%=i%>.focus();

					}
				if(gubun =="apporg")
					{
					opener.document.frm.app_be_org.value = org_name;
					opener.document.frm.app_be_orgcode.value = org_code;
					opener.document.frm.app_company.value = org_company;
					opener.document.frm.app_bonbu.value = org_bonbu;
					opener.document.frm.app_saupbu.value = org_saupbu;
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
					opener.document.frm.app_saupbu<%=org_id%>.value = org_saupbu;
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
					if(org_saupbu =="")
					  if(mg_level =="사업부") {
					    opener.document.frm.app_saupbu.value = org_name;
					    opener.document.frm.app_saupbu<%=org_id%>.value = org_name;
					}
					if(org_team =="")
					  if(mg_level =="팀") {
					    opener.document.frm.app_team.value = org_name;
					    opener.document.frm.app_team<%=org_id%>.value = org_name;
					}
					window.close();
					opener.document.frm.app_comment.focus();
					}
				if(gubun =="appbmorg")
					{
					opener.document.frm.app_bm_org.value = org_name;
					opener.document.frm.app_bm_orgcode.value = org_code;
					opener.document.frm.app_bm_company.value = org_company;
					opener.document.frm.app_bm_bonbu.value = org_bonbu;
					opener.document.frm.app_bm_saupbu.value = org_saupbu;
					opener.document.frm.app_bm_team.value = org_team;
					opener.document.frm.app_bm_reside_place.value = org_reside_place;
					opener.document.frm.app_bm_reside_company.value = org_reside_company;
					opener.document.frm.app_bm_org_level.value = mg_level;

					opener.document.frm.app_bm_org<%=org_id%>.value = org_name;
					opener.document.frm.app_bm_orgcode<%=org_id%>.value = org_code;
					opener.document.frm.app_bm_company<%=org_id%>.value = org_company;
					opener.document.frm.app_bm_bonbu<%=org_id%>.value = org_bonbu;
					opener.document.frm.app_bm_saupbu<%=org_id%>.value = org_saupbu;
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
					if(org_saupbu =="")
					  if(mg_level =="사업부") {
					    opener.document.frm.app_bm_saupbu.value = org_name;
					    opener.document.frm.app_bm_saupbu<%=org_id%>.value = org_name;
					}
					if(org_team =="")
					  if(mg_level =="팀") {
					    opener.document.frm.app_bm_team.value = org_name;
					    opener.document.frm.app_bm_team<%=org_id%>.value = org_name;
					}					window.close();
					opener.document.frm.app_bm_comment.focus();
					}
                if(gubun =="stock")
					{
					opener.document.frm.stock_name.value = org_name;
					opener.document.frm.stock_code.value = org_code;
					opener.document.frm.stock_company.value = org_company;
					opener.document.frm.stock_bonbu.value = org_bonbu;
					opener.document.frm.stock_saupbu.value = org_saupbu;
					opener.document.frm.stock_team.value = org_team;
					opener.document.frm.stock_manager_code.value = org_empno;
					opener.document.frm.stock_manager_name.value = org_emp_name;
					opener.document.frm.stock_open_date.value = org_date;

					opener.document.frm.stock_name<%=org_id%>.value = org_name;
					opener.document.frm.stock_code<%=org_id%>.value = org_code;
					opener.document.frm.stock_company<%=org_id%>.value = org_company;
					opener.document.frm.stock_bonbu<%=org_id%>.value = org_bonbu;
					opener.document.frm.stock_saupbu<%=org_id%>.value = org_saupbu;
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
					if(org_saupbu =="")
					  if(mg_level =="사업부") {
					    opener.document.frm.stock_saupbu.value = org_name;
					    opener.document.frm.stock_saupbu<%=org_id%>.value = org_name;
					}
					if(org_team =="")
					  if(mg_level =="팀") {
					    opener.document.frm.stock_team.value = org_name;
					    opener.document.frm.stock_team<%=org_id%>.value = org_name;
					}
					window.close();
					opener.document.frm.stock_team.focus();
					}
				if(gubun =="stay")
					{
					opener.document.frm.stay_org_name.value = org_name;
					opener.document.frm.stay_org_code.value = org_code;
					opener.document.frm.stay_reside_company.value = org_reside_company;

					opener.document.frm.stay_org_name<%=org_id%>.value = org_name;
					opener.document.frm.stay_org_code<%=org_id%>.value = org_code;
					opener.document.frm.stay_reside_company<%=org_id%>.value = org_reside_company;

					window.close();
					opener.document.frm.stay_sido.focus();
					}
				if(gubun =="car")
					{
					opener.document.frm.car_use_dept.value = org_name;
					opener.document.frm.car_use_dept<%=org_id%>.value = org_name;
					window.close();
					opener.document.frm.emp_name.focus();
					}
				if(gubun =="alba")
					{
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
				if(gubun =="costEmp")
					{
					opener.callbackOrgSelect('<%=target%>',org_cost_group);
					window.close();
					}
				<%
				'else
				'	{
				'	opener.document.frm.sido.value = sido;
				'   opener.document.frm.family_gugun.value = gugun;
				'   opener.document.frm.family_dong.value = dong;
				'   opener.document.frm.family_zip.value = zip;
				'    window.close();
				'    opener.document.frm.family_addr.focus();
				'	}
				%>
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if(document.frm.in_name.value =="") {
					alert('조직명을 입력하세요');
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
				<form action="insa_org_select.asp?gubun=<%=gubun%>&mg_level=<%=mg_level%>&target=<%=target%>&org_id=<%=org_id%>" method="post" name="frm">
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
						if first_view = "Y" then
							do until rs.eof or rs.bof
							%>
							<tr>
								<td class="first">
                                <a href="#" onClick="orgsel('<%=rs("org_name")%>','<%=rs("org_code")%>','<%=rs("org_company")%>','<%=rs("org_bonbu")%>','<%=rs("org_saupbu")%>','<%=rs("org_team")%>','<%=rs("org_reside_place")%>','<%=rs("org_reside_company")%>','<%=rs("org_owner_empno")%>','<%=rs("org_owner_empname")%>','<%=rs("org_empno")%>','<%=rs("org_emp_name")%>','<%=rs("org_date")%>','<%=rs("org_level")%>','<%=rs("org_cost_group")%>','<%=rs("org_cost_center")%>','<%=gubun%>','<%=org_id%>');"><%=rs("org_name")%></a>
                                </td>
								<td><%=rs("org_code")%></td>
                                <td><%=rs("org_emp_name")%></td>
                                <td><%=rs("org_empno")%></td>
								<td class="left"><%=rs("org_company")%> - <%=rs("org_bonbu")%> - <%=rs("org_saupbu")%> - <%=rs("org_team")%></td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
						<%
						end if
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

