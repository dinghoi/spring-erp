<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim page, page_cnt, be_pg, curr_date
Dim start_page, pg_cnt, view_sort
Dim pgsize, stpage, date_sw
Dim whereSql, rsCount, orderSql
Dim total_record, total_page, view_condi
Dim title_line, view_user, condi
Dim rsEmp, rsOrg
Dim pg_url

page = f_Request("page")
page_cnt = f_Request("page_cnt")
pg_cnt = CInt(f_Request("pg_cnt"))
view_sort = f_Request("view_sort")
view_condi = f_Request("view_condi")
view_user = f_Request("view_user")
condi = f_Request("condi")

be_pg = "/insa/insa_mg.asp"
curr_date = DateValue(Mid(CStr(Now()), 1, 10))

If view_condi = "" Then
	'view_condi = "케이원"
	view_condi = "전체"
End If

' 화면 한 페이지
pgsize = 10

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_sort="&view_sort&"&view_condi="&view_condi

If view_sort = "" Then
	view_sort = "ASC"
End If

whereSql = "WHERE (isNull(emp_end_date) OR emp_end_date = '1900-01-01' OR emp_end_date = '0000-00-00') "
whereSql = whereSql&"AND emtt.emp_no < '900000' "
If view_condi <> "전체" Then
	whereSql = whereSql&"AND eomt.org_company = '"&view_condi&"' "
End If

If f_toString(condi, "") <> "" Then
	whereSql = whereSql&"AND emtt."&view_user

	If view_user = "emp_name" Then
		whereSql = whereSql&" LIKE '"&condi&"%' "
	Else
		whereSql = whereSql&" ='"&condi&"' "
	End If
End If

'if team <> "인사총무" Then
'    whereSql = whereSql & " AND emtt.emp_team = '"&team&"' "
'End If

objBuilder.Append "SELECT COUNT(*) FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code " & whereSql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount
rsCount.Close() : Set rsCount = Nothing

orderSql = "ORDER BY eomt.org_company, eomt.org_bonbu, eomt.org_team, "
orderSql = orderSql & "eomt.org_reside_place, emtt.emp_no, emtt.emp_in_date " & view_sort

objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_in_date, emtt.emp_org_name, emtt.emp_first_date, "
objBuilder.Append "	emtt.emp_org_baldate, emtt.emp_grade_date, emtt.emp_birthday, "
objBuilder.Append "	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "
objBuilder.Append "	eomt.org_name, eomt.org_code "
'objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team  "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append whereSql & orderSql & " LIMIT "& stpage & "," &pgsize

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = " 직원 현황 "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>인사관리 시스템</title>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>
	<script type="text/javascript">
		function getPageCode(){
			return "1 1";
		}

		function frmcheck(){
			if(formcheck(document.frm) && chkfrm()){
				document.frm.submit();
			}
		}

		function chkfrm(){
			if(document.frm.view_condi.value == ""){
				alert ("필드조건을 선택하시기 바랍니다");
				return false;
			}
			return true;
		}

	</script>
</head>
<body>
	<div id="wrap">
		<!--#include virtual = "/include/insa_header.asp" -->
		<!--#include virtual = "/include/insa_sub_menu1.asp" -->
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/insa/insa_mg.asp" method="post" name="frm">

			<fieldset class="srch">
				<legend>조회영역</legend>
				<dl>
					<dt>조건검색</dt>
					<dd>
						<p>
							<strong>회사 : </strong>
							<%
							'objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE org_level = '회사' ORDER BY org_code ASC "
							objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
							objBuilder.Append "	AND org_level = '회사' AND org_code <> '6272' "
							objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

							Set rsOrg = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()
							%>
							<label>
								<select name="view_condi" id="view_condi" style="width:150px;" onchange="frmcheck();">
									<option value="">전체</option>
								<%
								Do Until rsOrg.EOF
								%>
									<option value='<%=rsOrg("org_name")%>' <%If view_condi = rsOrg("org_name") then %>selected<% end if %>><%=rsOrg("org_name")%></option>
								<%
									rsOrg.MoveNext()
								Loop
								rsOrg.Close() : Set rsOrg = Nothing
								%>
								</select>
								<select name="view_user" id = "view_user" style="width:60px;">
									<option value="emp_name" <%If view_user = "emp_name" Then%>selected<%End If%>>성명</option>
									<option value="emp_no" <%If view_user = "emp_no" Then%>selected<%End If%>>사번</option>
								</select>
								<strong>조건 : </strong>
								<input type="text" name="condi" id="condi" style="width:100px;" value="<%=condi%>"/>
							</label>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
						</p>
					</dd>
				</dl>
			</fieldset>
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableList">
					<colgroup>
						<col width="5%" >
						<col width="5%" >
						<col width="6%" >
						<col width="6%" >
						<col width="6%" >
						<col width="6%" >
						<col width="9%" >
						<col width="6%" >
						<col width="6%" >
						<col width="6%" >
						<col width="6%" >
						<col width="*" >
						<col width="3%" >
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">사번</th>
							<th scope="col">성  명</th>
							<th scope="col">직급</th>
							<th scope="col">직위</th>
							<th scope="col">직책</th>
							<th scope="col">입사일</th>
							<th scope="col">소속</th>
							<th scope="col">최초입사일</th>
							<th scope="col">소속발령일</th>
							<th scope="col">승진일</th>
							<th scope="col">생년월일</th>
							<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
							<th scope="col">비고</th>
						</tr>
					</thead>
				<tbody>
					<%
					Dim emp_org_baldate, emp_grade_date

					Do Until rsEmp.EOF
						If rsEmp("emp_org_baldate") = "1900-01-01" Then
						   emp_org_baldate = ""
						Else
						   emp_org_baldate = rsEmp("emp_org_baldate")
						End If

						If rsEmp("emp_grade_date") = "1900-01-01" Then
						   emp_grade_date = ""
						Else
						   emp_grade_date = rsEmp("emp_grade_date")
						End If
					%>
					<tr>
						<td class="first"><%=rsEmp("emp_no")%></td>
						<td>
							<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsEmp("emp_no")%>','인사 기록카드','scrollbars=yes,width=1250,height=670')"><%=rsEmp("emp_name")%></a>
						</td>
						<td><%=rsEmp("emp_grade")%>&nbsp;</td>
						<td><%=rsEmp("emp_job")%>&nbsp;</td>
						<td><%=rsEmp("emp_position")%>&nbsp;</td>
						<td><%=rsEmp("emp_in_date")%>&nbsp;</td>
						<td><%=rsEmp("org_name")%>&nbsp;</td>
						<td><%=rsEmp("emp_first_date")%>&nbsp;</td>
						<td><%=emp_org_baldate%>&nbsp;</td>
						<td><%=emp_grade_date%>&nbsp;</td>
						<td><%=rsEmp("emp_birthday")%>&nbsp;</td>
						<td class="left">
						<%
						Call EmpOrgCodeSelect(rsEmp("org_code"))
						%>
						</td>
						<%
						If insa_grade = "0" Or SysAdminYN = "Y" Then
						%>
						<td>
							<a href="#" onClick="pop_Window('/insa/insa_emp_add01.asp?view_condi=<%=view_condi%>&emp_no=<%=rsEmp("emp_no")%>&u_type=U','인사기본사항 변경','scrollbars=yes,width=1250,height=600')">수정</a>
						</td>
						<%Else %>
							<td>&nbsp;</td>
						<%End If %>
					</tr>
					<%
						rsEmp.MoveNext()
					Loop
					rsEmp.close() : Set rsEmp = Nothing
					%>
					</tbody>
				</table>
			</div>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr>
				<td width="20%">
					<div class="btnCenter">
						<a href="/insa/insa_excel_emp.asp?view_condi=<%=view_condi%>" class="btnType04">엑셀다운로드</a>
					</div>
				</td>
				<td>
				<%
				'Page Navi
				Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
				DBConn.Close() : Set DBConn = Nothing
				%>
				</td>
				<td width="20%">
					<div class="btnCenter">
						<a href="#" onClick="pop_Window('/insa/insa_emp_add01.asp?view_condi=<%=view_condi%>','insa_emp_add01_popup','scrollbars=yes,width=1250,height=600')" class="btnType04">신규채용등록</a>
					</div>
				</td>
			  </tr>
			  </table>
		</form>
	</div>
</div>
<input type="hidden" name="user_id">
<input type="hidden" name="pass">
</body>
</html>