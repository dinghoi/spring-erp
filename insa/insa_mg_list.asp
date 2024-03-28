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
Dim be_pg, page, view_condi, ck_sw, pgsize
Dim start_page, stpage, rsCount, total_record
Dim title_line, order_sql, where_sql, field_sql
Dim total_page, rsEmp
Dim emp_org_baldate, emp_birthday, emp_grade_date
Dim page_cnt, intstart, intend, first_page, i, pg_url

be_pg = "/insa/insa_mg_list.asp"

page = f_Request("page")
view_condi = f_Request("view_condi")

title_line = " 직원 현황 -인사자료미등록- "

If view_condi = "" Then
	view_condi = "emp_image"
End If

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi

order_sql = "ORDER BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_code, emtt.emp_in_date, emtt.emp_no ASC "
where_sql = "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01') AND (emtt.emp_no < '900000')"
field_sql = "AND (" & view_condi & " = '' OR isNull(" & view_condi & ")) "

objBuilder.Append "SELECT COUNT(*) FROM emp_master AS emtt " & where_sql & field_sql
Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT emp_org_baldate, emp_birthday, emp_grade_date, emp_no, emp_name, "
objBuilder.Append "	emp_grade, emp_position, emp_in_date, emp_org_name, emp_first_date, "
objBuilder.Append "	emp_reside_place, emp_company, emp_bonbu, emp_saupbu, emp_team, org_name, "
'objBuilder.Append "	org_company, org_bonbu, org_team, "
objBuilder.Append "	org_reside_place, eomt.org_code "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append where_sql & field_sql & order_sql & " LIMIT " & stpage & "," &pgsize

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>인사 관리 시스템</title>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>
	<script type="text/javascript">
		function getPageCode(){
			return "0 1";
		}

		function frmcheck(){
			if(formcheck(document.frm)){
				document.frm.submit();
			}
		}
		/*
		function delcheck(){
			if(form_chk(document.frm_del)){
				document.frm_del.submit();
			}
		}

		function form_chk(){
			a = confirm('삭제하시겠습니까?');

			if (a == true) {
				return true;
			}
			return false;
		}*/
	</script>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
	<div id="wrap">
		<!--#include virtual = "/include/insa_header.asp" -->
		<!--#include virtual = "/include/insa_report_menu.asp" -->
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/insa/insa_mg_list.asp" method="post" name="frm">
			<fieldset class="srch">
				<legend>조회영역</legend>
				<dl>
					<dt>조건 검색</dt>
					<dd>
						<p>
							<select name="view_condi" id="select3" style="width:100px;">
								<option value="cost_center" <%If view_condi = "cost_center" Then %>selected<%End If %>>비용배분구분</option>
								<option value="emp_image" <%If view_condi = "emp_image" Then %>selected<%End If %>>사진</option>
								<option value="emp_ename" <%If view_condi = "emp_ename" Then %>selected<%End If %>>성명(영문)</option>
								<option value="emp_person1" <%If view_condi = "emp_person1" Then %>selected<%End If %>>주민등록번호</option>
								<option value="emp_birthday" <%If view_condi = "emp_birthday" Then %>selected<%End If %>>생년월일</option>
								<option value="emp_sido" <%If view_condi = "emp_sido" Then %>selected<%End If %>>주소</option>
								<option value="emp_tel_no1" <%If view_condi = "emp_tel_no1" Then %>selected<%End If %>>전화번호</option>
								<option value="emp_hp_no1" <%If view_condi = "emp_hp_no1" Then %>selected<%End If %>>휴대폰번호</option>
								<option value="emp_emergency_tel" <%If view_condi = "emp_emergency_tel" Then %>selected<%End If %>>비상연락</option>
								<option value="emp_email" <%If view_condi = "emp_email" Then %>selected<%End If %>>이메일</option>
								<option value="emp_extension_no" <%If view_condi = "emp_extension_no" Then %>selected<%End If %>>내선번호</option>
								<option value="emp_last_edu" <%If view_condi = "emp_last_edu" Then %>selected<%End If %>>최종학력</option>
							</select>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
						</p>
					</dd>
				</dl>
			</fieldset>
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableList">
					<colgroup>
						<col width="6%" >
						<col width="6%" >
						<col width="6%" >
						<col width="6%" >
						<col width="6%" >
						<col width="6%" >
						<col width="9%" >
						<col width="6%" >
						<col width="6%" >
						<col width="10%" >
						<col width="*" >
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">사번</th>
							<th scope="col">성  명</th>
							<th scope="col">생년월일</th>
							<th scope="col">직급</th>
							<th scope="col">직책</th>
							<th scope="col">입사일</th>
							<th scope="col">소속</th>
							<th scope="col">최초입사일</th>
							<th scope="col">소속발령일</th>
							<th scope="col">상주처</th>
							<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
						</tr>
					</thead>
					<tbody>
					<%
					Do Until rsEmp.EOF
						If rsEmp("emp_org_baldate") = "1900-01-01" Then
						   emp_org_baldate = ""
						Else
						   emp_org_baldate = rsEmp("emp_org_baldate")
						End If

						If rsEmp("emp_birthday") = "1900-01-01" Then
						   emp_birthday = ""
						Else
						   emp_birthday = rsEmp("emp_birthday")
						End If

						If rsEmp("emp_grade_date") = "1900-01-01" Then
						   emp_grade_date = ""
						Else
						   emp_grade_date = rsEmp("emp_grade_date")
						End If
					%>
						<tr>
							<td class="first"><%=rsEmp("emp_no")%>&nbsp;</td>
							<td>
								<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsEmp("emp_no")%>','인사 기록 카드','scrollbars=yes,width=1250,height=670')"><%=rsEmp("emp_name")%></a>
							</td>
							<td><%=emp_birthday%>&nbsp;</td>
							<td><%=rsEmp("emp_grade")%>&nbsp;</td>
							<td><%=rsEmp("emp_position")%>&nbsp;</td>
							<td><%=rsEmp("emp_in_date")%>&nbsp;</td>
							<td><%=rsEmp("org_name")%>&nbsp;</td>
							<td><%=rsEmp("emp_first_date")%>&nbsp;</td>
							<td><%=emp_org_baldate%>&nbsp;</td>
							<td><%=rsEmp("org_reside_place")%>&nbsp;</td>
							<td class="left">
							<%
							Call EmpOrgCodeSelect(rsEmp("org_code"))
							%>
							</td>
						</tr>
					<%
						rsEmp.moveNext()
					Loop
					rsEmp.Close() : Set rsEmp = Nothing
					%>
					</tbody>
				</table>
			</div>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr>
				<td width="15%">
				<div class="btnCenter">
				<a href="/insa/insa_excel_emplist2.asp?view_condi=<%=view_condi%>" class="btnType04">엑셀다운로드</a>
				</div>
				</td>
				<td>
				<%
				'Page Navi
				Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
				DBConn.Close() : Set DBConn = Nothing
				%>
				</td>
			  </tr>
			  </table>
		</form>
	</div>
</div>
</body>
</html>

