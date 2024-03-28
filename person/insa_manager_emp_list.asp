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
Dim page, be_pg, view_condi, pgsize, start_page, stpage
Dim order_sql, where_sql, sqlStr, rsCount, total_record, total_page
Dim rsEmp, title_line, str_param, emp_name

emp_name =f_Request("emp_name")
page = f_Request("page")
be_pg = "/person/insa_manager_emp_list.asp"

view_condi = emp_company

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If
stpage = Int((page - 1) * pgsize)

str_param = "&bonbu="&bonbu&"&view_condi="&view_condi

where_sql = " WHERE (ISNULL(emp_end_date) OR emp_end_date = '1900-01-01') AND emp_no < '900000' "
where_sql = where_sql&"AND emp_bonbu = '"&bonbu&"' "

If f_toString(emp_name, "") <> "" Then
	where_sql = where_sql&"AND emp_name = '"&emp_name&"' "
End If

sqlStr = "SELECT COUNT(*) FROM emp_master "&where_sql
Set rsCount = Dbconn.Execute(sqlStr)

total_record = CInt(rsCount(0)) 'Result.RecordCount

If total_record Mod pgsize = 0 Then
	total_page = int(total_record / pgsize) 'Result.PageCount
Else
	total_page = int((total_record / pgsize) + 1)
End If

order_sql = "ORDER BY FIELD(emp_grade, '회장', '사장', '감사', '부사장', '전무이사', '상무이사', '연구소장', '이사', '전문위원',"
order_sql = order_sql & "'부장', '수석연구원', '차장', '책임연구원', '과장', '주임연구원', '대리', '대리1급', '대리2급', '연구원', '사원') ASC "

objBuilder.Append "SELECT emp_no, emp_name, emp_grade, emp_job, emp_position, emp_in_date, "
objBuilder.Append "	emp_org_code, emp_org_name, emp_first_date, emp_org_baldate, emp_grade_date, "
objBuilder.Append "	emp_birthday, emp_company, emp_bonbu, emp_saupbu, emp_team "
objBuilder.Append "FROM emp_master "
objBuilder.Append where_sql&order_sql
objBuilder.Append "LIMIT "&stpage&","&pgsize

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "직원 현황 - "&bonbu
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무관리</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.emp_name.value == ""){
					alert ("성명을 입력해 주세요.");
					return false;
				}
				return true;
			}
		</script>
	</head>
	<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">-->
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_plist_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/person/insa_manager_emp_list.asp" method="post" name="frm">

				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈조건 검색◈</dt>
						<dd>
							<p>
								<strong>성명 : </strong>
								<label>
									<input type="text" name="emp_name" id="emp_name" value="<%=emp_name%>" style="width:100px; text-align:left">
								</label>
								<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
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
                                <a href="#" onClick="pop_Window('/person/insa_individual_card00.asp?emp_no=<%=rsEmp("emp_no")%>','emp_card0_pop','scrollbars=yes,width=1300,height=650')"><%=rsEmp("emp_name")%></a>
								</td>
                                <td><%=rsEmp("emp_grade")%>&nbsp;</td>
                                <td><%=rsEmp("emp_job")%>&nbsp;</td>
                                <td><%=rsEmp("emp_position")%>&nbsp;</td>
                                <td><%=rsEmp("emp_in_date")%>&nbsp;</td>
                                <td><%=rsEmp("emp_org_name")%>&nbsp;</td>
                                <td><%=rsEmp("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=emp_grade_date%>&nbsp;</td>
                                <td><%=rsEmp("emp_birthday")%>&nbsp;</td>
                                <td class="left">
								<%Call EmpOrgCodeSelect(rsEmp("emp_org_code"))%>
								</td>
							</tr>
						<%
							rsEmp.MoveNext()
						Loop
						rsEmp.Close() : Set rsEmp = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<%
					'page navigator[허정호_20210720]
					Call Page_Navi(page, be_pg, str_param, total_page)

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