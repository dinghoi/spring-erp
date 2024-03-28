<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim be_pg, page, view_condi, condi, ck_sw
Dim view_company, title_line, condi_sql, pgsize
Dim pasize, start_page, stpage, be_page
Dim rsCount, totCount, total_page
Dim rsCareer, pg_url
Dim career_empno, task_memo
Dim emp_name, emp_grade, emp_job, emp_position
Dim emp_org_code, emp_org_name, view_memo
Dim page_cnt

be_page = "/insa/insa_career_list.asp"

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")

If f_toString(view_condi, "") = "" Then
	view_condi = "전체"
	condi_sql = " "
	condi = ""
End If

title_line = " 직원 경력 현황 "
pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&condi="&condi

objBuilder.Append "SELECT COUNT(*) FROM emp_career AS emct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emct.career_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "

If view_condi = "상주처회사" Then
	objBuilder.Append "AND eomt.org_reside_company  "
ElseIf view_condi = "경력업무" then
	objBuilder.Append "AND emct.career_task "
Else
	objBuilder.Append "AND emtt.emp_name "
End If

objBuilder.Append "LIKE '%"&condi&"%' "

Set rsCount = Dbconn.Execute(objBuilder.ToString())
objBuilder.Clear()

totCount = CInt(rsCount(0))

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT emct.career_task, emct.career_empno, emct.career_office, emct.career_join_date, "
objBuilder.Append "	emct.career_end_date, emct.career_dept, emct.career_position, "
objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_org_code, emtt.emp_org_name, emtt.emp_company, "
objBuilder.Append "	eomt.org_name, eomt.org_company "
objBuilder.Append "FROM emp_career AS emct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emct.career_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "

If view_condi = "상주처회사" Then
	objBuilder.Append "AND emtt.emp_reside_company  "
ElseIf view_condi = "경력업무" then
	objBuilder.Append "AND emct.career_task "
Else
	objBuilder.Append "AND emtt.emp_name "
End If

objBuilder.Append "LIKE '%"&condi&"%' "
objBuilder.Append "ORDER BY emct.career_empno ASC "
objBuilder.Append "LIMIT "&stpage& ", "&pgsize

Set rsCareer = DBConn.Execute(objBuilder.ToString())
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
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}

			function form_chk(){
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}*/
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_career_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
									<option value="전체" <%If view_condi = "전체" Then %>selected<%End If %>>전체</option>
									<option value="경력업무" <%If view_condi = "경력업무" Then %>selected<%End If %>>경력업무</option>
									<option value="상주처회사" <%If view_condi = "상주처회사" Then %>selected<%End If %>>상주처회사</option>
                                </select>
								<strong>조건 : </strong>
								<input type="text" name="condi" value="<%=condi%>" style="width:150px; text-align:left;"/>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="17%" >
                            <col width="14%" >
                            <col width="12%" >
                            <col width="10%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
                                <th scope="col">성명</th>
                                <th scope="col">직위</th>
								<th scope="col">회사</th>
								<th scope="col">소속</th>
                                <th scope="col">경력회사</th>
								<th scope="col">재직기간</th>
								<th scope="col">부서</th>
								<th scope="col">직위</th>
								<th scope="col">주요업무</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsCareer.EOF
							career_empno = rsCareer("career_empno")
							emp_name = rsCareer("emp_name")
							emp_grade = rsCareer("emp_grade")
							emp_job = rsCareer("emp_job")
							emp_position = rsCareer("emp_position")
							emp_org_code = rsCareer("emp_org_code")
							emp_org_name = rsCareer("org_name")
							emp_company = rsCareer("org_company")
							task_memo = Replace(rsCareer("career_task"), Chr(34), Chr(39))
							view_memo = task_memo

							If Len(task_memo) > 10 Then
								view_memo = Mid(task_memo, 1, 10) & ".."
							End If
	           			%>
							<tr>
								<td><%=rsCareer("career_empno")%>&nbsp;</td>
                                <td>
									<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsCareer("career_empno")%>','인사 기록 카드','scrollbars=yes,width=1250,height=670')"><%=emp_name%></a>
								</td>
                                <td><%=emp_job%>&nbsp;</td>
                                <td><%=emp_company%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rsCareer("career_office")%>&nbsp;</td>
                                <td><%=rsCareer("career_join_date")%>∼<%=rsCareer("career_end_date")%>&nbsp;</td>
                                <td><%=rsCareer("career_dept")%>&nbsp;</td>
                                <td><%=rsCareer("career_position")%>&nbsp;</td>
                                <td class="left"><p><span title="<%=task_memo%>"><%=view_memo%></span></p></td>
							</tr>
						<%
							rsCareer.MoveNext()
						Loop
						rsCareer.Close() : Set rsCareer = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="/insa/insa_excel_careerlist.asp?view_condi=<%=view_condi%>&condi=<%=condi%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, totCount, pgsize)
					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>