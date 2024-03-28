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
Dim in_name, in_empno
Dim be_pg, page, view_condi, condi, ck_sw
Dim condi_sql, pgsize, start_page, stpage, title_line
Dim rsCount, total_record, total_page, rsSch
Dim pg_url, sch_empno, emp_name, emp_grade, emp_job, emp_position
Dim emp_org_code, emp_org_name, page_cnt

be_pg = "/insa/insa_school_list.asp"

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")

If view_condi = "" Then
	view_condi = "전체"
	condi_sql = " "
	condi = ""
End If

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&condi="&condi

Select Case view_condi
	Case "전체"
		condi_sql = ""
	Case "상주처회사"
		condi_sql = "AND emtt.emp_reside_company LIKE '%"&condi&"%' "
	Case "성명"
		condi_sql = "AND emtt.emp_name LIKE '%"&condi&"%' "
	Case Else
		condi_sql = "AND emct." & view_condi & " LIKE '%"&condi&"%' "
End Select

objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM emp_school AS emct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emct.sch_empno  = emtt.emp_no "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01') "
objBuilder.Append condi_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0))

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT emct.sch_empno, emct.sch_school_name, emct.sch_start_date, emct.sch_end_date, "
objBuilder.Append "	emct.sch_dept, emct.sch_major, emct.sch_sub_major, emct.sch_degree, "
objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, emtt.emp_org_code, "
objBuilder.Append "	eomt.org_name, eomt.org_company "
objBuilder.Append "FROM emp_school AS emct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emct.sch_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01') "
objBuilder.Append condi_sql
objBuilder.Append "ORDER BY emct.sch_empno ASC "
objBuilder.Append "LIMIT "& stpage & "," &pgsize

Set rsSch = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = " 직원 학력 현황 "
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

				if(a == true){
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
				<form action="/insa/insa_school_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="전체" <%If view_condi = "전체" Then %>selected<%End If %>>전체</option>
                                  <option value="성명" <%If view_condi = "성명" Then %>selected<%End If %>>성명</option>
                                  <option value="sch_dept" <%If view_condi = "sch_dept" Then %>selected<%End If %>>학과</option>
                                  <option value="sch_major" <%If view_condi = "sch_major" Then %>selected<%End If %>>전공</option>
                                  <option value="sch_school_name" <%If view_condi = "sch_school_name" Then %>selected<%End If %>>학교</option>
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
							<col width="7%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="*" >
                            <col width="14%" >
                            <col width="12%" >
                            <col width="12%" >
                            <col width="8%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
                                <th scope="col">성명</th>
                                <th scope="col">직위</th>
								<th scope="col">회사</th>
								<th scope="col">소속</th>
                                <th scope="col">학교명</th>
								<th scope="col">기간</th>
								<th scope="col">학과</th>
								<th scope="col">전공</th>
								<th scope="col">부전공</th>
                                <th scope="col">학위</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsSch.EOF
							sch_empno = rsSch("sch_empno")
							emp_name = rsSch("emp_name")
							emp_grade = rsSch("emp_grade")
							emp_job = rsSch("emp_job")
							emp_position = rsSch("emp_position")
							emp_org_code = rsSch("emp_org_code")
							emp_org_name = rsSch("org_name")
							emp_company = rsSch("org_company")
	           			%>
							<tr>
								<td><%=rsSch("sch_empno")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsSch("sch_empno")%>','인사 기록 카드','scrollbars=yes,width=1250,height=670')"><%=emp_name%></a>
								</td>
                                <td><%=emp_job%>&nbsp;</td>
                                <td><%=emp_company%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rsSch("sch_school_name")%>&nbsp;</td>
                                <td><%=rsSch("sch_start_date")%>∼<%=rsSch("sch_end_date")%>&nbsp;</td>
                                <td><%=rsSch("sch_dept")%>&nbsp;</td>
                                <td><%=rsSch("sch_major")%>&nbsp;</td>
                                <td><%=rsSch("sch_sub_major")%>&nbsp;</td>
                                <td><%=rsSch("sch_degree")%>&nbsp;</td>
							</tr>
						<%
							rsSch.MoveNext()
						Loop
						rsSch.Close() : Set rsSch = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="/insa/insa_excel_schoollist.asp?view_condi=<%=view_condi%>&condi=<%=condi%>" class="btnType04">엑셀다운로드</a>
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
		</div>
	</div>
	</body>
</html>