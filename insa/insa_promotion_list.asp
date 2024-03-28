<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
On Error Resume Next

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
Dim curr_date, curr_year, curr_month, curr_day, be_pg
Dim page, to_date, in_grade, in_company, ck_sw
Dim condi_sql, pgsize, start_page, stpage, target_date
Dim total_record, rs_count, emp_grade_date
Dim year_cnt, mon_cnt, day_cnt, target_cnt, total_page
Dim title_line, page_cnt

curr_date = Mid(CStr(Now()), 1, 10)
curr_year = Mid(CStr(Now()), 1, 4)
curr_month = Mid(CStr(Now()), 6, 2)
curr_day = Mid(CStr(Now()), 9, 2)

be_pg = "/insa/insa_promotion_list.asp"

page = Request("page")
to_date = Request("to_date")
in_grade = Request("in_grade")
in_company = Request("in_company")
ck_sw = Request("ck_sw")

If ck_sw = "n" Then
	to_date = Request.Form("to_date")
    in_grade = Request.Form("in_grade")
	in_company = Request.Form("in_company")
Else
	to_date = Request("to_date")
    in_grade = Request("in_grade")
	in_company = Request("in_company")
End If

If in_company = "" Then
	'in_company = "케이원정보통신"
	in_company = "전체"
	to_date = curr_date
	in_grade = "대리2급"
End If

Select Case in_grade
	Case "대리2급"
		condi_sql = "AND emp_grade LIKE '%사원%' "
	Case "대리1급"
		condi_sql = "AND emp_grade LIKE '%대리2급%' "
	Case "과장"
		condi_sql = "AND (emp_grade LIKE '%대리2급%' OR emp_grade LIKE '%대리1급%') "
	Case "차장"
		condi_sql = "AND emp_grade LIKE '%과장%' "
	Case "부장"
		condi_sql = "AND emp_grade LIKE '%차장%' "
	Case Else
		condi_sql = ""
End Select

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

target_date = to_date
total_record = 0

objBuilder.Append "SELECT emtt.emp_grade_date, emtt.emp_first_date, emtt.emp_grade "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01') "
objBuilder.Append "	AND emtt.emp_no < '999990' "
If in_company <> "전체" Then
	objBuilder.Append "	AND eomt.org_company = '"&in_company&"' "
End If
objBuilder.Append condi_sql

Set rs_count = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rs_count.EOF
	If rs_count("emp_grade_date") = "1900-01-01" Then
		emp_grade_date = ""
	Else
		emp_grade_date = rs_count("emp_grade_date")
	End If

	If emp_grade_date <> "" Then
		year_cnt = DateDiff("yyyy", rs_count("emp_grade_date"), target_date)
		mon_cnt = DateDiff("m", rs_count("emp_grade_date"), target_date)
		day_cnt = DateDiff("d", rs_count("emp_grade_date"), target_date)
	Else
		year_cnt = DateDiff("yyyy", rs_count("emp_first_date"), target_date)
		mon_cnt = DateDiff("m", rs_count("emp_first_date"), target_date)
		day_cnt = DateDiff("d", rs_count("emp_first_date"), target_date)
	End If

	target_cnt = CInt(mon_cnt)

'   tottal_record = tottal_record + 1

	If (in_grade = "대리2급" Or in_grade = "대리1급") And target_cnt > 24 Then
		total_record = total_record + 1
    Else
		If in_grade = "과장" And rs_count("emp_grade") = "대리1급" And target_cnt > 36 Then
			total_record = total_record + 1
		Else
			If in_grade = "과장" And rs_count("emp_grade") = "대리2급" And target_cnt > 48 Then
		        total_record = total_record + 1
			End If
		End If
	End If

	rs_count.MoveNext()
Loop
rs_count.Close() : Set rs_count = Nothing

'tottal_record = cint(RsCount(0)) 'Result.RecordCount

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

title_line = " 승진대상자 현황 "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
				return "5 1";
			}

			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit();
				}
			}

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
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_asses_promo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="<%=be_pg%>?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>대상자 검색</dt>
                        <dd>
                            <p>
								<strong>회사 : </strong>
								<%Call SelectEmpOrgList("in_company", "in_company", "width:120px", in_company)%>
                                <strong>승진기준일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker">
                                <strong>승진직급 : </strong>
								<%Call SelectEmpEtcCodeList("in_grade", "in_grade", "width:70px;", "02", in_grade)%>
                                <span>&nbsp;※ 승진기준은 매년 1월 1일 기준입니다.</span>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
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
							<col width="8%" >
							<col width="6%" >
							<col width="12%" >
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
                                <th scope="col">생년월일</th>
								<th scope="col">현직급</th>
								<th scope="col">직책</th>
								<th scope="col">소속</th>
								<th scope="col">최초<br>입사일</th>
                                <th scope="col">입사일</th>
                                <th scope="col">최종<br>승진일</th>
                                <th scope="col">대상년한</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim rs_emp

						objBuilder.Append "SELECT emtt.emp_grade_date, emtt.emp_first_date, emtt.emp_no, emtt.emp_name, "
						objBuilder.Append "	emtt.emp_birthday, emtt.emp_grade, emtt.emp_position, "
						objBuilder.Append "	emtt.emp_in_date, "
						objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_team "
						objBuilder.Append "FROM emp_master AS emtt "
						objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
						objBuilder.Append "WHERE (ISNULL(emp_end_date) OR emp_end_date = '1900-01-01') "
						objBuilder.Append "	AND emtt.emp_no < '999990' "
						If in_company <> "전체" Then
							objBuilder.Append "	AND eomt.org_company = '"&in_company&"' "
						End If
						objBuilder.Append condi_sql
						objBuilder.Append "ORDER BY emp_first_date, emp_no DESC "
						objBuilder.Append "LIMIT "& stpage & "," &pgsize

						Set rs_emp = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						Do Until rs_emp.EOF
							If rs_emp("emp_grade_date") = "1900-01-01" Then
							   emp_grade_date = ""
							Else
							   emp_grade_date = rs_emp("emp_grade_date")
							End If

							If emp_grade_date <> "" Then
								year_cnt = DateDiff("yyyy", rs_emp("emp_grade_date"), target_date)
								mon_cnt = DateDiff("m", rs_emp("emp_grade_date"), target_date)
								day_cnt = DateDiff("d", rs_emp("emp_grade_date"), target_date)
							Else
							    year_cnt = DateDiff("yyyy", rs_emp("emp_first_date"), target_date)
								mon_cnt = DateDiff("m", rs_emp("emp_first_date"), target_date)
								day_cnt = DateDiff("d", rs_emp("emp_first_date"), target_date)
							End If

							target_cnt = CInt(mon_cnt)

							If (in_grade = "대리2급" Or in_grade = "대리1급") And target_cnt > 24 Then
	           			%>
							<tr>
								<td class="first"><%=rs_emp("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rs_emp("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs_emp("emp_name")%></a>
								</td>
                                <td><%=rs_emp("emp_birthday")%>&nbsp;</td>
                                <td><%=rs_emp("emp_grade")%>&nbsp;</td>
                                <td><%=rs_emp("emp_position")%>&nbsp;</td>
                                <td><%=rs_emp("org_name")%>&nbsp;</td>
                                <td><%=rs_emp("emp_first_date")%>&nbsp;</td>
                                <td><%=rs_emp("emp_in_date")%>&nbsp;</td>
                                <td><%=emp_grade_date%>&nbsp;</td>
                                <td><%=mon_cnt%>&nbsp;개월</td>
                                <td class="left">
									<%Call EmpOrgText(rs_emp("org_company"), rs_emp("org_bonbu"), rs_emp("org_team"))%>
								</td>
							</tr>
						<%
						      Else
								If in_grade = "과장" And rs_emp("emp_grade") = "대리1급" And target_cnt > 36 Then
	           			%>
							<tr>
								<td class="first"><%=rs_emp("emp_no")%>&nbsp;<td>
                                <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rs_emp("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs_emp("emp_name")%></a>
								</td>
                                <td><%=rs_emp("emp_birthday")%>&nbsp;</td>
                                <td><%=rs_emp("emp_grade")%>&nbsp;</td>
                                <td><%=rs_emp("emp_position")%>&nbsp;</td>
                                <td><%=rs_emp("org_name")%>&nbsp;</td>
                                <td><%=rs_emp("emp_first_date")%>&nbsp;</td>
                                <td><%=rs_emp("emp_in_date")%>&nbsp;</td>
                                <td><%=emp_grade_date%>&nbsp;</td>
                                <td><%=mon_cnt%>&nbsp;개월</td>
                                <td class="left">
									<%Call EmpOrgText(rs_emp("org_company"), rs_emp("org_bonbu"), rs_emp("org_team"))%>
								</td>
							</tr>
						<%
						      Else
								If in_grade = "과장" And rs_emp("emp_grade") = "대리2급" And target_cnt > 48 Then
	           			%>
							<tr>
								<td class="first"><%=rs_emp("emp_no")%>&nbsp;<td>
                                <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rs_emp("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs_emp("emp_name")%></a>
								</td>
                                <td><%=rs_emp("emp_birthday")%>&nbsp;</td>
                                <td><%=rs_emp("emp_grade")%>&nbsp;</td>
                                <td><%=rs_emp("emp_position")%>&nbsp;</td>
                                <td><%=rs_emp("org_name")%>&nbsp;</td>
                                <td><%=rs_emp("emp_first_date")%>&nbsp;</td>
                                <td><%=rs_emp("emp_in_date")%>&nbsp;</td>
                                <td><%=emp_grade_date%>&nbsp;</td>
                                <td><%=mon_cnt%>&nbsp;개월</td>
                                <td class="left">
									<%Call EmpOrgText(rs_emp("org_company"), rs_emp("org_bonbu"), rs_emp("org_team"))%>
								</td>
							</tr>
						<%
								End If
							End If
						End If
							rs_emp.MoveNext()
						Loop
						rs_emp.Close() : Set rs_emp = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page, i

                intstart = (Int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                If intend > total_page Then
                    intend = total_page
                End If
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="/insa/excel/insa_excel_promotlist.asp?in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "<%=be_pg%>?page=<%=first_page%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% If intstart > 1 Then %>
                        <a href="<%=be_pg%>?page=<%=intstart -1%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                    <% End If %>
                    <% For i = intstart To intend %>
           				<% If i = Int(page) Then %>
                        <b>[<%=i%>]</b>
						<% Else %>
                        <a href="<%=be_pg%>?page=<%=i%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
						<% End If %>
                    <% Next %>
           				<% If intend < total_page Then %>
                        <a href="<%=be_pg%>?page=<%=intend+1%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="<%=be_pg%>?page=<%=total_page%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	Else %>
                        [다음]&nbsp;[마지막]
                      <% End If %>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>