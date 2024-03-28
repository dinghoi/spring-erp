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
Dim be_pg, from_date, to_date, page, view_condi, ck_sw
Dim curr_dd, pgsize, start_page, stpage, where_sql, rs_emp
Dim page_cnt, title_line
Dim rs_count, total_record, total_page, pg_url
Dim emp_end_date, target_date, first_date, emp_org_baldate, emp_grade_date
Dim year_cnt, mon_cnt, day_cnt, y_cnt, m_cnt, app_empno, rs_app
Dim app_id_type, app_comment, app_task, task_memo, view_memo

be_pg = "/insa/insa_report_emp_out.asp"

from_date = f_Request("from_date")
to_date = f_Request("to_date")

page = f_Request("page")
view_condi = f_Request("view_condi")

If view_condi = "" Then
	view_condi = "전체"
	curr_dd = CStr(DatePart("d", Now()))
	to_date = Mid(CStr(Now()), 1, 10)
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
End If

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&from_date="&from_date&"&to_date="&to_date

If view_condi <> "전체" Then
	where_sql = "AND eomt.org_company = '"&view_condi&"' "
Else
	where_sql = ""
End If

objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (emtt.emp_end_date >= '"&from_date&"' AND emtt.emp_end_date <= '"&to_date&"') "
objBuilder.Append where_sql

Set rs_count = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rs_count(0)) 'Result.RecordCount

rs_count.Close() : Set rs_count = Nothing

objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_end_date, emtt.emp_first_date, "
objBuilder.Append "	emtt.emp_org_baldate, emtt.emp_grade_date, emtt.emp_birthday, "
objBuilder.Append "	emtt.emp_grade, emtt.emp_position, emtt.emp_in_date, emtt.emp_end_date, "
objBuilder.Append "	emtt.emp_last_edu, emtt.emp_disabled, emtt.emp_disab_grade, "
objBuilder.Append "	emtt.emp_org_name, eomt.org_code, eomt.org_name "
'objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_team "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emtt.emp_end_date >= '" & from_date & "' AND emtt.emp_end_date <= '" & to_date & "' "
objBuilder.Append where_sql
objBuilder.Append "ORDER BY emtt.emp_no, emtt.emp_name ASC "
objBuilder.Append "LIMIT "& stpage & "," &pgsize

Set rs_emp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = view_condi & " - 퇴직자 현황(" & from_date & " ∼ " & to_date & ")"
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

			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit ();
				}
			}
			/*
			function delcheck(){
				if(form_chk(document.frm_del)){
					document.frm_del.submit ();
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
				<form action="/insa/insa_report_emp_out.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%

							  'Call SelectEmpOrgList("view_condi", "view_condi", "width:150px", view_condi)
							  %>
            					<%
							   Dim rs_org
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								'objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC;"
								objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
								objBuilder.Append "	AND org_level = '회사' AND org_code <> '6272' "
								objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

								Set rs_org = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
								<select name="view_condi" id="view_condi" type="text" style="width:110px;">
									<option value="전체">전체</option>
									<%
									Do Until rs_org.EOF
										%>
										<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") Then %>selected<%End If %>><%=rs_org("org_name")%></option>
										<%
										rs_org.MoveNext()
									Loop
									rs_org.Close() : Set rs_org = Nothing
									%>
								</select>
                                </label>
								<label>
								<strong>퇴사일(From) : </strong>
                                	<input type="text" name="from_date" value="<%=from_date%>" style="width:70px;" id="datepicker"/>
								</label>
								<label>
								<strong> ∼ To : </strong>
                                	<input type="text" name="to_date" value="<%=to_date%>" style="width:70px;" id="datepicker1"/>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="7%" >
							<col width="6%" >
							<col width="*" >
                            <col width="7%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">생년월일</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
                                <th scope="col">퇴직일</th>
                                <th scope="col">근무<br>기간</th>
                                <th scope="col">소속</th>
                                <th scope="col">최종학력</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <th scope="col">장애여부</th>
								<th scope="col">퇴직사유</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rs_emp.EOF
							emp_end_date = rs_emp("emp_end_date")
							target_date = rs_emp("emp_end_date")
							first_date = rs_emp("emp_first_date")

							If rs_emp("emp_org_baldate") = "1900-01-01" Then
							   emp_org_baldate = ""
							Else
							   emp_org_baldate = rs_emp("emp_org_baldate")
							End If

							If rs_emp("emp_grade_date") = "1900-01-01" Then
							   emp_grade_date = ""
							Else
							   emp_grade_date = rs_emp("emp_grade_date")
							End If

							year_cnt = DateDiff("yyyy", first_date, target_date)
							mon_cnt = DateDiff("m", first_date, target_date)
							day_cnt = DateDiff("d", first_date, target_date)

							year_cnt = Int(year_cnt) + 1
							mon_cnt = Int(mon_cnt) + 1
							day_cnt = Int(day_cnt) + 1
							y_cnt = Int(mon_cnt / 12)
							m_cnt = mon_cnt - (y_cnt * 12)

							app_empno = rs_emp("emp_no")

							objBuilder.Append "SELECT app_id_type, app_comment "
							objBuilder.Append "FROM emp_appoint "
							objBuilder.Append "WHERE app_empno = '"&app_empno&"' "
							objBuilder.Append "	AND app_id = '퇴직발령' "
							objBuilder.Append "	AND app_be_enddate = '"&emp_end_date&"' "

							Set rs_app = DbConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							If Not rs_app.EOF Then
							  app_id_type = rs_app("app_id_type")
							  app_comment = rs_app("app_comment")
							Else
							  app_id_type = ""
							  app_comment = ""
							End If
							rs_app.Close()

							app_task = app_id_type & " - " & app_comment
							task_memo = replace(app_task, Chr(34), Chr(39))
							view_memo = task_memo

							If Len(task_memo) > 10 Then
								view_memo = Mid(task_memo, 1, 10) & ".."
							End If
	           			%>
							<tr>
								<td class="first"><%=rs_emp("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rs_emp("emp_no")%>','인사 기록 카드','scrollbars=yes,width=1250,height=670')"><%=rs_emp("emp_name")%></a>
								</td>
                                <td><%=rs_emp("emp_birthday")%>&nbsp;</td>
                                <td><%=rs_emp("emp_grade")%>&nbsp;</td>
                                <td><%=rs_emp("emp_position")%>&nbsp;</td>
                                <td><%=rs_emp("emp_in_date")%>&nbsp;</td>
                                <td><%=rs_emp("emp_end_date")%>&nbsp;</td>
								<% If y_cnt > 0 And m_cnt > 0 Then %>
                                <td><%=y_cnt%>년&nbsp;<%=m_cnt%>개월</td>
								 <% End If %>
								<% If y_cnt > 0 And m_cnt = 0 Then %>
                                <td><%=y_cnt%>년&nbsp;</td>
								<% End If %>
								<% If y_cnt = 0 And m_cnt > 0 Then %>
                                <td><%=m_cnt%>개월&nbsp;</td>
								<% End If %>
								<% If y_cnt = 0 And m_cnt = 0 Then %>
                                <td><%=m_cnt%>개월&nbsp;</td>
								<% End If %>
                                <td><%=rs_emp("org_name")%>&nbsp;</td>
                                <td><%=rs_emp("emp_last_edu")%>&nbsp;</td>
                                <td class="left">
								<%
								Call EmpOrgCodeSelect(rs_emp("org_code"))
								%>
								</td>
                                <td><%=rs_emp("emp_disabled")%>&nbsp;<%=rs_emp("emp_disab_grade")%>&nbsp;</td>
                                <td class="left">
									<p style="cursor:pointer">
										<span title="<%=task_memo%>"><%=view_memo%></span>
									</p>
								</td>
							</tr>
						<%
							rs_emp.MoveNext()
						Loop
						Set rs_app = Nothing
						rs_emp.close() : Set rs_emp = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="/insa/insa_excel_emp_out.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
                    <td>
				    <td width="15%">
					<div class="btnCenter">
			            <a href="#" onClick="pop_Window('/insa/insa_emp_out_print.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>','pop_report','scrollbars=yes,width=1050,height=500')" class="btnType04">출력</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>

