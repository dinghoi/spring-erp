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
Dim be_pg, from_date, to_date, page, view_condi, ck_sw
Dim curr_dd, pgsize, start_page, stpage
Dim rsCount, total_record, total_page
Dim rsEmp, title_line, where_sql
Dim emp_org_baldate, emp_grade_date, pg_url

be_pg = "/insa/insa_report_emp_in.asp"

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
	where_sql = "	AND eomt.org_company='"&view_condi&"' "
Else
	where_sql = ""
End If

objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (emtt.emp_in_date >= '" & from_date & "' AND emtt.emp_in_date <= '" & to_date & "') "
objBuilder.Append "	AND emtt.emp_no < '900000' "
objBuilder.Append where_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT emtt.emp_org_baldate, emtt.emp_grade_date, emtt.emp_no, emtt.emp_name, "
objBuilder.Append "	emtt.emp_birthday, emtt.emp_grade, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_in_date, emtt.emp_last_edu, emtt.emp_disabled, "
objBuilder.Append "	emtt.emp_disab_grade, emtt.emp_reside_company, eomt.org_name, eomt.org_code "
'objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (emtt.emp_in_date >= '" & from_date & "' AND emtt.emp_in_date <= '" & to_date & "') "
objBuilder.Append "	AND emtt.emp_no < '900000' "
objBuilder.Append where_sql
objBuilder.Append "ORDER BY emtt.emp_no, emtt.emp_name ASC "
objBuilder.Append "LIMIT "& stpage & "," & pgsize

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "" & view_condi & " - 입사자 현황(" & from_date & " ∼ " & to_date & ")"
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
					document.frm.submit();
				}
			}
			/*
			function delcheck(){
				if (form_chk(document.frm_del)){
					document.frm_del.submit();
				}
			}

			function form_chk(){
				a=confirm('삭제하시겠습니까?');

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
				<form action="/insa/insa_report_emp_in.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
							   <label>
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
								<strong>입사일(From) : </strong>
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
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="8%" >
							<col width="9%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">생년월일</th>
								<th scope="col">직급</th>
								<th scope="col">직위</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
                                <th scope="col">소속</th>
                                <th scope="col">최종학력</th>
								<th scope="col">장애여부</th>
								<th scope="col">상주처회사</th>
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
                                <td><%=rsEmp("emp_birthday")%>&nbsp;</td>
                                <td><%=rsEmp("emp_grade")%>&nbsp;</td>
                                <td><%=rsEmp("emp_job")%>&nbsp;</td>
                                <td><%=rsEmp("emp_position")%>&nbsp;</td>
                                <td><%=rsEmp("emp_in_date")%>&nbsp;</td>
                                <td><%=rsEmp("org_name")%>&nbsp;</td>
                                <td><%=rsEmp("emp_last_edu")%>&nbsp;</td>
                                <td><%=rsEmp("emp_disabled")%>&nbsp;<%=rsEmp("emp_disab_grade")%>&nbsp;</td>
                                <td><%=rsEmp("emp_reside_company")%>&nbsp;</td>
                                <td class="left">
								<%
								Call EmpOrgCodeSelect(rsEmp("org_code"))
								%>
								</td>
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
                  	<td width="15%">
					<div class="btnCenter">
						<a href="/insa/insa_excel_emp_in.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
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
						<a href="#" onClick="pop_Window('/insa/insa_emp_in_print.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>','pop_report','scrollbars=yes,width=1050,height=500')" class="btnType04">출력</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>