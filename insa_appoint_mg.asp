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
Dim be_pg, curr_date
Dim view_condi, owner_view
Dim ck_sw, title_line

Dim rs

Dim at_name, at_empno, at_position

be_pg = "insa_appoint_mg.asp"
curr_date = DateValue(Mid(CStr(Now()), 1,10))

view_condi = Request("view_condi" )
owner_view = Request("owner_view")

ck_sw = Request("ck_sw")

If ck_sw = "n" Then
	owner_view=Request.form("owner_view")
	view_condi = request.form("view_condi")
Else
	owner_view=request("owner_view")
	view_condi = request("view_condi")
End If

If view_condi = "" Then
	view_condi = ""
	owner_view = "C"
	ck_sw = "n"
End If

title_line = " 인사발령 처리  "
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
				return "2 1";
			}
			function goAction () {
			   window.close () ;
			}
		</script>

		<script type="text/javascript">
			/*$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%'=from_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%'=to_date%>" );
			});
			*/
			function frmcheck(){
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if (document.frm.view_condi.value == "") {
					alert ("조건을 입력하시기 바랍니다");
					return false;
				}
				return true;
			}
		</script>

	</head>

	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_appoint_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_appoint_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈조건 검색◈</dt>
                        <dd>
                            <p>
                                <label>
                                <input name="owner_view" type="radio" value="T" <%If owner_view = "T" Then %>checked<%End If %> style="width:25px">사번
                                <input name="owner_view" type="radio" value="C" <%If owner_view = "C" Then %>checked<%End If %> style="width:25px">성명
                                </label>
							<strong>조건 : </strong>
								<label>
        						<input name="view_condi" type="text" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left; ime-mode:active">
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
							<col width="9%" >
                            <col width="6%" >
							<col width="*" >
                            <col width="4%" >
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
							   <th scope="col">상주처</th>
                               <th scope="col">생년월일</th>
							   <th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                               <th>처리</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						Dim emp_org_baldate, emp_grade_date, emp_type
						Dim page, view_sort, date_sw, page_cnt

						If view_condi <> "" Then
							objBuilder.Append "SELECT emp_org_baldate, emp_grade_date, emp_type, emp_no, emp_name, "
							objBuilder.Append "emp_grade, emp_job, emp_position, emp_in_date, emp_org_name, "
							objBuilder.Append "emp_first_date, emp_reside_place, emp_birthday, emp_company, emp_bonbu, "
							objBuilder.Append "emp_saupbu, emp_team "
							objBuilder.Append "FROM emp_master "
							'objBuilder.Append "WHERE (isNull(emp_end_date) OR emp_end_date = '1900-01-01') AND (emp_no < '900000') "
							objBuilder.Append "WHERE (isNull(emp_end_date) OR emp_end_date = '1900-01-01') "

							If owner_view = "C" Then
								'sql = "select * from emp_master where emp_name like '%"+view_condi+"%' and (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_no,emp_name ASC"
								objBuilder.Append "AND emp_name LIKE '%"&view_condi&"%' "
							Else
								'sql = "select * from emp_master where emp_no = '"+view_condi+"' and (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_no,emp_name ASC"
								objBuilder.Append "AND emp_no = '"&view_condi&"' "
							End If

							objBuilder.Append "ORDER BY emp_no, emp_name ASC "

							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open objBuilder.ToString(), DBConn, 1
							objBuilder.Clear()

							Do Until rs.EOF

								If rs("emp_org_baldate") = "1900-01-01" Then
								   emp_org_baldate = ""
								Else
								   emp_org_baldate = rs("emp_org_baldate")
								End If

								If rs("emp_grade_date") = "1900-01-01" Then
								   emp_grade_date = ""
								Else
								   emp_grade_date = rs("emp_grade_date")
								End If

								emp_type = rs("emp_type")
						%>
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rs("emp_reside_place")%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
							    <td>
                                <a href="insa_appoint_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&be_pg=<%=be_pg%>&u_type=<%="U"%>">발령</a>
                                </td>
							</tr>
						<%
								rs.MoveNext()
							Loop
							rs.Close() : Set rs = Nothing

						End If

						DBConn.Close : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>

