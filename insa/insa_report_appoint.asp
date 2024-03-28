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
Dim be_pg, from_date, to_date, Page, view_condi
Dim ck_sw, app_id, curr_dd, pgsize, start_page
Dim stpage, RsCount, tottal_record, total_page
Dim rsApppoint, title_line

be_pg = "/insa/insa_report_appoint.asp"

from_date = Request.Form("from_date")
to_date = Request.Form("to_date")

Page = Request("page")
view_condi = Request("view_condi")

ck_sw = Request("ck_sw")

If ck_sw = "n" Then
	view_condi = Request.Form("view_condi")
	app_id = Request.Form("app_id")
	from_date = Request.Form("from_date")
    to_date = Request.Form("to_date")
Else
	view_condi = Request("view_condi")
	app_id = Request("app_id")
	from_date = Request("from_date")
    to_date = Request("to_date")
End If

If view_condi = "" Then
	view_condi = "전체"
	app_id = "전체"
	curr_dd = CStr(DatePart("d", Now()))
	to_date = Mid(CStr(Now()), 1, 10)
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
End If

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM emp_appoint "
objBuilder.Append "WHERE (app_date >= '"&from_date&"' AND app_date <= '"&to_date&"') "
objBuilder.Append "	AND app_empno < '900000' "

If view_condi <> "전체" Then
	'Sql = "SELECT count(*) from emp_appoint where app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000')"
'Else
	'Sql = "select count(*) from emp_appoint where app_to_company='"+view_condi+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000')"
	objBuilder.Append "	AND app_to_company = '"&view_condi&"' "
End If

Set RsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

tottal_record = CInt(RsCount(0))	'Result.RecordCount

If tottal_record Mod pgsize = 0 Then
	'Result.PageCount
	total_page = Int(tottal_record / pgsize)
Else
	total_page = Int((tottal_record / pgsize) + 1)
End If

objBuilder.Append "SELECT app_empno, app_emp_name, app_date, app_id, app_id_type, "
objBuilder.Append "	app_to_company, app_to_org, app_to_orgcode, app_to_grade, "
objBuilder.Append "	app_to_position, app_be_company, app_be_org, app_be_orgcode, "
objBuilder.Append "	app_be_grade, app_be_position, app_start_date, app_finish_date, "
objBuilder.Append "	app_be_enddate, app_reward, app_comment "
objBuilder.Append "FROM emp_appoint "
objBuilder.Append "WHERE (app_date >= '"&from_date&"' AND app_date <= '"&to_date&"') "
objBuilder.Append "	AND app_empno < '900000' "

If view_condi = "전체" Then
	If app_id <> "전체" Then
		'Sql = "select * from emp_appoint where app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC limit "& stpage & "," &pgsize

	'Else
		'Sql = "select * from emp_appoint where app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC limit "& stpage & "," &pgsize
		objBuilder.Append "	AND app_id = '"&app_id&"' "
	End If
 Else
 	objBuilder.Append "	AND app_to_company = '"&view_condi&"' "
	If app_id <> "전체" Then
		'Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC limit "& stpage & "," &pgsize
	'Else
		'Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC limit "& stpage & "," &pgsize
		objBuilder.Append "	AND app_id = '"&app_id&"' "
	End If
End If

objBuilder.Append "ORDER BY app_date,app_empno ASC "
objBuilder.Append "LIMIT "& stpage & "," &pgsize

'Set rsApppoint = Server.CreateObject("ADODB.RecordSet")
'rsApppoint.Open objBuilder.ToString(), DBConn, 1
Set rsApppoint = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = view_condi &" - 인사발령 현황(" & from_date & " ∼ " & to_date & ")"
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
				return "2 1";
			}

			$(function(){
				$("#datepicker").datepicker();
				$("#datepicker").datepicker("option", "dateFormat", "yy-mm-dd" );
				$("#datepicker").datepicker("setDate", "<%=from_date%>" );
			});

			$(function(){
				$("#datepicker1").datepicker();
				$("#datepicker1").datepicker("option", "dateFormat", "yy-mm-dd" );
				$("#datepicker1").datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit ();
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
	<body oncontextmenu="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_appoint_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/insa_report_appoint.asp?ck_sw=n" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
							  	Dim rs_org, rs_etc

								objBuilder.Append "SELECT org_name "
								objBuilder.Append "FROM emp_org_mst "
								objBuilder.Append "WHERE (isNull(org_end_date) OR org_end_date = '0000-00-00') "
								objBuilder.Append "	AND org_level = '회사' AND org_code <> '6272' "
								objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

								Set rs_org = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px;">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<%End If %>>전체</option>
                			  <%
								Do Until rs_org.EOF
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.MoveNext()
								Loop
								rs_org.Close() : Set rs_org = Nothing
							  %>
            					</select>
                                </label>
                                <label>
                                <strong>발령구분</strong>
                            <%
								'Sql="select * from emp_etc_code where emp_etc_type = '10' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name "
								objBuilder.Append "FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '10' "
								objBuilder.Append "ORDER BY emp_etc_code ASC "

								'rs_etc.Open objBuilder.ToString(), DBConn, 1
								Set rs_etc = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							%>
								<select name="app_id" id="select" type="text" style="width:150px">
                                    <option value="전체" <%If app_id = "전체" then %>selected<% end if %>>전체</option>
                			<%
								Do Until rs_etc.EOF
			  				%>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If app_id = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%>&nbsp;</option>
                			<%
									rs_etc.MoveNext()
								Loop
								rs_etc.Close() : Set rs_etc = Nothing
							%>
            					</select>
								</label>
								<label>
								<strong>발령일(From) : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong> ∼ To : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
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
							<col width="9%" >
							<col width="10%" >
							<col width="9%" >
							<col width="9%" >
							<col width="10%" >
                            <col width="9%" >
                            <col width="*" >
						</colgroup>
						<thead>
                            <tr>
				                <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">사번</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">성명</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령일</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령구분</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령유형</th>
                                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령전</th>
				                <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령후</th>
			                </tr>
                            <tr>
                                <th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">직급/책</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">직급/책</th>
                                <th scope="col">발령내용</th>
                            </tr>
						</thead>
						<tbody>
						<%
					  	   Do Until rsApppoint.EOF

	           			%>
							<tr>
								<td><%=rsApppoint("app_empno")%>&nbsp;</td>
                                <td><%=rsApppoint("app_emp_name")%>&nbsp;</td>
                                <td><%=rsApppoint("app_date")%>&nbsp;</td>
								<td><%=rsApppoint("app_id")%>&nbsp;</td>
                                <td><%=rsApppoint("app_id_type")%>&nbsp;</td>
                                <td><%=rsApppoint("app_to_company")%>&nbsp;</td>
                                <td><%=rsApppoint("app_to_org")%>(<%=rsApppoint("app_to_orgcode")%>)&nbsp;</td>
                                <td><%=rsApppoint("app_to_grade")%>-<%=rsApppoint("app_to_position")%>&nbsp;</td>
                                <td><%=rsApppoint("app_be_company")%>&nbsp;</td>
                                <td><%=rsApppoint("app_be_org")%>(<%=rsApppoint("app_be_orgcode")%>)&nbsp;</td>
                                <td><%=rsApppoint("app_be_grade")%>-<%=rsApppoint("app_be_position")%>&nbsp;</td>
                                <td class="left">
									<%=rsApppoint("app_start_date")%>&nbsp;-&nbsp;<%=rsApppoint("app_finish_date")%>&nbsp;
									<%=rsApppoint("app_be_enddate")%>&nbsp;
									<%=rsApppoint("app_reward")%>&nbsp;:&nbsp;<%=rsApppoint("app_comment")%>&nbsp;
								</td>
							</tr>
						<%
							  rsApppoint.MoveNext()
						  Loop
						  rsApppoint.Close() : Set rsApppoint = Nothing
						  DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page, i

                intstart = (Int((page - 1) / 10) * 10) + 1
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
						<a href="/insa/insa_excel_appoint.asp?view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
					<div id="paging">
						<a href = "<%=be_pg%>?page=<%=first_page%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% If intstart > 1 Then %>
                        <a href="<%=be_pg%>?page=<%=intstart -1%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
					<% End If %>

					<% For i = intstart To intend %>
           				<% If i = Int(page) Then %>
                        <b>[<%=i%>]</b>
						<% Else %>
                        <a href="<%=be_pg%>?page=<%=i%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
						<% End If %>
					<% Next %>
					<% If intend < total_page Then %>
                        <a href="<%=be_pg%>?page=<%=intend+1%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="<%=be_pg%>?page=<%=total_page%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                    <% Else %>
                        [다음]&nbsp;[마지막]
					<% End If %>
                    </div>
                    </td>
                    <td>
				    <td width="15%">
					<div class="btnCenter">
			            <a href="#" onClick="pop_Window('/insa/insa_appoint_print.asp?view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>','pop_report','scrollbars=yes,width=1250,height=600')" class="btnType04">출력</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>