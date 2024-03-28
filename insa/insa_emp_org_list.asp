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
Dim Page, be_pg, curr_date, page_cnt, pg_cnt, ck_sw
Dim view_condi, pgsize, start_page, stpage, view_sort
Dim order_Sql, where_sql, rsCount, total_record, total_page
Dim Rs, rs_org
Dim title_line

Page = Request("page")
page_cnt = Request.Form("page_cnt")
pg_cnt = cint(Request("pg_cnt"))
be_pg = "./insa_emp_org_list.asp"
curr_date = DateValue(Mid(CStr(Now()), 1, 10))

ck_sw = Request("ck_sw")

If ck_sw = "y" Then
	view_condi = Request("view_condi")
Else
	view_condi = Request.Form("view_condi")
End if

If view_condi = "" Then
	view_condi = "전체"
End If

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

view_sort = Request("view_sort")

If view_sort = "" Then
	view_sort = "ASC "
End If

order_Sql = "ORDER BY eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name, emtt.emp_no, emtt.emp_in_date "&view_sort
where_sql = "WHERE (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "

If view_condi = "전체" Then
	where_sql = where_sql & "AND emtt.emp_no < '900000' "
Else
	where_sql = where_sql & "AND eomt.org_company = '"&view_condi&"' AND emtt.emp_no < '900000' "
End If

objBuilder.Append "SELECT COUNT(*) FROM emp_org_mst AS eomt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON eomt.org_code = emtt.emp_org_code "
objBuilder.Append where_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(RsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

title_line = " 조직전체 직원 현황 "
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
				return "1 1";
			}

			function frmcheck(){
				if (formcheck(document.frm) && chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if (document.frm.view_condi.value == ""){
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
			<!--include virtual = "/include/insa_asses_promo_menu.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_org_list.asp" method="post" name="frm">

				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>회사 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								'Sql="select * from emp_org_mst where (org_level = '회사') ORDER BY org_code ASC"
								objBuilder.Append "SELECT org_name FROM emp_org_mst "
								objBuilder.Append "WHERE (org_level = '회사') "
								objBuilder.Append "AND (org_end_date IS NULL OR org_end_date = '0000-00-00') "
								objBuilder.Append "ORDER BY FIELD(org_company, '케이원') DESC,"
								objBuilder.Append "org_code DESC "

								Set rs_org = Server.CreateObject("ADODB.Recordset")
	                            rs_org.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
                                    <option value="전체" <%If view_condi = "0" Then %>selected<%End If%>>전체</option>

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
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>

				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
                            <col width="5%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
                            <col width="6%" >
							<col width="*" >
                            <col width="12%" >
                            <col width="16%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>

								<th scope="col">직위</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
                                <th scope="col">소속</th>
								<th scope="col">소속발령일</th>

                                <th scope="col">생년월일</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <th scope="col">상주회사</th>
                                <th scope="col">실근무지</th>
							</tr>
						</thead>
					<tbody>
						<%
						Dim date_sw, emp_org_baldate

						objBuilder.Append "SELECT eomt.org_code, eomt.org_level, eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name, "
						objBuilder.Append "	eomt.org_reside_place, eomt.org_reside_company, "
						objBuilder.Append "	emtt.emp_no, emtt.emp_name, emtt.emp_job, emtt.emp_position, emtt.emp_in_date, "
						objBuilder.Append "	emtt.emp_org_baldate, emtt.emp_birthday, emtt.emp_stay_name "
						objBuilder.Append "FROM emp_org_mst AS eomt "
						objBuilder.Append "INNER JOIN emp_master AS emtt ON eomt.org_code = emtt.emp_org_code "
						objBuilder.Append where_sql & order_sql
						objBuilder.Append "LIMIT "& stpage & "," &pgsize

						Set Rs = Server.CreateObject("ADODB.Recordset")
						Rs.Open objBuilder.ToString(), DBConn, 1
						objBuilder.Clear()

						Do Until rs.EOF

							If rs("emp_org_baldate") = "1900-01-01" Then
							   emp_org_baldate = ""
							Else
							   emp_org_baldate = rs("emp_org_baldate")
							End If
						%>
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('../insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("org_name")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td class="left"><%=rs("org_company")%> > <%=rs("org_bonbu")%> > <%=rs("org_team")%></td>
                                <td class="left"><%=rs("org_reside_company")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_stay_name")%>&nbsp;</td>
							</tr>
						<%
							rs.MoveNext()
						Loop
						rs.Close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page
				Dim i

                intstart = (Int((page - 1)/10) * 10) + 1
                intend = intstart + 9
                first_page = 1

                If intend > total_page Then
                    intend = total_page
                End If
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="./insa_emp_org_list_excel.asp?view_condi=<%=view_condi%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_emp_org_list.asp?page=<%=first_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<%If intstart > 1 Then %>
                        <a href="insa_emp_org_list.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[이전]</a>
                    <%End If %>

                    <%For i = intstart To intend %>
                  		<%If i = Int(page) Then %>
                        <b>[<%=i%>]</b>
						<%Else %>
                        <a href="insa_emp_org_list.asp?page=<%=i%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
						<%End If %>
                    <%Next %>

                  	<%If intend < total_page Then %>
                        <a href="insa_emp_org_list.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_emp_org_list.asp?page=<%=total_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[마지막]</a>
                    <%Else %>
                        [다음]&nbsp;[마지막]
                    <%End If %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">

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

