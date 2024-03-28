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
Dim Repeat_Rows, page_cnt, pg_cnt
Dim Page, be_pg, curr_date, ck_sw, view_condi
Dim field_check, field_bonbu, field_saupbu, field_team
Dim view_c, pgsize, start_page, stpage
Dim rs, rs_org, rsCount
Dim order_Sql, owner_sql
Dim total_record, total_page, title_line

Page = Request("page")
page_cnt = Request.form("page_cnt")
pg_cnt = CInt(Request("pg_cnt"))
be_pg = "insa_org_mg.asp"
curr_date = DateValue(Mid(CStr(Now()), 1, 10))

ck_sw = Request("ck_sw")

If ck_sw = "y" Then
	view_condi = Request("view_condi")
	field_check = Request("field_check")
	field_bonbu = Request("field_bonbu")
	field_saupbu = Request("field_saupbu")
	field_team = Request("field_team")
	view_c = Request("view_c")
  else
	view_condi = Request.form("view_condi")
	field_check = Request.form("field_check")
	field_bonbu = Request.form("field_bonbu")
	field_saupbu = Request.form("field_saupbu")
	field_team = Request.form("field_team")
	view_c = Request.form("view_c")
End if

If view_condi = "" Then
	view_condi = "케이원정보통신"
	'view_condi = "케이시스템"
End If

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

order_Sql = " ORDER BY org_code, org_company,org_bonbu,org_saupbu,org_team,org_reside_place ASC"

If view_c = "" Then
	ck_sw = "n"
	field_check = "total"
	view_c = "bonbu"
End If

owner_sql = "WHERE (isNull(org_end_date) OR org_end_date = '1900-01-01' or org_end_date = '000-00-00') "

If field_check = "total" Then
	'owner_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '000-00-00') and (org_company = '"&view_condi&"')"
	owner_sql = owner_sql & "AND org_company = '"&view_condi&"' "
	field_check = ""
Else
	If view_c = "bonbu" Then
		'owner_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '000-00-00') and (org_company = '"&view_condi&"') and (org_bonbu like '%" + field_bonbu + "%')"
		owner_sql = owner_sql & "AND org_company = '"&view_condi&"' AND org_bonbu LIKE '%" + field_bonbu + "%' "
	End If

	If view_c = "saupbu" Then
		owner_sql = owner_sql & "AND org_company = '"&view_condi&"' AND org_saupbu LIKE '%" + field_saupbu + "%' "
	End If

	If view_c = "team" Then
		'owner_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '000-00-00') and (org_company = '"&view_condi&"') and (org_team like '%" + field_team + "%')"
		owner_sql = owner_sql & "AND org_company = '"&view_condi&"' AND org_team LIKE '%" + field_team + "%' "
	End If
End If

'Sql = "SELECT count(*) FROM emp_org_mst " + owner_sql
objBuilder.Append "SELECT COUNT(*) FROM emp_org_mst "
objBuilder.Append owner_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount
rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

title_line = " 조직 현황 "
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
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck(){
				//if (formcheck(document.frm) && chkfrm()) {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if (document.frm.view_condi.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}

			function condi_view(){

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('bonbu1').style.display = '';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = 'none';
				}
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = '';
					document.getElementById('team1').style.display = 'none';
				}
				if (eval("document.frm.view_c[2].checked")) {
					document.getElementById('bonbu1').style.display = 'none';
					document.getElementById('saupbu1').style.display = 'none';
					document.getElementById('team1').style.display = '';
				}
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_org_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>검색 조건</dt>
                        <dd>
                            <p>
                               <strong>회사</strong>
                              <%
								objBuilder.Append "SELECT org_name FROM emp_org_mst "
								objBuilder.Append "WHERE (org_level = '회사') ORDER BY org_code ASC "

								Set rs_org = Server.CreateObject("ADODB.Recordset")
	                            rs_org.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">

                			  <%
								Do Until rs_org.EOF
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") Then %>selected<% End If %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.MoveNext()
								Loop
								rs_org.Close() : Set rs_org = Nothing
							  %>
            					</select>
                                </label>
								<label>
									<input type="radio" name="view_c" value="bonbu" <% If view_c = "bonbu" Then %>checked<% End If %> style="width:25px" onClick="condi_view()">본부
								</label>
								<label>
									<input type="radio" name="view_c" value="saupbu" <% If view_c = "saupbu" Then %>checked<% End If %> style="width:25px" onClick="condi_view()">사업부
								</label>
				                <label>
									<input type="radio" name="view_c" value="team" <% If view_c = "team" Then %>checked<% End If %> style="width:25px" onClick="condi_view()">팀
								</label>
                                <label id="bonbu1">
								 <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;본부 명</strong>
                                	<input name="field_bonbu" type="text" value="<%=field_bonbu%>" style="width:120px; text-align:left; ime-mode:active" id="field_view">
								 </label>
								 <label id="saupbu1">
								 <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;사업부 명</strong>
                                	<input name="field_saupbu" type="text" value="<%=field_saupbu%>" style="width:120px; text-align:left; ime-mode:active" id="field_view">
								 </label>
                                 <label id="team1">
								 <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;팀 명</strong>
                                	<input name="field_team" type="text" value="<%=field_team%>" style="width:120px; text-align:left; ime-mode:active" id="field_view">
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
				      <col width="9%" >
                      <col width="6%" >
                      <col width="4%" >
				      <col width="5%" >
				      <col width="6%" >
                      <col width="8%" >
				      <col width="8%" >
				      <col width="8%" >
				      <col width="8%" >
                      <col width="11%" >
				      <col width="6%" >
                      <col width="5%" >
				      <col width="5%" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th colspan="4" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;직&nbsp;&nbsp;장</th>
                        <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
				        <th rowspan="2" scope="col">상주회사</th>
                        <th rowspan="2" scope="col">조직생성일</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">상위&nbsp;조직장</th>
                        <th rowspan="2" scope="col">수정</th>
			          </tr>
                      <tr>
				        <th class="first"scope="col">코드</th>
				        <th scope="col">조직명</th>
                        <th scope="col">조직<br>Lvel</th>
                        <th scope="col">T.O</th>
				        <th scope="col">사번</th>
				        <th scope="col">성명</th>
                        <th scope="col">회&nbsp;&nbsp;사</th>
				        <th scope="col">본&nbsp;&nbsp;부</th>
				        <th scope="col">사업부</th>
				        <th scope="col">팀</th>
				        <th scope="col">사번</th>
                        <th scope="col">성명</th>
                      </tr>
			        </thead>
				    <tbody>
					<%
					Dim view_sort, date_sw

					objBuilder.Append "SELECT org_code, org_name, org_level, org_table_org, org_empno, "
					objBuilder.Append "org_emp_name, org_company, org_bonbu, org_saupbu, org_team, "
					objBuilder.Append "org_reside_company, org_date, org_owner_empno, org_owner_empname "
					objBuilder.Append "FROM emp_org_mst "
					objBuilder.Append owner_sql & order_sql
					objBuilder.Append " LIMIT "& stpage &"," &pgsize

					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.Open objBuilder.ToString(), DBConn, 1
					objBuilder.Clear()

					Do Until rs.EOF
					%>
				      <tr>
				        <td class="first"><%=rs("org_code")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_org_view.asp?org_code=<%=rs("org_code")%>&org_name=<%=org_name%>&u_type=<%="U"%>','insa_org_view_pop','scrollbars=yes,width=750,height=350')"><%=rs("org_name")%></a>&nbsp;</td>
                        <td><%=rs("org_level")%>&nbsp;</td>
                        <td><%=rs("org_table_org")%>&nbsp;</td>
                        <td><%=rs("org_empno")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("org_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("org_emp_name")%></a>
						</td>
                        <td><%=rs("org_company")%>&nbsp;</td>
				        <td><%=rs("org_bonbu")%>&nbsp;</td>
                        <td><%=rs("org_saupbu")%>&nbsp;</td>
                        <td><%=rs("org_team")%>&nbsp;</td>
                        <td><%=rs("org_reside_company")%>&nbsp;</td>
                        <td><%=rs("org_date")%>&nbsp;</td>
                        <td><%=rs("org_owner_empno")%>&nbsp;</td>
                        <td><%=rs("org_owner_empname")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_org_reg.asp?org_code=<%=rs("org_code")%>&view_condi=<%=view_condi%>&u_type=<%="U"%>','insa_org_modi_pop','scrollbars=yes,width=1250,height=400')">수정</a>&nbsp;</td>
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
				Dim intstart, intend, first_page, field_view
				Dim i

                intstart = (Int((page-1)/10)*10) + 1
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
                    <a href="insa_excel_org.asp?view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_org_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&ck_sw=<%="y"%>">[처음]</a>
                        <% If intstart > 1 Then %>
                            <a href="insa_org_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&ck_sw=<%="y"%>">[이전]</a>
                        <% End If %>
                        <% For i = intstart To intend %>
                            <% If i = Int(page) Then %>
                                <b>[<%=i%>]</b>
                            <% Else %>
                                <a href="insa_org_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <% End If %>
                        <% next %>
                        <% If intend < total_page Then %>
                            <a href="insa_org_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_org_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_bonbu=<%=field_bonbu%>&field_saupbu=<%=field_saupbu%>&field_team=<%=field_team%>&ck_sw=<%="y"%>">[마지막]</a>
                        <% Else %>
                            [다음]&nbsp;[마지막]
                        <% End If %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_org_reg.asp?view_condi=<%=view_condi%>','insa_org_reg_popup','scrollbars=yes,width=1250,height=400')" class="btnType04">신규조직등록</a>
					</div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
        <input type="hidden" name="field_check" value="<%=field_view%>" ID="field_check">
	</body>
</html>

