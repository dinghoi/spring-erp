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
Dim page, page_cnt, pg_cnt, be_pg, curr_date
Dim ck_sw, view_condi, condi, start_page, stpage, pgsize
Dim order_sql, where_sql, rsCount, total_record, total_page
Dim title_line, rsOrg, rs_org, date_sw, i
Dim intstart, intend, first_page, pageNavi, view_sort
Dim pg_url

page = f_Request("page")
page_cnt = f_Request("page_cnt")
pg_cnt = CInt(f_Request("pg_cnt"))
view_condi = f_Request("view_condi")
condi = f_Request("condi")

be_pg = "/insa/insa_org_name_view.asp"
curr_date = DateValue(Mid(CStr(Now()), 1, 10))
title_line = " 조직명 조회 "

pageNavi = "/insa/insa_org_name_view.asp"

If view_condi = "" Then
	view_condi = "케이원"
	condi = ""
End If

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	Page = 1
	start_page = 1
End If
stpage = Int((page - 1) * pgsize)

pg_url = "&view_condi="&view_condi&"&condi="&condi

order_Sql = " ORDER BY org_company, org_bonbu, org_saupbu, org_team, org_name ASC"

where_sql = " WHERE (isNull(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '0000-00-00') "
where_sql = where_sql & "AND (org_company = '"&view_condi&"') AND (org_name like '%"&condi&"%') "

objbuilder.Append "SELECT COUNT(*) FROM emp_org_mst "&where_sql

Set rsCount = Dbconn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

objBuilder.Append "SELECT org_code, org_name, org_level, org_table_org, org_empno, "
objBuilder.Append "	org_emp_name, org_company, org_bonbu, org_saupbu, org_team, org_name, "
objBuilder.Append "	org_reside_company, org_date, org_owner_empno, org_owner_empname "
objBuilder.Append "FROM emp_org_mst " & where_sql & order_sql & " LIMIT "& stpage & "," &pgsize

Set rsOrg = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				return "5 2";
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if (document.frm.view_condi.value == ""){
					alert("필드조건을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_org_name_view.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>회사 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
							   <%
								objBuilder.Append "SELECT org_name FROM emp_org_mst "
								objBuilder.Append "WHERE (isNull(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '0000-00-00') "
								objBuilder.Append "AND org_level = '회사' "
								objBuilder.Append "ORDER BY org_code ASC "

								Set rs_org = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  	%>
                                <label>
								<select type="text" name="view_condi" id="view_condi" style="width:150px;">
                			  <%
								Do Until rs_org.EOF
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") Then%>selected<%End If %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.MoveNext()
								Loop
								rs_org.Close() : Set rs_org = Nothing
							  %>
            					</select>
                                </label>
                                <label>
                                <strong>조직명 : </strong>
									<input type="text" name="condi" value="<%=condi%>" style="width:150px; text-align:left;" >
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>

				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="3%" >
				      <col width="10%" >
                      <col width="6%" >
                      <col width="4%" >
				      <col width="4%" >
				      <col width="6%" >
                      <col width="8%" >
				      <col width="8%" >
				      <col width="8%" >
				      <col width="8%" >
                      <col width="8%" >
				      <col width="6%" >
                      <col width="5%" >
				      <col width="6%" >
                      <col width="3%" >
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
                        <th rowspan="2" scope="col">상위<br>조직</th>
                        <th rowspan="2" scope="col">기타<br>정보</th>
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
						Do Until rsOrg.EOF
					%>
				      <tr>
				        <td class="first"><%=rsOrg("org_code")%>&nbsp;</td>
						<td>
							<a href="#" onClick="pop_Window('/insa/insa_org_view.asp?org_code=<%=rsOrg("org_code")%>&org_name=<%=org_name%>&u_type=U','insa_org_view_pop','scrollbars=yes,width=750,height=350')"><%=rsOrg("org_name")%></a>&nbsp;
						</td>
						<td><%=rsOrg("org_level")%>&nbsp;</td>
						<td><%=rsOrg("org_table_org")%>&nbsp;</td>
						<td><%=rsOrg("org_empno")%>&nbsp;</td>
						<td>
                			<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsOrg("org_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=670')"><%=rsOrg("org_emp_name")%></a>
						</td>
                		<td><%=rsOrg("org_company")%>&nbsp;</td>
				        <td><%=rsOrg("org_bonbu")%>&nbsp;</td>
						<td><%=rsOrg("org_saupbu")%>&nbsp;</td>
						<td><%=rsOrg("org_team")%>&nbsp;</td>
						<td><%=rsOrg("org_reside_company")%>&nbsp;</td>
						<td><%=rsOrg("org_date")%>&nbsp;</td>
						<td><%=rsOrg("org_owner_empno")%>&nbsp;</td>
						<td><%=rsOrg("org_owner_empname")%>&nbsp;</td>
						<td>
							<a href="#" onClick="pop_Window('/insa/insa_org_owner_modify.asp?org_code=<%=rsOrg("org_code")%>&u_type=<%="U"%>','insa_org_modi_pop','scrollbars=yes,width=1250,height=400')">변경</a>&nbsp;
						</td>
						<td>
							<a href="#" onClick="pop_Window('/insa/insa_org_modify.asp?org_code=<%=rsOrg("org_code")%>&u_type=<%="U"%>','insa_org_modi_pop','scrollbars=yes,width=1250,height=400')">수정</a>&nbsp;
						</td>
			          </tr>
				      <%
							rsOrg.MoveNext()
						Loop
						rsOrg.close() : Set rsOrg = Nothing
						%>
			        </tbody>
			      </table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)

					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
			      </tr>
			  </table>
			</form>
		</div>
	</div>
		<input type="hidden" name="user_id" />
		<input type="hidden" name="pass" />
	</body>
</html>