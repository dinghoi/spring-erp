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
Dim page, view_codi, condi, owner_view, be_pg, view_condi
Dim pgsize, start_page, stpage, rsCount, total_record, total_page, pg_url
Dim base_sql, condi_sql, rsEmp, title_line, rs_org

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")
owner_view = f_Request("owner_view")

be_pg = "/pay/insa_bank_account_mg.asp"

If view_condi = "" Then
	view_condi = "케이원"
	condi = ""
	owner_view = "C"
End If

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&condi="&condi&"&owner_view="&owner_view

base_sql = "FROM emp_master AS emtt "
base_sql = base_sql & "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
base_sql = base_sql & "LEFT OUTER JOIN pay_bank_account AS pbat ON emtt.emp_no = pbat.emp_no "
base_sql = base_sql & "WHERE (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01') "

If condi = "" Then
	'Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"&view_condi&"')  and (emp_no < '900000')"
	condi_sql = "	AND emtt.emp_no < '900000' "
Else
	If owner_view = "C" Then
		'Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"&view_condi&"') and (emp_name like '%"&condi&"%')"
		condi_sql = "	AND emtt.emp_name LIKE '%"&condi&"%' "
	Else
		'Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"&view_condi&"') and (emp_no = '"&condi&"')"
		condi_sql = "	AND emtt.emp_no = '"&condi&"' "
	End If
End If

'전체 Count
objBuilder.Append "SELECT COUNT(*) " & base_sql
objBuilder.Append "	AND eomt.org_company = '"&view_condi&"' " & condi_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'Result.RecordCount
total_record = CInt(rsCount(0))

'Result.PageCount
If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize)
Else
	total_page = Int((total_record / pgsize) + 1)
End If

rsCount.Close() : Set rsCount = Nothing

title_line = "직원 은행계좌 현황 "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
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
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}

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
			}//-->
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_code_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_bank_account_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								'Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '회사' ORDER BY org_code ASC"
								objBuilder.Append "SELECT org_name FROM emp_org_mst "
								objBuilder.Append "WHERE (isNull(org_end_date) OR org_end_date = '0000-00-00') "
								objBuilder.Append "	AND org_level = '회사' "
								objBuilder.Append "ORDER BY org_code ASC "

								Set rs_org = DBConn.Execute(objBuilder.ToString())
	                            objBuilder.Clear()
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
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
									<input name="owner_view" type="radio" value="T" <%If owner_view = "T" Then %>checked<%End If %> style="width:25px">사번
									<input name="owner_view" type="radio" value="C" <%If owner_view = "C" Then %>checked<%End If %> style="width:25px">성명
                                </label>
								<strong>조건 : </strong>
								<label>
        							<input name="condi" type="text" id="condi" value="<%=condi%>" style="width:100px; text-align:left">
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
                            <col width="6%" >
							<col width="12%" >
                            <col width="9%" >
							<col width="*" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
                                <th scope="col">소속</th>
                                <th scope="col">거래은행</th>
								<th scope="col">계좌번호</th>
                                <th scope="col">예금주</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <th colspan="2" scope="col">은행계좌</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim bank_name, account_no, account_holder, emp_person1, emp_person2

						If view_condi <> "" Then
							objBuilder.Append "SELECT emtt.emp_no, emtt.emp_person1, emtt.emp_person2, emtt.emp_name, "
							objBuilder.Append "	emtt.emp_grade, emtt.emp_position, emtt.emp_in_date, "
							objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
							objBuilder.Append "	pbat.bank_name, pbat.account_no, pbat.account_holder "
							objBuilder.Append base_sql
							objBuilder.Append "	AND eomt.org_company = '"&view_condi&"' " & condi_sql
							objBuilder.Append "ORDER BY emtt.emp_no, emtt.emp_name ASC "
							objBuilder.Append "LIMIT "& stpage & "," &pgsize

							Set rsEmp = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							Do Until rsEmp.EOF
								emp_no = rsEmp("emp_no")
								emp_person1 = rsEmp("emp_person1")
								emp_person2 = rsEmp("emp_person2")

								bank_name = rsEmp("bank_name")
								account_no = rsEmp("account_no")
								account_holder = rsEmp("account_holder")
	           			%>
							<tr>
								<td class="first"><%=rsEmp("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsEmp("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>','인사 카드','scrollbars=yes,width=1250,height=650')"><%=rsEmp("emp_name")%></a>
								</td>
                                <td><%=rsEmp("emp_grade")%>&nbsp;</td>
                                <td><%=rsEmp("emp_position")%>&nbsp;</td>
                                <td><%=rsEmp("emp_in_date")%>&nbsp;</td>
                                <td><%=rsEmp("org_name")%>&nbsp;</td>
                                <td><%=bank_name%>&nbsp;</td>
                                <td><%=account_no%>&nbsp;</td>
                                <td><%=account_holder%>&nbsp;</td>
                                <td class="left">
								<%
								Call EmpOrgInSaupbuText(rsEmp("org_company"), rsEmp("org_bonbu"), rsEmp("org_saupbu"), rsEmp("org_team"))
								%>
								</td>
                                <td><a href="#" onClick="pop_Window('/pay/insa_bank_account_add.asp?emp_no=<%=rsEmp("emp_no")%>&emp_name=<%=rsEmp("emp_name")%>&emp_person1=<%=rsEmp("emp_person1")%>&emp_person2=<%=rsEmp("emp_person2")%>&u_type=U','insa_bank_add_pop','scrollbars=yes,width=750,height=300')">수정</a></td>
                                <td><a href="#" onClick="pop_Window('/pay/insa_bank_account_add.asp?emp_no=<%=rsEmp("emp_no")%>&emp_name=<%=rsEmp("emp_name")%>&emp_person1=<%=rsEmp("emp_person1")%>&emp_person2=<%=rsEmp("emp_person2")%>','insa_bank_add_pop','scrollbars=yes,width=750,height=300')">등록</a></td>
							</tr>
						<%
								rsEmp.MoveNext()
							Loop
							rsEmp.close() : Set rsEmp = Nothing
						End If
						DBConn.Close : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="/pay/insa_excel_banklist.asp?view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
					%>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>