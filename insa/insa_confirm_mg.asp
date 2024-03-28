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
Dim be_pg, page, page_cnt, pg_cnt, from_date, to_date
Dim cfm_type, company, ck_sw, pgsize, start_page, stpage
Dim com_sql, type_sql, rsCount, totRecord, total_page
Dim curr_dd, rsCfm, title_line

be_pg = "/insa/insa_confirm_mg.asp"

page = Request("page")
page_cnt = Request.Form("page_cnt")
pg_cnt = CInt(Request("pg_cnt"))
from_date = Request("from_date")
to_date = Request("to_date")
cfm_type = Request("cfm_type")
company = Request("company")
ck_sw = Request("ck_sw")

If ck_sw = "n" Then
	from_date = Request.Form("from_date")
    to_date = Request.Form("to_date")
    cfm_type = Request.Form("cfm_type")
    company = Request.Form("company")
Else
	from_date = Request("from_date")
    to_date = Request("to_date")
    cfm_type = Request("cfm_type")
    company = Request("company")
End If

If to_date = "" Or from_date = "" Then
	curr_dd = CStr(DatePart("d", Now()))
	to_date = Mid(CStr(Now()), 1, 10)
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
	cfm_type = "재직증명서"
	company = "전체"
End If

title_line = " 제증명 발급 현황 "

if company = "전체" then
	com_sql = ""
else
  	'com_sql = "AND cfm_company ='"&company&"' "
	com_sql = "AND eomt.org_company = '"&company&"' "
end If

If cfm_type = "전체" Then
	type_sql = ""
Else
  	type_sql = "AND ecft.cfm_type ='"&cfm_type&"' "
End If

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

'Sql = "SELECT COUNT(*) FROM emp_confirm where "+com_sql+type_sql+" cfm_date >= '"+from_date+"' and cfm_date <= '"+to_date+"'"
objBuilder.Append "SELECT COUNT(*) FROM emp_confirm AS ecft  "
objBuilder.Append "INNER JOIN emp_master AS emtt ON ecft.cfm_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (ecft.cfm_date >= '"&from_date&"' AND ecft.cfm_date <= '"&to_date&"') "
objBuilder.Append com_sql & type_sql

Set rsCount = Dbconn.Execute(objBuilder.ToString())
objBuilder.Clear()

totRecord = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If totRecord Mod pgsize = 0 Then
	total_page = Int(totRecord / pgsize) 'Result.PageCount
Else
	total_page = Int((totRecord / pgsize) + 1)
End If

'Sql = "SELECT * FROM emp_confirm where "+com_sql+type_sql+" cfm_date >= '"+from_date+"' and cfm_date <= '"+to_date+"' ORDER BY cfm_type,cfm_seq DESC limit "& stpage & "," &pgsize
objBuilder.Append "SELECT ecft.cfm_empno, ecft.cfm_emp_name, ecft.cfm_company, ecft.cfm_org_name, ecft.cfm_date, "
objBuilder.Append "	ecft.cfm_number, ecft.cfm_seq, ecft.cfm_type, ecft.cfm_use, ecft.cfm_use_dept, "
objBuilder.Append "	ecft.cfm_person1, ecft.cfm_person2, ecft.cfm_comment, "
objBuilder.Append "	emtt.emp_position, emtt.emp_job, "
objBuilder.Append "	eomt.org_name, eomt.org_company "
objBuilder.Append "FROM emp_confirm AS ecft "
objBuilder.Append "INNER JOIN emp_master AS emtt ON ecft.cfm_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (ecft.cfm_date >= '"&from_date&"' AND ecft.cfm_date <= '"&to_date&"') "
objBuilder.Append com_sql & type_sql
objBuilder.Append "ORDER BY cfm_type,cfm_seq DESC "
objBuilder.Append "LIMIT "& stpage & "," &pgsize

Set rsCfm = Server.CreateObject("ADODB.RecordSet")
rsCfm.Open objBuilder.ToString(), Dbconn, 1
objBuilder.Clear()
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
				return "3 1";
			}

			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});
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
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_welfare_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_confirm_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>발급일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>  ∼  </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<strong>회사</strong>
                                <%
								Dim rs_org

								'Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
								objBuilder.Append "SELECT org_name FROM emp_org_mst "
								objBuilder.Append "WHERE (isNull(org_end_date) Or org_end_date = '1900-01-01' Or org_end_date = '0000-00-00') "
								objBuilder.Append "	AND org_level = '회사' "
								objBuilder.Append "ORDER BY org_code ASC "

								Set rs_org = Server.CreateObject("ADODB.RecordSet")
	                            rs_org.Open objBuilder.ToString(), Dbconn, 1
								objBuilder.Clear()
							  %>
                                <label>
								<select name="company" id="company" type="text" style="width:150px">
									<option value="전체" <%If company = "전체" then %>selected<% end if %>>전체</option>
                			  <%
								do until rs_org.eof
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()
								Loop
								rs_org.Close() : Set rs_org = Nothing
							  %>
            					</select>
                                </label>
								<strong>제증명종류</strong>
                                <select name="cfm_type" id="cfm_type" style="width:100px">
                                    <option value="전체" <%If cfm_type = "전체" then %>selected<% end if %>>전체</option>
                                    <option value="재직증명서" <%If cfm_type = "재직증명서" then %>selected<% end if %>>재직증명서</option>
                                    <option value="경력증명서" <%If cfm_type = "경력증명서" then %>selected<% end if %>>경력증명서</option>
                                    <option value="원천징수" <%If cfm_type = "원천징수" then %>selected<% end if %>>원천징수</option>
                                    <option value="갑근세명세" <%If cfm_type = "갑근세명세" then %>selected<% end if %>>갑근세명세</option>
                                </select>
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
							<col width="6%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="*" >
						</colgroup>
						<thead>
						  <tr>
							<th class="first" scope="col">사번</th>
							<th scope="col">성명</th>
							<th scope="col">직위</th>
							<th scope="col">직책</th>
							<th scope="col">회사</th>
                            <th scope="col">소속</th>
                            <th scope="col">발급일</th>
							<th scope="col">발급번호</th>
							<th scope="col">제증명</th>
							<th scope="col">용도</th>
                            <th scope="col">사용처</th>
							<th scope="col">주민번호</th>
                            <th scope="col">비고</th>
						  </tr>
						</thead>
						<tbody>
						<%
						Dim cfm_empno, emp_job, emp_position

						Do Until rsCfm.EOF

		                  cfm_empno = rsCfm("cfm_empno")
						  emp_job = rsCfm("emp_job")
		                  emp_position = rsCfm("emp_position")
	           			%>
							<tr>
								<td class="first"><%=rsCfm("cfm_empno")%></td>
                                <td>
								 <a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsCfm("cfm_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rsCfm("cfm_emp_name")%></a>
								</td>
                                <td><%=emp_job%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=rsCfm("org_company")%>&nbsp;</td>
                                <td><%=rsCfm("org_name")%>&nbsp;</td>
                                <td><%=rsCfm("cfm_date")%>&nbsp;</td>
                                <td>제&nbsp;<%=rsCfm("cfm_number")%>-<%=rsCfm("cfm_seq")%>&nbsp;호</td>
                                <td><%=rsCfm("cfm_type")%>&nbsp;</td>
                                <td><%=rsCfm("cfm_use")%>&nbsp;</td>
								<td><%=rsCfm("cfm_use_dept")%>&nbsp;</td>
                                <td><%=rsCfm("cfm_person1")%>-<%=rsCfm("cfm_person2")%>&nbsp;</td>
                                <td><%=rsCfm("cfm_comment")%>&nbsp;</td>
							</tr>
						<%
							rsCfm.MoveNext()
						Loop
						rsCfm.close() : Set rsCfm = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page, i

                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="/insa/insa_excel_cfmlist.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="<%=be_pg%>?from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="<%=be_pg%>?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="<%=be_pg%>?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="<%=be_pg%>?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[다음]</a> <a href="<%=be_pg%>?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&cfm_type=<%=cfm_type%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>

