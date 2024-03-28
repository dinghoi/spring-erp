<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_mg_list.asp"

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
  else
	view_condi = request("view_condi")
end if

if view_condi = "" then
	view_condi = "emp_image"
end if


pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code,emp_in_date,emp_no ASC"

where_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"

field_sql = " and ( " + view_condi + " = '' or isNull(" + view_condi + ")) "

'if view_condi = "emp_image" then 
'        where_sql = " WHERE (emp_image = '' or isNull(emp_image)) and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
'   elseif view_condi = "emp_person1" then 
'                where_sql = " WHERE (emp_person1 = '' or isNull(emp_person1) or emp_person1 = '' or isNull(emp_person1)) and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
'end if				

Sql = "SELECT count(*) FROM emp_master " + where_sql + field_sql
'response.write(sql)
'response.End()
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_master " + where_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
'Sql = "SELECT * FROM emp_master where "+condi_sql+"isNull(emp_end_date) ORDER BY emp_no,emp_name ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " 직원 현황-인사자료미등록- "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_mg_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="cost_center" <%If view_condi = "cost_center" then %>selected<% end if %>>비용배분구분</option>
                                  <option value="emp_image" <%If view_condi = "emp_image" then %>selected<% end if %>>사진</option>
                                  <option value="emp_ename" <%If view_condi = "emp_ename" then %>selected<% end if %>>영문명</option>
                                  <option value="emp_person1" <%If view_condi = "emp_person1" then %>selected<% end if %>>주민등록번호</option>
                                  <option value="emp_birthday" <%If view_condi = "emp_birthday" then %>selected<% end if %>>생년월일</option>
                                  <option value="emp_family_sido" <%If view_condi = "emp_family_sido" then %>selected<% end if %>>본적</option>
                                  <option value="emp_sido" <%If view_condi = "emp_sido" then %>selected<% end if %>>주소</option>
                                  <option value="emp_tel_no1" <%If view_condi = "emp_tel_no1" then %>selected<% end if %>>전화번호</option>
                                  <option value="emp_hp_no1" <%If view_condi = "emp_hp_no1" then %>selected<% end if %>>핸드폰</option>
                                  <option value="emp_emergency_tel" <%If view_condi = "emp_emergency_tel" then %>selected<% end if %>>비상연락</option>
                                  <option value="emp_email" <%If view_condi = "emp_email" then %>selected<% end if %>>이메일</option>
                                  <option value="emp_extension_no" <%If view_condi = "emp_extension_no" then %>selected<% end if %>>내선번호</option>
                                  <option value="emp_last_edu" <%If view_condi = "emp_last_edu" then %>selected<% end if %>>최종학력</option>
                                </select>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
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
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">생년월일</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
                                <th scope="col">소속</th>
                                <th scope="col">최초입사일</th>
								<th scope="col">소속발령일</th>
								<th scope="col">상주처</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

						if rs("emp_org_baldate") = "1900-01-01" then
						   emp_org_baldate = ""
						   else 
						   emp_org_baldate = rs("emp_org_baldate")
						end if
						if rs("emp_birthday") = "1900-01-01" then
						   emp_birthday = ""
						   else 
						   emp_birthday = rs("emp_birthday")
						end if
						if rs("emp_grade_date") = "1900-01-01" then
						   emp_grade_date = ""
						   else 
						   emp_grade_date = rs("emp_grade_date")
						end if

	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=emp_birthday%>&nbsp;</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rs("emp_reside_place")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_emplist2.asp?view_condi=<%=view_condi%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_mg_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_mg_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_mg_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_mg_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_mg_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[마지막]</a>
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

