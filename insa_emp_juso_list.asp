<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_emp_juso_list.asp"

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_condi = "전체"
	condi_sql = " "
	condi = ""
end if

if view_condi = "소속조직별" then
	condi_sql = "emp_org_name like '%" + condi + "%' and "
end if
if view_condi = "성명" then
	condi_sql = "emp_name like '%" + condi + "%' and "
end if
if view_condi = "회사별" then
	condi_sql = "emp_company like '%" + condi + "%' and "
end if
if view_condi = "본부별" then
	condi_sql = "emp_bonbu like '%" + condi + "%' and "
end if
if view_condi = "사업부별" then
	condi_sql = "emp_saupbu like '%" + condi + "%' and "
end if
if view_condi = "팀별" then
	condi_sql = "emp_team like '%" + condi + "%' and "
end if
if view_condi = "상주처 회사별" then
	condi_sql = "emp_reside_company like '%" + condi + "%' and "
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


Sql = "SELECT count(*) FROM emp_master where "+condi_sql+" (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM emp_master where "+condi_sql+" (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_no,emp_name ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = ""+ view_condi +" - 직원 주소록 "
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_juso_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="소속조직별" <%If view_condi = "소속조직별" then %>selected<% end if %>>소속조직별</option>
                                  <option value="성명" <%If view_condi = "성명" then %>selected<% end if %>>성명</option>
                                  <option value="회사별" <%If view_condi = "회사별" then %>selected<% end if %>>회사별</option>
                                  <option value="본부별" <%If view_condi = "본부별" then %>selected<% end if %>>본부별</option>
                                  <option value="사업부별" <%If view_condi = "사업부별" then %>selected<% end if %>>사업부별</option>
                                  <option value="팀별" <%If view_condi = "팀별" then %>selected<% end if %>>팀별</option>
                                  <option value="상주처 회사별" <%If view_condi = "상주처 회사별" then %>selected<% end if %>>상주처 회사별</option>
                                </select>
								<strong>조건 : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="15%" >
                            <col width="11%" >
                            <col width="11%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col" class="first">소속</th>
                                <th scope="col">성  명</th>
								<th scope="col">직위</th>
								<th scope="col">직책</th>
								<th scope="col">메일주소</th>
                                <th scope="col">내선번호</th>
                                <th scope="col">휴대전화</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

                        emp_email = rs("emp_email") + "@k-won.co.kr"
	           			%>
							<tr>
                                <td class="first"><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><a href="#" onClick="pop_Window('insa_emp_card.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_emp_card_pop','scrollbars=yes,width=500,height=500')"><%=rs("emp_name")%></a>&nbsp;</td>

                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td class="left"><%=emp_email%>&nbsp;</td>
                                <td><%=rs("emp_extension_no")%>&nbsp;</td>
                                <td><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%>&nbsp;</td>
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
				    <td>
                    <div id="paging">
                        <a href = "insa_emp_juso_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_emp_juso_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_emp_juso_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_emp_juso_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_emp_juso_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
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

