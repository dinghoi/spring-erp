<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim year_tab(3,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

be_pg = "insa_pay_year_nps_mg.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

'당월 입사일이 15일 이전이면 당월 급여대상임
st_es_date = mid(cstr(curr_date),1,4) + "-" + mid(cstr(curr_date),6,2) + "-" + "01"
st_in_date = mid(cstr(curr_date),1,4) + "-" + mid(cstr(curr_date),6,2) + "-" + "16"

' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")
owner_view=request("owner_view")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	owner_view=Request.form("owner_view")
	condi = request.form("condi")
	inc_yyyy=Request.form("inc_yyyy")
  else
	view_condi = request("view_condi")
	owner_view=request("owner_view")
	condi = request("condi")
	inc_yyyy=request("inc_yyyy") 
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	condi = ""
	owner_view = "C"
	ck_sw = "n"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	inc_yyyy = mid(cstr(from_date),1,4)
end if

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'Sql = "SELECT * FROM k1_memb where "+condi_sql+"mg_group = '"+mg_group+"' ORDER BY user_name ASC"
'where_sql = " WHERE isNull(emp_end_date) or emp_end_date = '1900-01-01'"

if condi = "" then
       Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_company = '"+view_condi+"')  and (emp_no < '900000')"
   else  
      if owner_view = "C" then 
            Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_company = '"+view_condi+"') and (emp_name like '%"+condi+"%')"
		 else	
            Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_company = '"+view_condi+"') and (emp_no = '"+condi+"')"
	  end if
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if condi = "" then
       Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_company = '"+view_condi+"')  and (emp_no < '900000') ORDER BY emp_no ASC limit "& stpage & "," &pgsize 
   else 
       if owner_view = "C" then  
             Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_company = '"+view_condi+"') and (emp_name like '%"+condi+"%') ORDER BY emp_no ASC limit "& stpage & "," &pgsize 
          else
             Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_company = '"+view_condi+"') and (emp_no = '"+condi+"') ORDER BY emp_no ASC limit "& stpage & "," &pgsize 
	   end if
end if

Rs.Open Sql, Dbconn, 1

title_line = "국민연금 표준월액 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				return "3 1";
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
			<!--#include virtual = "/include/insa_pay_income_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_year_nps_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '회사' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">

                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <label>
							<strong>귀속년도 : </strong>
                                <select name="inc_yyyy" id="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:90px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If inc_yyyy = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                </select>
								</label>
							    <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">사번
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">성명
                                </label>
							<strong>조건 : </strong>
								<label>
        						<input name="condi" type="text" id="condi" value="<%=condi%>" style="width:100px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
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
                            <col width="7%" >
							<col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="*" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
                                <th scope="col">소속</th>
                                <th scope="col">표준보수<br>월액</th>
								<th scope="col">국민연금료</th>
                                <th scope="col">고용보험<br>적용</th>
                                <th scope="col">산재보험<br>적용</th>
                                <th scope="col">장기요양<br>보험적용</th>
                                <th scope="col">청년<br>소득세감면</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <th scope="col">연금</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
                              bank_name = ""
							  account_no = ""
							  account_holder = ""
							  emp_no = rs("emp_no")
	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                        <%
						      Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&inc_yyyy&"'"
                              Set rs_year = DbConn.Execute(SQL)
							  if not rs_year.eof then
                                    incom_nps_amount = rs_year("incom_nps_amount")
								    incom_nps = rs_year("incom_nps")
									incom_go_yn = rs_year("incom_go_yn")
                                    incom_san_yn = rs_year("incom_san_yn")
                                    incom_long_yn = rs_year("incom_long_yn")
                                    incom_incom_yn = rs_year("incom_incom_yn")
	                             else
                                    incom_nps_amount = 0
								    incom_nps = 0
									incom_go_yn = ""
                                    incom_san_yn = ""
                                    incom_long_yn = ""
                                    incom_incom_yn = ""
                              end if
                              rs_year.close()
                          %>
                                <td class="right"><%=formatnumber(incom_nps_amount,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(incom_nps,0)%>&nbsp;</td>
                                <td><%=incom_go_yn%>&nbsp;</td>
                                <td><%=incom_san_yn%>&nbsp;</td>
                                <td><%=incom_long_yn%>&nbsp;</td>
                                <td><%=incom_incom_yn%>&nbsp;</td>
                                <td class="left"><%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
                                <td><a href="#" onClick="pop_Window('insa_pay_year_nps_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&inc_yyyy=<%=inc_yyyy%>&u_type=<%="U"%>','insa_income_add_pop','scrollbars=yes,width=800,height=280')">등록</a></td>
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
                    <a href="insa_excel_year_income.asp?view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_year_nps_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_year_nps_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_year_nps_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_year_nps_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_year_nps_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
                    <% if owner_view = "T" then 
                              emp_no = condi
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
							  end if
							  rs_emp.close()
				    %>
					<a href="#" onClick="pop_Window('insa_pay_year_nps_add.asp?emp_no=<%=emp_no%>&inc_yyyy=<%=inc_yyyy%>&emp_name=<%=emp_name%>&u_type=<%=""%>','insa_income_add_pop','scrollbars=yes,width=800,height=280')" class="btnType04">국민연금표준월액 등록</a>
					</div> 
                    <% end if %>                 
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

