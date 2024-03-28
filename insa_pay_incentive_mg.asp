<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

be_pg = "insa_pay_incentive_mg.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

in_pmg_id=Request.form("in_pmg_id")

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")
owner_view=request("owner_view")

ck_sw=Request("ck_sw")

if ck_sw = "y" then
	view_condi = request("view_condi")
	in_pmg_id = request("in_pmg_id") 
	pmg_yymm=request("pmg_yymm")
    to_date=request("to_date") 
	owner_view=request("owner_view")
	condi = request("condi")
else
	view_condi = request.form("view_condi")
	in_pmg_id = Request.Form("in_pmg_id") 
	pmg_yymm=Request.form("pmg_yymm")
    to_date=Request.form("to_date")
	owner_view=Request.form("owner_view")
	condi = request.form("condi")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	in_pmg_id = "2"
	condi = ""
	owner_view = "C"
	ck_sw = "n"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
	to_date = ""
end if

give_date = to_date '지급일


' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

' 분기 테이블 생성
curr_mm = mid(now(),6,2)
if curr_mm > 0 and curr_mm < 4 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "1"
end if
if curr_mm > 3 and curr_mm < 7 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "2"
end if
if curr_mm > 6 and curr_mm < 10 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "3"
end if
if curr_mm > 9 and curr_mm < 13 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "4"
end if

quarter_tab(8,2) = cstr(mid(quarter_tab(8,1),1,4)) + "년 " + cstr(mid(quarter_tab(8,1),5,1)) + "/4분기"

for i = 7 to 1 step -1
	cal_quarter = cint(quarter_tab(i+1,1)) - 1
	if cstr(mid(cal_quarter,5,1)) = "0" then
		quarter_tab(i,1) = cstr(cint(mid(cal_quarter,1,4))-1) + "4"
	  else
		quarter_tab(i,1) = cal_quarter
	end if	 
	quarter_tab(i,2) = cstr(mid(quarter_tab(i,1),1,4)) + "년 " + cstr(mid(quarter_tab(i,1),5,1)) + "/4분기"
next

'당월 퇴사자 포함
st_es_date = mid(cstr(pmg_yymm),1,4) + "-" + mid(cstr(pmg_yymm),5,2) + "-" + "01"


' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'Sql = "SELECT * FROM k1_memb where "+condi_sql+"mg_group = '"+mg_group+"' ORDER BY user_name ASC"
'where_sql = " WHERE isNull(emp_end_date) or emp_end_date = '1900-01-01'"

if condi = "" then
       Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_es_date&"') and (emp_company = '"+view_condi+"')  and (emp_no < '900000')"
   else  
      if owner_view = "C" then 
            Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_es_date&"') and (emp_company = '"+view_condi+"') and (emp_name like '%"+condi+"%')"
		 else	
            Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_es_date&"') and (emp_company = '"+view_condi+"') and (emp_no = '"+condi+"')"
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
       Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_es_date&"') and (emp_company = '"+view_condi+"')  and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC limit "& stpage & "," &pgsize 
   else 
       if owner_view = "C" then  
             Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_es_date&"') and (emp_company = '"+view_condi+"') and (emp_name like '%"+condi+"%') ORDER BY emp_in_date,emp_no ASC limit "& stpage & "," &pgsize 
          else
             Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_es_date&"') and (emp_company = '"+view_condi+"') and (emp_no = '"+condi+"') ORDER BY emp_in_date,emp_no ASC limit "& stpage & "," &pgsize 
	   end if
end if
'Response.write sql&"<br>"
Rs.Open Sql, Dbconn, 1

if in_pmg_id = "2" then 
    pmg_id_name = "상여금" 
elseif in_pmg_id = "3" then 
    pmg_id_name = "추천인인센티브" 
elseif in_pmg_id = "4" then 
	pmg_id_name = "연차수당" 
end if
		  
title_line = ""+ pmg_id_name +" - 자료 입력 "

etc_code = "9999"

sql = "select * from emp_etc_code where emp_etc_code = '" + etc_code + "'"
'Response.write sql&"<br>"
Rs_etc.Open Sql, Dbconn, 1
emp_payend_date = Rs_etc("emp_payend_date")
emp_payend_yn = Rs_etc("emp_payend_yn")

Rs_etc.close()

if pmg_yymm > emp_payend_date then
       emp_payend = "N"
else   
	   emp_payend = "Y"
end if   

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				return "0 1";
			}
		</script>
		<script type="text/javascript">
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_incentive_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where  org_level = '회사' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:120px">
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
								<strong>귀속년월 : </strong>
                                <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
								<label>
								<strong>지급일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<label>
                                <strong>수당구분</strong>
                                <select name="in_pmg_id" id="in_pmg_id" type="text" value="<%=in_pmg_id%>" style="width:100px">
                                    <option value="2" <%If in_pmg_id = "2" then %>selected<% end if %>>상여금</option>
                                    <option value="3" <%If in_pmg_id = "3" then %>selected<% end if %>>추천인인센티브</option>
                                    <option value="4" <%If in_pmg_id = "4" then %>selected<% end if %>>연차수당</option>
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
                            <col width="6%" >
                            <col width="9%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="*" >
                            <col width="3%" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">최초입사일</th>
                                <th scope="col">입사일</th>
                                <th scope="col">소속</th>
                                <%
								  if in_pmg_id = "2" then %>
								  <th scope="col">상여금</th>
                                <%   elseif in_pmg_id = "3" then %>
                                     <th scope="col">추천인<br>인센티브</th>
                                <%          elseif in_pmg_id = "4" then %>
                                            <th scope="col">연차수당</th>
                                <% end if %>
                                <th scope="col">고용보험</th>
                                <th scope="col">세액계</th>
                                <th scope="col">차인지급액</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <th colspan="2" scope="col">자료</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  emp_no = rs("emp_no")
	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td>
                                    <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>								
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                        <%
                              Sql = "SELECT * FROM pay_month_give where pmg_yymm = '"&pmg_yymm&"' and pmg_emp_no = '"&emp_no&"' and pmg_id = '"&in_pmg_id&"' and (pmg_company = '"+view_condi+"')"
                              'Response.write sql&"<br>"                              
                              Set rs_give = DbConn.Execute(SQL)
							  if not rs_give.eof then
                                    pmg_base_pay = rs_give("pmg_base_pay")
									pmg_give_tot = rs_give("pmg_give_total")
	                             else
                                    pmg_base_pay = 0
								    pmg_give_tot = 0
                              end if
                              rs_give.close()
                        %>
                                <td class="right"><%=formatnumber(pmg_base_pay,0)%>&nbsp;</td>
                        <%
                              Sql = "SELECT * FROM pay_month_deduct where de_yymm = '"&pmg_yymm&"' and de_emp_no = '"&emp_no&"' and de_id = '"&in_pmg_id&"' and (de_company = '"+view_condi+"')"
                              'Response.write sql&"<br>"                              
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
									de_epi_amt = Rs_dct("de_epi_amt")
									de_income_tax = Rs_dct("de_income_tax")
									de_wetax = Rs_dct("de_wetax")
									de_deduct_tot = Rs_dct("de_deduct_total")
	                             else
                                    de_deduct_tot = 0
									de_epi_amt = 0
									de_income_tax = 0
									de_wetax = 0
                              end if
                              Rs_dct.close()
							  
							  de_tax = de_income_tax + de_wetax
							  pmg_curr_pay = pmg_give_tot - de_deduct_tot
                        %>
                              
                                <td class="right"><%=formatnumber(de_epi_amt,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(de_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pmg_curr_pay,0)%>&nbsp;</td>
                                
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
                        <% if emp_payend = "N" then %> 
                                <td><a href="#" onClick="pop_Window('insa_pay_incentive_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&pmg_yymm=<%=pmg_yymm%>&in_pmg_id=<%=in_pmg_id%>&give_date=<%=give_date%>&view_condi=<%=view_condi%>&u_type=<%=""%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=430')">입력</a></td>
                                <td><a href="#" onClick="pop_Window('insa_pay_incentive_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&pmg_yymm=<%=pmg_yymm%>&in_pmg_id=<%=in_pmg_id%>&give_date=<%=give_date%>&view_condi=<%=view_condi%>&u_type=<%="U"%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=430')">수정</a></td>
                        <%     else %>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                        <% end if %>
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
                    <a href="insa_excel_pay_incentive.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_id=<%=in_pmg_id%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_incentive_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&in_pmg_id=<%=in_pmg_id%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_incentive_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&in_pmg_id=<%=in_pmg_id%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_incentive_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&in_pmg_id=<%=in_pmg_id%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_incentive_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&in_pmg_id=<%=in_pmg_id%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_incentive_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&in_pmg_id=<%=in_pmg_id%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
                    <% if emp_payend = "N" then %>
					<a href="#" onClick="pop_Window('insa_pay_incentive_upload.asp?emp_no=<%=in_empno%>&emp_name=<%=in_name%>&pmg_yymm=<%=pmg_yymm%>&give_date=<%=give_date%>&view_condi=<%=view_condi%>&u_type=<%=""%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=700')" class="btnType04">자료UPload</a>
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

