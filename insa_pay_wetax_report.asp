<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

be_pg = "insa_pay_wetax_report.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	pmg_yymm=Request.form("pmg_yymm")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
	sum_tax_yes = 0
	sum_tax_no = 0
	sum_tax_reduced = 0
	sum_give_tot = 0

    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_deduct_tot = 0
	
	a20_income_tax = 0
    a20_wetax = 0
	
	a30_income_tax = 0
    a30_wetax = 0
	
	a40_income_tax = 0
    a40_wetax = 0
	
	a50_income_tax = 0
    a50_wetax = 0
	
	pay_count = 0
	p_cnt = 0	
	sum_curr_pay = 0	
	
end if

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
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select count(*) from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
    pmg_give_tot = rs("pmg_give_total")
    pay_count = pay_count + 1
				  
    sum_tax_yes = sum_tax_yes + int(rs("pmg_tax_yes"))
    sum_tax_no = sum_tax_no + int(rs("pmg_tax_no"))
    sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
    if not Rs_dct.eof then

            de_income_tax = int(Rs_dct("de_income_tax"))
            de_wetax = int(Rs_dct("de_wetax"))
			de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
            de_year_wetax = int(Rs_dct("de_year_wetax"))
		    de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	     else
            de_income_tax = 0
            de_wetax = 0
			de_year_incom_tax = 0
            de_year_wetax = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()

     sum_income_tax = sum_income_tax + de_income_tax
     sum_wetax = sum_wetax + de_wetax
	 sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
     sum_year_wetax = sum_year_wetax + de_year_wetax
	 sum_deduct_tot = sum_deduct_tot + de_deduct_tot

	rs.movenext()
loop
rs.close()

a30_income_tax = sum_income_tax + a20_income_tax
a30_wetax = sum_wetax + a20_wetax


Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC limit "& stpage & "," &pgsize 

Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = " 지방소득세명세서(근로퇴직소득) "

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
				return "5 1";
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_tax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_wetax_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
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
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
                            <col width="5%" >
                            <col width="5%" >
                            <col width="*" >
                            <col width="8%" >
                            <col width="6%" >
                            <col width="8%" >
                            <col width="6%" >
                            <col width="8%" >
							<col width="6%" >
                            <col width="8%" >
                            <col width="6%" >
							<col width="8%" > 
                            <col width="6%" >
                            <col width="6%" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
				                <th rowspan="2" class="first" scope="col">순번</th>
                                <th rowspan="2" scope="col">사번</th>
                                <th rowspan="2" scope="col">성명</th>
                                <th rowspan="2" scope="col">주민번호</th>
				                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">근로소득</th>
                                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">퇴직소득</th>
                                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">합계</th>
                                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">연말정산</th>
                                <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">조정액</th>
                                <th rowspan="2" scope="col">지급년월</th>
                                <th rowspan="2" scope="col">적요</th>
			                </tr>
                            <tr>
							    <th scope="col" style=" border-left:1px solid #e3e3e3;">과세표준</th>
								<th scope="col">세액</th>  
								<th scope="col">과세표준</th>
                                <th scope="col">세액</th>
                                <th scope="col">과세표준</th>
                                <th scope="col">세액</th>
                                <th scope="col">과세표준</th>
                                <th scope="col">세액</th>
                                <th scope="col">과세표준</th>
                                <th scope="col">세액</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  p_cnt = p_cnt + 1
							  emp_no = rs("pmg_emp_no")
							  pmg_give_tot = rs("pmg_give_total")

							  sub_give_hap = int(rs("pmg_postage_pay")) + int(rs("pmg_re_pay")) + int(rs("pmg_car_pay")) + int(rs("pmg_position_pay")) + int(rs("pmg_custom_pay")) + int(rs("pmg_job_pay")) + int(rs("pmg_job_support")) + int(rs("pmg_jisa_pay")) + int(rs("pmg_long_pay")) + int(rs("pmg_disabled_pay"))
							  
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
		                      if not rs_emp.eof then
		                    		emp_in_date = rs_emp("emp_in_date")
									emp_person1 = rs_emp("emp_person1")
									emp_person2 = "*******"
	                             else
	                    			emp_in_date = ""
									emp_person1 = ""
									emp_person2 = ""
                              end if
                              rs_emp.close()
							  
	           			%>
							<tr>
								<td class="first"><%=p_cnt%></td>
                                <td><%=rs("pmg_emp_no")%></td>
                                <td><%=rs("pmg_emp_name")%></td>
                                <td><%=emp_person1%>-<%=emp_person2%></td>
                         <%
						      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
                                    de_income_tax = int(Rs_dct("de_income_tax"))
                                    de_wetax = int(Rs_dct("de_wetax"))
									de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
                                    de_year_wetax = int(Rs_dct("de_year_wetax"))
	                             else
                                    de_income_tax = 0
                                    de_wetax = 0
									de_year_incom_tax = 0
                                    de_year_wetax = 0
		                            de_deduct_tot = 0
                              end if
                              Rs_dct.close()
						'퇴직급여
						        a21_income_tax = 0	
								a21_wetax = 0	
								a31_income_tax = de_income_tax + a21_income_tax
							    a31_wetax = de_wetax + a21_wetax
						'연말정산		
								a41_income_tax = 0	
								a41_wetax = 0	
						'조정액		
								a51_income_tax = 0	
								a51_wetax = 0	
                          %>
                                <td class="right" style=" border-left:1px solid #e3e3e3;"><%=formatnumber(de_income_tax,0)%></td>
                                <td class="right"><%=formatnumber(de_wetax,0)%></td>
                                <td class="right"><%=formatnumber(a21_income_tax,0)%></td>
                                <td class="right"><%=formatnumber(a21_wetax,0)%></td>
                                <td class="right"><%=formatnumber(a31_income_tax,0)%></td>
                                <td class="right"><%=formatnumber(a31_wetax,0)%></td>
                                <td class="right"><%=formatnumber(a41_income_tax,0)%></td>
                                <td class="right"><%=formatnumber(a41_wetax,0)%></td>
                                <td class="right"><%=formatnumber(a51_income_tax,0)%></td>
                                <td class="right"><%=formatnumber(a51_wetax,0)%></td>
                                <td><%=rs("pmg_date")%>&nbsp;</td>
                                <td>근로</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						%>
                          	<tr>
                                <th colspan="4" class="first">총계</th>
                                <th class="right"><%=formatnumber(sum_income_tax,0)%></th>
                                <th class="right"><%=formatnumber(sum_wetax,0)%></th>
                                <th class="right"><%=formatnumber(a20_income_tax,0)%></th>
                                <th class="right"><%=formatnumber(a20_wetax,0)%></th>
                                <th class="right"><%=formatnumber(a30_income_tax,0)%></th>
                                <th class="right"><%=formatnumber(a30_wetax,0)%></th>
                                <th class="right"><%=formatnumber(a40_income_tax,0)%></th>
                                <th class="right"><%=formatnumber(a40_wetax,0)%></th>
                                <th class="right"><%=formatnumber(a50_income_tax,0)%></th>
                                <th class="right"><%=formatnumber(a50_wetax,0)%></th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
							</tr>
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
                    <a href="insa_excel_pay_wetax_report.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_wetax_report.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_wetax_report.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_wetax_report.asp?page=<%=i%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_wetax_report.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_wetax_report.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_wetax_print.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>','insa_pay_wetax_pop','scrollbars=yes,width=900,height=550')" class="btnType04">지방소득세계산서/납부서</a>
					</div>                  
                    </td> 
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

