<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim year_tab(3,2)

be_pg = "insa_pay_pay_report_person.asp"

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

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
	
	sum_base_pay = 0
	sum_meals_pay = 0
	sum_postage_pay = 0
	sum_re_pay = 0
	sum_overtime_pay = 0
	sum_car_pay = 0
	sum_position_pay = 0
	sum_custom_pay = 0
	sum_job_pay = 0
	sum_job_support = 0
	sum_jisa_pay = 0
	sum_long_pay = 0
	sum_disabled_pay = 0
	sum_family_pay = 0
	sum_school_pay = 0
	sum_qual_pay = 0
	sum_other_pay1 = 0
	sum_other_pay2 = 0
	sum_other_pay3 = 0
	sum_tax_yes = 0
	sum_tax_no = 0
	sum_tax_reduced = 0
	sum_give_tot = 0
    sum_nps_amt = 0
    sum_nhis_amt = 0
    sum_epi_amt = 0
    sum_longcare_amt = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_year_incom_tax2 = 0
    sum_year_wetax2 = 0
    sum_other_amt1 = 0
    sum_sawo_amt = 0
    sum_hyubjo_amt = 0
    sum_school_amt = 0
    sum_nhis_bla_amt = 0
    sum_long_bla_amt = 0
	sum_deduct_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0	
	
end if

inc_yyyyf = inc_yyyy + "01"
inc_yyyyl = inc_yyyy + "12"

give_date = to_date '지급일

' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

pgsize = 12 ' 화면 한 페이지 

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

if condi <> "" then
      if owner_view = "C" then 
             Sql = "select count(*) from pay_month_give where (pmg_yymm >= '"+inc_yyyyf+"' and pmg_yymm <= '"+inc_yyyyl+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"') and (pmg_emp_name like '%"+condi+"%')"
         else
             Sql = "select count(*) from pay_month_give where (pmg_yymm >= '"+inc_yyyyf+"' and pmg_yymm <= '"+inc_yyyyl+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"') and (pmg_emp_no = '"+condi+"')"
	  end if

Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF
end if

if condi <> "" then
      if owner_view = "C" then 
             Sql = "select * from pay_month_give where (pmg_yymm >= '"+inc_yyyyf+"' and pmg_yymm <= '"+inc_yyyyl+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"') and (pmg_emp_name like '%"+condi+"%') ORDER BY pmg_emp_no,pmg_yymm ASC"
		 else	 
			 Sql = "select * from pay_month_give where (pmg_yymm >= '"+inc_yyyyf+"' and pmg_yymm <= '"+inc_yyyyl+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"') and (pmg_emp_no = '"+condi+"') ORDER BY pmg_emp_no,pmg_yymm ASC"
	  end if
 
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
	pmg_yymm = rs("pmg_yymm")
    pmg_give_tot = rs("pmg_give_total")
    pay_count = pay_count + 1
				  
    sum_base_pay = sum_base_pay + int(rs("pmg_base_pay"))
    sum_meals_pay = sum_meals_pay + int(rs("pmg_meals_pay"))
    sum_postage_pay = sum_postage_pay + int(rs("pmg_postage_pay"))
    sum_re_pay = sum_re_pay + int(rs("pmg_re_pay"))
    sum_overtime_pay = sum_overtime_pay + int(rs("pmg_overtime_pay"))
    sum_car_pay = sum_car_pay + int(rs("pmg_car_pay"))
    sum_position_pay = sum_position_pay + int(rs("pmg_position_pay"))
    sum_custom_pay = sum_custom_pay + int(rs("pmg_custom_pay"))
    sum_job_pay = sum_job_pay + int(rs("pmg_job_pay"))
    sum_job_support = sum_job_support + int(rs("pmg_job_support"))
    sum_jisa_pay = sum_jisa_pay + int(rs("pmg_jisa_pay"))
    sum_long_pay = sum_long_pay + int(rs("pmg_long_pay"))
    sum_disabled_pay = sum_disabled_pay + int(rs("pmg_disabled_pay"))
    sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
    if not Rs_dct.eof then
	    	de_nps_amt = int(Rs_dct("de_nps_amt"))
            de_nhis_amt = int(Rs_dct("de_nhis_amt"))
            de_epi_amt = int(Rs_dct("de_epi_amt"))
		    de_longcare_amt = int(Rs_dct("de_longcare_amt"))
            de_income_tax = int(Rs_dct("de_income_tax"))
            de_wetax = int(Rs_dct("de_wetax"))
			de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
            de_year_wetax = int(Rs_dct("de_year_wetax"))
			de_year_incom_tax2 = int(Rs_dct("de_year_incom_tax2"))
            de_year_wetax2 = int(Rs_dct("de_year_wetax2"))
            de_other_amt1 = int(Rs_dct("de_other_amt1"))
            de_sawo_amt = int(Rs_dct("de_sawo_amt"))
            de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
            de_school_amt = int(Rs_dct("de_school_amt"))
            de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
            de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
		    de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	     else
			de_nps_amt = 0
            de_nhis_amt = 0
            de_epi_amt = 0
		    de_longcare_amt = 0
            de_income_tax = 0
            de_wetax = 0
			de_year_incom_tax = 0
            de_year_wetax = 0
			de_year_incom_tax2 = 0
            de_year_wetax2 = 0
            de_other_amt1 = 0
            de_sawo_amt = 0
            de_hyubjo_amt = 0
            de_school_amt = 0
            de_nhis_bla_amt = 0
            de_long_bla_amt = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()
     sum_nps_amt = sum_nps_amt + de_nps_amt
     sum_nhis_amt = sum_nhis_amt + de_nhis_amt
     sum_epi_amt = sum_epi_amt + de_epi_amt
	 sum_longcare_amt = sum_longcare_amt + de_longcare_amt
     sum_income_tax = sum_income_tax + de_income_tax
     sum_wetax = sum_wetax + de_wetax
	 sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
     sum_year_wetax = sum_year_wetax + de_year_wetax
	 sum_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
     sum_year_wetax2 = sum_year_wetax2 + de_year_wetax2
     sum_other_amt1 = sum_other_amt1 + de_other_amt1
     sum_sawo_amt = sum_sawo_amt + de_sawo_amt
     sum_hyubjo_amt = sum_hyubjo_amt + de_hyubjo_amt
     sum_school_amt = sum_school_amt + de_school_amt
     sum_nhis_bla_amt = sum_nhis_bla_amt + de_nhis_bla_amt
     sum_long_bla_amt = sum_long_bla_amt + de_long_bla_amt
	 sum_deduct_tot = sum_deduct_tot + de_deduct_tot

	rs.movenext()
loop
rs.close()
end if	  
if condi <> "" then
      if owner_view = "C" then 
             Sql = "select * from pay_month_give where (pmg_yymm >= '"+inc_yyyyf+"' and pmg_yymm <= '"+inc_yyyyl+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"') and (pmg_emp_name like '%"+condi+"%') ORDER BY pmg_emp_no,pmg_yymm ASC limit "& stpage & "," &pgsize 
		 else	 
			 Sql = "select * from pay_month_give where (pmg_yymm >= '"+inc_yyyyf+"' and pmg_yymm <= '"+inc_yyyyl+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"') and (pmg_emp_no = '"+condi+"') ORDER BY pmg_emp_no,pmg_yymm ASC limit "& stpage & "," &pgsize 
      end if

Rs.Open Sql, Dbconn, 1
end if

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(inc_yyyy) + "년 " + " 급여현황(개인)"

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
				return "7 1";
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_pay_report_person.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
							<col width="8%" >
                            <col width="4%" >
                            <col width="5%" >
                            <col width="*" >
                            <col width="8%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="8%" >
							<col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
							<col width="8%" > 
                            <col width="8%" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col">성명</th>
                               <th rowspan="2" scope="col">지급월</th>
                               <th rowspan="2" scope="col">직급</th>
                               <th rowspan="2" scope="col">소속</th>
				               <th colspan="5" scope="col" style="background:#FFFFE6;">기본급여 및 제수당</th>
                               <th colspan="4" scope="col" style="background:#E0FFFF;">공제 및 차인지급액</th>
                               <th rowspan="2" scope="col">지급액</th>
                               <th rowspan="2" scope="col">조회</th>
			                </tr>
                            <tr>
								<td scope="col" style=" border-left:1px solid #e3e3e3; background:#f8f8f8">기본급</td>
								<td scope="col" style="background:#f8f8f8;">식대</td>  
								<td scope="col" style="background:#f8f8f8;">연장근로<br>수당</td>
                                <td scope="col" style="background:#f8f8f8;">통신비 등</td>
                                <td scope="col" style="background:#f8f8f8;">지급소계</td>
								<td scope="col" style="background:#f8f8f8;">4대보험</td>
                                <td scope="col" style="background:#f8f8f8;">소득세 등</td>
								<td scope="col" style="background:#f8f8f8;">기타공제등</td>
                                <td scope="col" style="background:#f8f8f8;">예수금계</td>
							</tr>
						</thead>
						<tbody>
						<% if condi <> "" then
						     do until rs.eof
							  emp_no = rs("pmg_emp_no")
							  pmg_yymm = rs("pmg_yymm")
							  pmg_company = rs("pmg_company")
							   
							  pmg_give_tot = rs("pmg_give_total")

							  sub_give_hap = int(rs("pmg_postage_pay")) + int(rs("pmg_re_pay")) + int(rs("pmg_car_pay")) + int(rs("pmg_position_pay")) + int(rs("pmg_custom_pay")) + int(rs("pmg_job_pay")) + int(rs("pmg_job_support")) + int(rs("pmg_jisa_pay")) + int(rs("pmg_long_pay")) + int(rs("pmg_disabled_pay"))
							  
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
		                      if not rs_emp.eof then
		                    		emp_in_date = rs_emp("emp_in_date")
	                             else
	                    			emp_in_date = ""
                              end if
                              rs_emp.close()
							  
	           			%>
							<tr>
								<td class="first"><%=rs("pmg_emp_name")%>(<%=rs("pmg_emp_no")%>)</td>
                                <td style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_yymm")%></td>
                                <td style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_grade")%></td>
                                <td style=" border-left:1px solid #e3e3e3;"><%=rs("pmg_org_name")%></td>
                                <td class="right"><%=formatnumber(rs("pmg_base_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_meals_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_overtime_pay"),0)%></td>
                                <td class="right"><%=formatnumber(sub_give_hap,0)%></td>
                                <td class="right"><%=formatnumber(rs("pmg_give_total"),0)%></td>
                         <%
						      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+pmg_company+"')"
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
									de_nps_amt = int(Rs_dct("de_nps_amt"))
                                    de_nhis_amt = int(Rs_dct("de_nhis_amt"))
                                    de_epi_amt = int(Rs_dct("de_epi_amt"))
		                            de_longcare_amt = int(Rs_dct("de_longcare_amt"))
                                    de_income_tax = int(Rs_dct("de_income_tax"))
                                    de_wetax = int(Rs_dct("de_wetax"))
									de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
                                    de_year_wetax = int(Rs_dct("de_year_wetax"))
									de_year_incom_tax2 = int(Rs_dct("de_year_incom_tax2"))
                                    de_year_wetax2 = int(Rs_dct("de_year_wetax2"))
                                    de_other_amt1 = int(Rs_dct("de_other_amt1"))
                                    de_sawo_amt = int(Rs_dct("de_sawo_amt"))
                                    de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
                                    de_school_amt = int(Rs_dct("de_school_amt"))
                                    de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
                                    de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
		                            de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	                             else
									de_nps_amt = 0
                                    de_nhis_amt = 0
                                    de_epi_amt = 0
		                            de_longcare_amt = 0
                                    de_income_tax = 0
                                    de_wetax = 0
									de_year_incom_tax = 0
                                    de_year_wetax = 0
									de_year_incom_tax2 = 0
                                    de_year_wetax2 = 0
                                    de_other_amt1 = 0
                                    de_sawo_amt = 0
                                    de_hyubjo_amt = 0
                                    de_school_amt = 0
                                    de_nhis_bla_amt = 0
                                    de_long_bla_amt = 0
		                            de_deduct_tot = 0
                              end if
                              Rs_dct.close()
							  pmg_curr_pay = pmg_give_tot - de_deduct_tot
							  
							  hap_de_insur = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt
							  hap_de_tax = de_income_tax + de_wetax + de_year_incom_tax + de_year_wetax + de_year_incom_tax2 + de_year_wetax2
							  hap_de_other = de_other_amt1 + de_sawo_amt + de_hyubjo_amt + de_school_amt + de_nhis_bla_amt + de_long_bla_amt
							  
                          %>
                                <td class="right"><%=formatnumber(hap_de_insur,0)%></td>
                                <td class="right"><%=formatnumber(hap_de_tax,0)%></td>
                                <td class="right"><%=formatnumber(hap_de_other,0)%></td>
                                <td class="right"><%=formatnumber(de_deduct_tot,0)%></td>
                                <td class="right"><%=formatnumber(pmg_curr_pay,0)%></td>
                                <td class="right"><a href="#" onClick="pop_Window('insa_pay_person_view.asp?emp_no=<%=rs("pmg_emp_no")%>&emp_name=<%=rs("pmg_emp_name")%>&pmg_yymm=<%=pmg_yymm%>&pmg_date=<%=give_date%>&pmg_company=<%=rs("pmg_company")%>&pmg_org_code=<%=rs("pmg_org_code")%>&pmg_org_name=<%=rs("pmg_org_name")%>&pmg_grade=<%=rs("pmg_grade")%>&pmg_position=<%=rs("pmg_position")%>','insa_pay_person_pop','scrollbars=yes,width=750,height=700')">상세</a></td>
                                
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						end if
						sum_curr_pay = sum_give_tot - sum_deduct_tot
						
						sum_give_hap = sum_postage_pay + sum_re_pay + sum_car_pay + sum_position_pay + sum_custom_pay + sum_job_pay + sum_job_support + sum_jisa_pay + sum_long_pay + sum_disabled_pay
						sum_de_insur =sum_nps_amt +sum_nhis_amt +sum_epi_amt +sum_longcare_amt
						sum_de_tax =sum_income_tax + sum_wetax + sum_year_incom_tax + sum_year_wetax + sum_year_incom_tax2 + sum_year_wetax2
						sum_de_other =sum_other_amt1 +sum_sawo_amt +sum_hyubjo_amt +sum_school_amt +sum_nhis_bla_amt +sum_long_bla_amt
						
						%>
                          	<tr>
                                <th colspan="2" class="first">총계</th>
                                <th colspan="2" class="right">&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_base_pay,0)%></th>
                                <th class="right"><%=formatnumber(sum_meals_pay,0)%></th>
                                <th class="right"><%=formatnumber(sum_overtime_pay,0)%></th>
                                <th class="right"><%=formatnumber(sum_give_hap,0)%></th>
                                <th class="right"><%=formatnumber(sum_give_tot,0)%></th>
                                <th class="right"><%=formatnumber(sum_de_insur,0)%></th>
                                <th class="right"><%=formatnumber(sum_de_tax,0)%></th>
                                <th class="right"><%=formatnumber(sum_de_other,0)%></th>
                                <th class="right"><%=formatnumber(sum_deduct_tot,0)%></th>
                                <th class="right"><%=formatnumber(sum_curr_pay,0)%></th>
                                <th class="right">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/12)*12) + 1
                intend = intstart + 11
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_pay_report_person.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_pay_report_person.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_pay_report_person.asp?page=<%=i%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_pay_report_person.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_pay_report_person.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&inc_yyyy=<%=inc_yyyy%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
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

