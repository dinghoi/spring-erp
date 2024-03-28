<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

be_pg = "insa_pay_month_saup_list.asp"

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
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
    to_date=request("to_date") 
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
'	view_condi = "전체"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
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

give_date = to_date '지급일

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

if view_condi = "전체" then

      Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
  else
     Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
end if

Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
	pmg_company = rs("pmg_company")
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

'slq = " select  count(*) " & _
'             " from pay_month_give " & _
'			 " where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') " & _
'			 " group by pmg_saupbu "

'Sql = "select count(*) from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
'Set RsCount = Dbconn.Execute (sql)

'tottal_record = cint(RsCount(0)) 'Result.RecordCount

'IF tottal_record mod pgsize = 0 THEN
'	total_page = int(tottal_record / pgsize) 'Result.PageCount
'  ELSE
'	total_page = int((tottal_record / pgsize) + 1)
'END IF

if view_condi = "전체" then
   Sql = " SELECT pmg_saupbu, saup_count, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, " & _
            "   pmg_car_pay, pmg_position_pay, pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay, " & _
			"   pmg_disabled_pay,pmg_give_total, " & _
			"   de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax, " & _
			"   de_year_incom_tax2,de_year_wetax2, " & _
			"   de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_other_amt1,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total " & _
			"   FROM ( " & _
			" select pmg_saupbu,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            "   sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			"   sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			"   sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			"   sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			"   from pay_month_give " & _
			"   where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') group by pmg_saupbu " & _
			"   order by pmg_bonbu,pmg_saupbu " & _
			"   ) a, " & _
			" ( select de_saupbu,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            "   sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			"   sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax," & _
			"   sum(de_year_incom_tax2) as de_year_incom_tax2,sum(de_year_wetax2) as de_year_wetax2,sum(de_sawo_amt) as de_sawo_amt," & _
			"   sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			"   sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			"   sum(de_deduct_total) as de_deduct_total " & _
			"   from pay_month_deduct " & _
			"   where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') group by de_saupbu " & _	
			"   order by de_bonbu,de_saupbu " & _
			"   ) b " & _		
			"  WHERE a.pmg_saupbu = b.de_saupbu " & _
			"  ORDER BY pmg_saupbu ASC " 

   else
 
   Sql = " SELECT pmg_saupbu, saup_count, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, " & _
            "   pmg_car_pay, pmg_position_pay, pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay, " & _
			"   pmg_disabled_pay,pmg_give_total, " & _
			"   de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax, " & _
			"   de_year_incom_tax2,de_year_wetax2, " & _
			"   de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_other_amt1,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total " & _
			"   FROM ( " & _
			" select pmg_saupbu,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            "   sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			"   sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			"   sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			"   sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			"   from pay_month_give " & _
			"   where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') group by pmg_saupbu " & _
			"   order by pmg_company,pmg_bonbu,pmg_saupbu " & _
			"   ) a, " & _
			" ( select de_saupbu,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            "   sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			"   sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax," & _
			"   sum(de_year_incom_tax2) as de_year_incom_tax2,sum(de_year_wetax2) as de_year_wetax2,sum(de_sawo_amt) as de_sawo_amt," & _
			"   sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			"   sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			"   sum(de_deduct_total) as de_deduct_total " & _
			"   from pay_month_deduct " & _
			"   where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_company = '"+view_condi+"') group by de_saupbu " & _	
			"   order by de_company,de_bonbu,de_saupbu " & _
			"   ) b " & _		
			"  WHERE a.pmg_saupbu = b.de_saupbu " & _
			"  ORDER BY pmg_saupbu ASC " 

end if

'base_sql = "SELECT pay_month_give.pmg_company,pay_month_give.pmg_saupbu,COUNT(pay_month_give.pmg_emp_no) AS emp_cnt FROM pay_month_give INNER JOIN pay_month_deduct ON (pay_month_give.pmg_company = pay_month_deduct.de_company) AND (pay_month_give.pmg_id = pay_month_deduct.de_id) AND (pay_month_give.pmg_yymm = pay_month_deduct.de_yymm) AND (pay_month_give.pmg_saupbu = pay_month_deduct.de_saupbu) where (pay_month_give.pmg_yymm = '"+pmg_yymm+"' ) and (pay_month_give.pmg_id = '1') and (pay_month_give.pmg_company = '"+view_condi+"')"

'base_sql = "SELECT pay_month_give.pmg_company,pay_month_give.pmg_saupbu,COUNT(pay_month_give.pmg_emp_no) AS emp_cnt FROM pay_month_give LEFT JOIN pay_month_deduct ON (pay_month_give.pmg_company = pay_month_deduct.de_company) AND (pay_month_give.pmg_saupbu = pay_month_deduct.de_saupbu) where (pay_month_give.pmg_yymm = '"+pmg_yymm+"' ) and (pay_month_give.pmg_id = '1') and (pay_month_give.pmg_company = '"+view_condi+"')"

'group_sql = " group by pay_month_give.pmg_company, pay_month_give.pmg_saupbu order by pay_month_give.pmg_company, pay_month_give.pmg_saupbu"
'sql = base_sql + group_sql

'response.write(sql)

'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC limit "& stpage & "," &pgsize 

Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여 내역서(조직)"

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
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_month_saup_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
                                    <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
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
							<col width="*" >
							<col width="5%" >
                            <col width="9%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="9%" >
							<col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
							<col width="9%" > 
                            <col width="9%" >
                            <col width="5%" >
						</colgroup>
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">조직(사업부)</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">인원</th>
				               <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;">기본급여 및 제수당</th>
                               <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;">공제 및 차인지급액</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">지급액</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">비고</th>
			                </tr>
                            <tr>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">기본급</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">식대</td>  
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">연장근로<br>수당</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">통신비 등</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">지급소계</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">4대보험</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">소득세 등</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">기타공제등</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">예수금계</td>
							</tr>
						</thead>
						<tbody>
					<%
						do until rs.eof
							  pmg_give_tot = cdbl(rs("pmg_give_total"))
							  
							  sub_give_hap = cdbl(rs("pmg_postage_pay")) + cdbl(rs("pmg_re_pay")) + cdbl(rs("pmg_car_pay")) + cdbl(rs("pmg_position_pay")) + cdbl(rs("pmg_custom_pay")) + cdbl(rs("pmg_job_pay")) + cdbl(rs("pmg_job_support")) + cdbl(rs("pmg_jisa_pay")) + cdbl(rs("pmg_long_pay")) + cdbl(rs("pmg_disabled_pay"))
							
							saupbu_name = rs("pmg_saupbu")
							if saupbu_name = "" or saupbu_name = " " or isnull(saupbu_name) then
							    saupbu_name = view_condi
							end if
							  
	           			%>
							<tr>
								<td class="first"><%=saupbu_name%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rs("saup_count")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_base_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_meals_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_overtime_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sub_give_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_give_total"),0)%>&nbsp;</td>
                       <%  
                                  pmg_curr_pay = cdbl(rs("pmg_give_total")) - cdbl(rs("de_deduct_total"))
							  
							      hap_de_insur = cdbl(rs("de_nps_amt")) + cdbl(rs("de_nhis_amt")) + cdbl(rs("de_epi_amt")) + cdbl(rs("de_longcare_amt"))
							      hap_de_tax = cdbl(rs("de_income_tax")) + cdbl(rs("de_wetax")) + cdbl(rs("de_year_incom_tax")) + cdbl(rs("de_year_wetax")) + cdbl(rs("de_year_incom_tax2")) + cdbl(rs("de_year_wetax2"))
							      hap_de_other = cdbl(rs("de_other_amt1")) + cdbl(rs("de_sawo_amt")) + cdbl(rs("de_hyubjo_amt")) + cdbl(rs("de_school_amt")) + cdbl(rs("de_nhis_bla_amt")) + cdbl(rs("de_long_bla_amt"))
								  hap_deduct_tot = hap_de_insur + hap_de_tax + hap_de_other
                       %>
                                <td class="right"><%=formatnumber(hap_de_insur,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(hap_de_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(hap_de_other,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(hap_deduct_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pmg_curr_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
					   <%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_tot - sum_deduct_tot
						
						sum_give_hap = sum_postage_pay + sum_re_pay + sum_car_pay + sum_position_pay + sum_custom_pay + sum_job_pay + sum_job_support + sum_jisa_pay + sum_long_pay + sum_disabled_pay
						
						sum_de_insur =sum_nps_amt +sum_nhis_amt +sum_epi_amt +sum_longcare_amt
						sum_de_tax =sum_income_tax + sum_wetax + sum_year_incom_tax + sum_year_wetax + sum_year_incom_tax2 + sum_year_wetax2
						sum_de_other =sum_other_amt1 +sum_sawo_amt +sum_hyubjo_amt +sum_school_amt +sum_nhis_bla_amt +sum_long_bla_amt
						
						sum_deduct_tot = sum_de_insur + sum_de_tax + sum_de_other
						%>
                          	<tr>
                                <th class="first">총계</th>
                                <th class="right"><%=formatnumber(pay_count,0)%>&nbsp;명</th>
                                <th class="right"><%=formatnumber(sum_base_pay,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_meals_pay,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_overtime_pay,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_postage_pay,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_give_tot,0)%>&nbsp;</th>
                                
                                <th class="right"><%=formatnumber(sum_de_insur,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_de_tax,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_de_other,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_deduct_tot,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</th>
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
                  	<td width="40%">
					<div class="btnleft">
                    <a href="insa_excel_pay_month_saup.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
                    <a href="insa_pay_month_saup_excel2.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>" class="btnType04">엑셀다운(팀)</a>
                    <a href="insa_pay_month_saup_excel3.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>" class="btnType04">엑셀다운(사업부)</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_month_saup_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_month_saup_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_month_saup_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_month_saup_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_month_saup_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_month_saup_print2.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>','insa_pay_month_saup_list_pop','scrollbars=yes,width=1250,height=700')" class="btnType04">급여내역서(팀) 출력</a>
                    <a href="#" onClick="pop_Window('insa_pay_month_saup_print3.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>','insa_pay_month_saup_list_pop','scrollbars=yes,width=1250,height=700')" class="btnType04">급여내역서(사업부) 출력</a>
					</div>                  
                    </td>                    
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

