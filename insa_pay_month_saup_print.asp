<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sch_tab(10,10)

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))
to_yyyy = mid(cstr(to_date),1,4)
to_mm = mid(cstr(to_date),6,2)
to_dd = mid(cstr(to_date),9,2)

give_date = to_date '지급일

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
main_title = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여 내역서"

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
    sum_other_amt1 = 0
    sum_sawo_amt = 0
    sum_hyubjo_amt = 0
    sum_school_amt = 0
    sum_nhis_bla_amt = 0
    sum_long_bla_amt = 0
	sum_deduct_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")

DbConn.Open dbconnect

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
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

sql = "select pmg_saupbu,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            " sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			" sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			" sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			" sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			" from pay_month_give " & _
			" where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') group by pmg_saupbu " & _
			" order by pmg_company,pmg_bonbu"
Rs.Open Sql, Dbconn, 1

sql = "select de_saupbu,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            " sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			" sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax,sum(de_sawo_amt) as de_sawo_amt," & _
			" sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			" sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			" sum(de_deduct_total) as de_deduct_total " & _
			" from pay_month_deduct " & _
			" where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_company = '"+view_condi+"') group by de_saupbu " & _
			" order by de_company,de_bonbu"

Set Rs_dct = DbConn.Execute(SQL)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        <title>인사급여 시스템</title>
        <script src="/java/common.js" type="text/javascript"></script>
        <script type="text/javascript">
	function printWindow(){
//		viewOff("button");   
		factory.printing.header = ""; //머리말 정의
		factory.printing.footer = ""; //꼬리말 정의
		factory.printing.portrait = true; //출력방향 설정: true - 가로, false - 세로
		factory.printing.leftMargin = 13; //외쪽 여백 설정
		factory.printing.topMargin = 25; //윗쪽 여백 설정
		factory.printing.rightMargin = 13; //오른쯕 여백 설정
		factory.printing.bottomMargin = 15; //바닦 여백 설정
//		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
//		factory.printing.printer = ""; //프린터 할 프린터 이름
//		factory.printing.paperSize = "A4"; //용지선택
//		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
//		factory.printing.collate = true; //순서대로 출력하기
//		factory.printing.copies = "1"; //인쇄할 매수
//		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
//		factory.printing.Printer(true); //출력하기
		factory.printing.Preview(); //윈도우를 통해서 출력
		factory.printing.Print(false); //윈도우를 통해서 출력
	}
	function printW() {
        window.print();
    }
	function goBefore () {
		history.back() ;
	}
	
</script>
<title>월 급여지급대장</title>
<style type="text/css">
<!--
    	.style10C {font-size: 10px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style10BC {font-size: 10px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style14BC {font-size: 14px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style24BC {font-size: 24px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style32BC {font-size: 32px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
-->
</style>
<style media="print"> 
.noprint     { display: none }
</style>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<div class="noprint">
<p><a href="#" onClick="printWindow()"><img src="image/printer.jpg" width="39" height="36" border="0" alt="출력하기" /></a></p>
</div>
<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
</object>    
   
<table width="1030" cellpadding="0" cellspacing="0">
  <tr>
     <td colspan="3" align="center" class="style24BC"><%=main_title%></td>
  </tr>
  <tr>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
  </tr>
  <tr>
	 <td width="33%" height="30" align="left"><span class="style14BC"><%=view_condi%>&nbsp;&nbsp;(부서별)</span></td>
	 <td width="*" height="30" align="center"><span class="style14BC">[귀속:<%=curr_yyyy%>년<%=curr_mm%>]&nbsp;[지급:<%=to_yyyy%>년<%=to_mm%>월<%=curr_yyyy%>일]</span></td>
	 <td width="33%" height="30" align="left"><span class="style14BC">&nbsp;&nbsp;</span></td>
  </tr>  
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">조직(사업부)</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">인원</span></td>
    <td width="10%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">기본급</span></td>
    <td width="10%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">식대</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">통신비</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">소급급여</span></td>
    <td width="9%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">연장근로<br>수당</span></td>
    <td width="9%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">주차지원금</span></td>
    <td width="8%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">직책수당</span></td>
    <td width="7%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">고객관리<br>수당</span></td>
    <td width="7%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">직무<br>보조비</span></td>
    <td width="7%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">업무<br>장려비</span></td>
    <td width="7%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">본지사<br>근무비</span></td>
    <td width="7%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">근속수당</span></td>
    <td width="9%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">장애인수당</span></td>
    <td width="10%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><strong class="style12C">지급합계</strong></td>
    <td width="7%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">국민연금</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">건강보험</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">고용보험</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">장기요양<br>보험료</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">소득세</span></td>
    <td width="8%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">지방소득세</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">기타공제</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">사우회<br>회비</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">학자금<br>상환</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">협조비</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">건강보험<br>료정산</span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style10C">장기요양<br>보험료정산</span></td>
    <td width="10%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><strong class="style12C">공제합계</strong></td>
    <td width="10%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><strong class="style12C">차인지급액</strong></td>
  </tr>

 <%
	do until rs.eof
	  pmg_give_tot = cdbl(rs("pmg_give_total"))
							  
	  sql = "select de_saupbu,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            " sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			" sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax,sum(de_sawo_amt) as de_sawo_amt," & _
			" sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			" sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			" sum(de_deduct_total) as de_deduct_total " & _
			" from pay_month_deduct " & _
			" where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_company = '"+view_condi+"') group by de_saupbu " & _
			" order by de_company,de_bonbu"

     Set Rs_dct = DbConn.Execute(SQL)
							  
							  sub_give_hap = cdbl(rs("pmg_postage_pay")) + cdbl(rs("pmg_re_pay")) + cdbl(rs("pmg_car_pay")) + cdbl(rs("pmg_position_pay")) + cdbl(rs("pmg_custom_pay")) + cdbl(rs("pmg_job_pay")) + cdbl(rs("pmg_job_support")) + cdbl(rs("pmg_jisa_pay")) + cdbl(rs("pmg_long_pay")) + cdbl(rs("pmg_disabled_pay"))

  saupbu_name = rs("pmg_saupbu")
  if saupbu_name = "" or saupbu_name = " " or isnull(saupbu_name) then
	    saupbu_name = view_condi
  end if
							  
 %>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=saupbu_name%></span></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=rs("saup_count")%></span></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_base_pay"),0)%></span></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_meals_pay"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_postage_pay"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_re_pay"),0)%></span></td>
    <td width="9%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_overtime_pay"),0)%></span></td>
    <td width="9%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_car_pay"),0)%></span></td>
    <td width="8%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_position_pay"),0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_custom_pay"),0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_job_pay"),0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_job_support"),0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_jisa_pay"),0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_long_pay"),0)%></span></td>
    <td width="9%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(rs("pmg_disabled_pay"),0)%></span></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><strong class="style12C"><%=formatnumber(rs("pmg_give_total"),0)%></strong></td>

  <%  do until Rs_dct.eof
       if rs("pmg_saupbu") = Rs_dct("de_saupbu") then 
          pmg_curr_pay = cdbl(rs("pmg_give_total")) - cdbl(Rs_dct("de_deduct_total"))
					  
 	      hap_de_insur = cdbl(Rs_dct("de_nps_amt")) + cdbl(Rs_dct("de_nhis_amt")) + cdbl(Rs_dct("de_epi_amt")) + cdbl(Rs_dct("de_longcare_amt"))
		  hap_de_tax = cdbl(Rs_dct("de_income_tax")) + cdbl(Rs_dct("de_wetax")) + cdbl(Rs_dct("de_year_incom_tax")) + cdbl(Rs_dct("de_year_wetax"))
		  hap_de_other = cdbl(Rs_dct("de_other_amt1")) + cdbl(Rs_dct("de_sawo_amt")) + cdbl(Rs_dct("de_hyubjo_amt")) + cdbl(Rs_dct("de_school_amt")) + cdbl(Rs_dct("de_nhis_bla_amt")) + cdbl(Rs_dct("de_long_bla_amt"))
  %>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_nps_amt"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_nhis_amt"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_epi_amt"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_longcare_amt"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_income_tax"),0)%></span></td>
    <td width="8%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_wetax"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_other_amt1"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_sawo_amt"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_school_amt"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_hyubjo_amt"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_nhis_bla_amt"),0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(Rs_dct("de_long_bla_amt"),0)%></span></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><strong class="style12C"><%=formatnumber(Rs_dct("de_deduct_total"),0)%></strong></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><strong class="style12C"><%=formatnumber(pmg_curr_pay,0)%></strong></td>
 <%     end if 
     Rs_dct.movenext()
   loop 
   Rs_dct.close()
 %>
  </tr>                              

 <%
    	rs.movenext()
	loop
	rs.close()
		
	sum_curr_pay = sum_give_tot - sum_deduct_tot
						
	sum_give_hap = sum_postage_pay + sum_re_pay + sum_car_pay + sum_position_pay + sum_custom_pay + sum_job_pay +       sum_job_support + sum_jisa_pay + sum_long_pay + sum_disabled_pay
	sum_de_insur =sum_nps_amt +sum_nhis_amt +sum_epi_amt +sum_longcare_amt
	sum_de_tax =sum_income_tax +sum_wetax
	sum_de_other =sum_other_amt1 +sum_sawo_amt +sum_hyubjo_amt +sum_school_amt +sum_nhis_bla_amt +sum_long_bla_amt
						
 %>  

  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><strong class="style12C">총계</strong></td>
    <td width="6%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(pay_count,0)%>&nbsp;명</span></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_base_pay,0)%></span></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_meals_pay,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_postage_pay,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_re_pay,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_overtime_pay,0)%></span></td>
    <td width="9%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_car_pay,0)%></span></td>
    <td width="9%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_position_pay,0)%></span></td>
    <td width="8%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_custom_pay,0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_job_pay,0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_job_support,0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_jisa_pay,0)%></span></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_long_pay,0)%></span></td>
    <td width="9%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_disabled_pay,0)%></span></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><strong class="style12C"><%=formatnumber(sum_give_tot,0)%></strong></td>
    <td width="7%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_nps_amt,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_nhis_amt,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_epi_amt,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_longcare_amt,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_income_tax,0)%></span></td>
    <td width="8%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_wetax,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_other_amt1,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_sawo_amt,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_school_amt,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_hyubjo_amt,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_nhis_bla_amt,0)%></span></td>
    <td width="6%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><span class="style12C"><%=formatnumber(sum_long_bla_amt,0)%></span></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><strong class="style12C"><%=formatnumber(sum_deduct_tot,0)%></strong></td>
    <td width="10%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><strong class="style12C"><%=formatnumber(sum_curr_pay,0)%></strong></td>
  </tr>       
</table>
</p>	

</body>
</html>
