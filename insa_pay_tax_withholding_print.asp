<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sch_tab(10,10)

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
in_tax_id = request("in_tax_id") 

curr_date = datevalue(mid(cstr(now()),1,10))
to_yyyy = mid(cstr(to_date),1,4)
to_mm = mid(cstr(to_date),6,2)
to_dd = mid(cstr(to_date),9,2)

give_date = to_date '지급일

tax_man_name = ""

if view_condi = "케이원정보통신" then
      company_name = "(주)" + "케이원정보통신"
	  owner_name = "김승일"
	  addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	  trade_no = "107-81-54150"
	  tel_no = "02) 853-5250"
	  e_mail = "js10547@k-won.co.kr"
   elseif view_condi = "휴디스" then
              company_name = "(주)" + "휴디스"
			  owner_name = "김한종"
	          addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	          trade_no = "107-81-54150"
	          tel_no = "02) 853-5250"
	          e_mail = "js10547@k-won.co.kr"
		  elseif view_condi = "케이네트웍스" then
                     company_name = "케이네트웍스" + "(주)"
					 owner_name = "이중원"
	                 addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                 trade_no = "107-81-54150"
	                 tel_no = "02) 853-5250"
	                 e_mail = "js10547@k-won.co.kr"
				 elseif view_condi = "에스유에이치" then
                        company_name = "(주)" + "에스유에이치"	
						owner_name = "박미애"
	                    addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                    trade_no = "119-86-78709"
	                    tel_no = "02) 6116-8248"
	                    e_mail = "pshwork27@k-won.co.kr"
end if 

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
main_title = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여대장"

	sum_give_tot = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_special_tax = 0
	sum_deduct_tot = 0
	pay_count = 0	
	sum_curr_pay = 0
	
	a02_give_tot = 0
    a02_income_tax = 0
    a02_wetax = 0
	a02_count = 0	
	
	a03_give_tot = 0
    a03_income_tax = 0
    a03_wetax = 0
	a03_count = 0	
	
	a04_give_tot = 0
    a04_income_tax = 0
    a04_wetax = 0
	a04_count = 0	
	
	a10_give_tot = 0
    a10_income_tax = 0
    a10_wetax = 0
	a10_count = 0	
	
	a21_give_tot = 0
    a21_income_tax = 0
    a21_wetax = 0
	a21_count = 0	
	
	a22_give_tot = 0
    a22_income_tax = 0
    a22_wetax = 0
	a22_count = 0	
	
	a20_give_tot = 0
    a20_income_tax = 0
    a20_wetax = 0
	a20_count = 0	
	
	sum_alba_give_total = 0
    sum_tax_amt1 = 0
    sum_tax_amt2 = 0
	sum_deduct_tot = 0
	
	a32_give_tot = 0
    a32_income_tax = 0
    a32_wetax = 0
	a32_count = 0	
	
	a30_give_tot = 0
    a30_income_tax = 0
    a30_wetax = 0
	a30_count = 0
	
	tot_give_tot = 0
    tot_income_tax = 0
    tot_wetax = 0
	tot_year_incom_tax = 0
    tot_year_wetax = 0
	tot_special_tax = 0
	tot_deduct_tot = 0
	tot_pay_count = 0	
	tot_curr_pay = 0		

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

'근로소득
Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
    pmg_give_tot = rs("pmg_give_total")
    pay_count = pay_count + 1
				  
    sub_give_hap = int(rs("pmg_postage_pay")) + int(rs("pmg_re_pay")) + int(rs("pmg_car_pay")) + int(rs("pmg_position_pay")) + int(rs("pmg_custom_pay")) + int(rs("pmg_job_pay")) + int(rs("pmg_job_support")) + int(rs("pmg_jisa_pay")) + int(rs("pmg_long_pay")) + int(rs("pmg_disabled_pay"))
	
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

a10_give_tot = sum_give_tot + a02_give_tot + a03_give_tot + a03_give_tot 
a10_income_tax = sum_income_tax + a02_income_tax + a03_income_tax + a04_income_tax
a10_wetax = sum_wetax + a02_wetax + a03_wetax + a04_wetax
a10_count = pay_count + a02_count + a03_count + a04_count

'퇴직소득
a20_give_tot = a21_give_tot + a22_give_tot
a20_income_tax = a21_income_tax + a22_income_tax
a20_wetax = a21_wetax + a22_wetax
a20_count = a21_count + a22_count

'사업소득
Sql = "select * from pay_alba_cost where (rever_yymm = '"+pmg_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    alba_count = alba_count + 1
				  
    sum_alba_give_total = sum_alba_give_total + int(rs("alba_give_total"))
    sum_tax_amt1 = sum_tax_amt1 + int(rs("tax_amt1"))
    sum_tax_amt2 = sum_tax_amt2 + int(rs("tax_amt2"))
	sum_deduct_tot = sum_deduct_tot + (int(rs("tax_amt1")) + int(rs("tax_amt2")) + int(rs("de_other")))
	
	rs.movenext()
loop
rs.close()

a30_give_tot = sum_alba_give_total + a32_give_tot
a30_income_tax = sum_tax_amt1 + a32_income_tax
a30_wetax = sum_tax_amt2 + a32_wetax
a30_count = alba_count + a32_count

'총계
tot_give_tot = a10_give_tot + a20_give_tot + a30_give_tot
tot_income_tax = a10_income_tax + a20_income_tax + a30_income_tax
tot_wetax = a10_wetax + a20_wetax + a30_wetax
tot_pay_count = a10_count + a20_count + a30_count

if in_tax_id = "1" then 
   tax_id_name = "정기신고" 
   elseif in_tax_id = "2" then 
          tax_id_name = "분기" 
          elseif in_tax_id = "3" then 
		         tax_id_name = "연말" 
end if

title_line = " ■소득세법 시행규칙[별지 제21호서식]<개정2014.3.14> "

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        <title>급여관리 시스템</title>
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
<title>원천징수이행상황신고서</title>
<style type="text/css">
<!--
    	.style10C {font-size: 10px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
		.style10BC {font-size: 10px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style12L {font-size: 12px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style14BC {font-size: 14px; font-weight: bold; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style24BC {font-size: 24px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style32BC {font-size: 32px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
		.style10L {font-size: 8px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
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
    <td width="33%" height="20" align="left"><span class="style12L"><%=title_line%></span></td>
    <td width="*" height="20" align="center"><span class="style12L">&nbsp;&nbsp;</span></td>
    <td width="33%" height="20" align="right"><span class="style12L">[전자신고제출분]</span></td>
  </tr>  
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="6" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">① 신고구분</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:0px solid #ffffff;"><strong class="style14BC">■&nbsp;&nbsp; 원천징수이행상황신고서<br>□&nbsp;&nbsp; 원천징수세액환급신청서</strong></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">② 귀속년월</span></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=curr_yyyy%>년<%=curr_mm%>월</span></td>
  </tr>
  <tr>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">매월</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">반기</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">수정</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">연말</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">소득처분</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">환급신청</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">③ 지급연월</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=curr_yyyy%>년<%=curr_mm%>월</span></td>
  </tr>
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td rowspan="4" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">원천<br>징수<br>의무자</span></td>
    <td rowspan="2" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">법인명(상호)</span></td>
    <td rowspan="2" width="20%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=company_name%></span></td>
    <td rowspan="2" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">대표자<br>(성명)</span></td>
    <td rowspan="2" width="20%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=owner_name%></span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">일괄납부 여부</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">여 , 부</span></td>
  </tr>
  <tr>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">사업자단위과세여부</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">여 , 부</span></td>
  </tr>
  <tr>
    <td rowspan="2" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">사업자(주민)<br>등록번호</span></td>
    <td rowspan="2" width="20%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=trade_no%></span></td>
    <td rowspan="2" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">사업장<br>소재지</span></td>
    <td rowspan="2" width="20%" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=addr_name%></span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">전화번호</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=tel_no%></span></td>
  </tr>
  <tr>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">전자우편주소</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=e_mail%></span></td>
  </tr>
  <tr>
    <td colspan="7" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><strong class="style12C">1. 원천징수명세 및 납부세액 (단위: 원)</strong></td>
  </tr>
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td rowspan="3" colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">사업소득자구분</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">코드</span></td>
    <td colspan="5" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">원천징수명세</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑨ 당월조정<br>환급세액</span></td>
    <td colspan="5" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">납부세액</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">소득지급<br>(과세미달,일부비과세포함)</span></td>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">징수세액</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑩ 소득세등<br>(가산세포함)</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑪ 농어촌<br>특별세</span></td>
  </tr>
  <tr>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">④ 인원</span></td>
    <td width="12%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑤ 총지급액</span></td>
    <td width="12%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑥ 소득세등</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑦ 농어촌특별세</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑧ 가산세</span></td>
  </tr>
  <tr>
    <td rowspan="22" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">개<br>인<br>∧<br>거<br>주<br>자<br>·<br>·<br>비<br>거<br>주<br>자<br>∨</span></td>
    <td rowspan="5" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">근<br>
    로<br>소<br>득</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">간이세액</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A01</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(pay_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">중도퇴사</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A02</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a02_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a02_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a02_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">일용근로</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A03</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">연말정산</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A04</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a04_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">가감계</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A10</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a10_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a10_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a10_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a10_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td rowspan="3" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">퇴<br>직<br>소<br>득</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">연금계좌</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A21</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a21_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a21_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a21_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">그외</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A22</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">가감계</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A20</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td rowspan="3" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">사<br>업<br>소<br>득</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">매월징수</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A25</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(alba_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_alba_give_total,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_tax_amt1,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">연말정산</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A26</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a32_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a32_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a32_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">가감계</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A30</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a30_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a30_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a30_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a30_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td rowspan="3" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">기<br>타<br>소<br>득</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">연금계좌</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A41</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">그외</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A42</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">가감계</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A40</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td rowspan="4" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">연<br>금<br>소<br>득</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">연금계좌</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A48</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">공적연금(매월)</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A45</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">연말정산</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A46</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">가감계</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A47</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">이자소득</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A50</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">배당소득</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A60</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style10C">저축해지추징세액등</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A69</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">비거주자양도소득</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A70</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">법인</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">내외국법인원천</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A80</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">수정신고(세액)</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A90</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">총합계</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A99</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(tot_pay_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(tot_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(tot_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(tot_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="12" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><strong class="style12C">2. 환급세액 조정 (단위: 원)</strong></td>
  </tr>
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">전월 미환급 세액의 계산</span></td>
    <td colspan="4" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">당월 발생 환급세액</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">?조정대상<br>환급세액<br>(⑭+⑮+<br>?+?)</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? 당월조정<br>환급세액</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? 차월이월<br>환급세액<br>(?-?)</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? 환급신청액</span></td>
  </tr>
  <tr>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑫ 전월<br>미환급세액</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑬ 기환급<br>신청세액</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑭ 차감잔액<br>(⑫-⑬)</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">⑮ 일반환급</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? 신탁재산<br>(금융회사등)</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? 그밖의환급세액</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">금융회사등</span></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">합병 등</span></td>
  </tr>
  <tr>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
</table>

<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td rowspan="11" width="70%" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;&nbsp;&nbsp;&nbsp;원천징수의무자는 「소득세법 시행령」 제185조제1항에따라 위의 내용을 제출하며, <strong><br>&nbsp;&nbsp;위 내용을 충분히 검토하였고 원천징수의무자가 알고 있는 사실 그대<br>&nbsp;&nbsp;로를 정확하게 적었음을 확인합니다</strong>
    <br>
    <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <%=mid(cstr(now()),1,4)%>년&nbsp;<%=mid(cstr(now()),6,2)%>월&nbsp;<%=mid(cstr(now()),9,2)%>일&nbsp;&nbsp;&nbsp;
    <br>
    <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    원천징수의무자&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=company_name%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(서명 또는 인)&nbsp;&nbsp;&nbsp;
    <br>
    <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=owner_name%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
    <br>
    <br>
    &nbsp;&nbsp;&nbsp;&nbsp;<strong>세무대리인은 조세전문자격자로서 위 신고서를 성실하고 공정하게<br>&nbsp;&nbsp;작성하였음을 확인합니다</strong>
    <br>
    <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    세무대리인&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=tax_man_name%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(서명 또는 인)&nbsp;&nbsp;&nbsp;
    <br>
    <br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;금천&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>세무서장</strong>&nbsp;&nbsp;귀하
    </span>
    </td>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">신고서 부포등 작성여부<br>※ 해당란에 "0"표시를 합니다</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">부표(4-5)쪽</span></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">환급(7쪽∼9쪽)</span></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">승계명세(10쪽)</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">세무대리인</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">성명</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">사업자등록번호</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">전화번호</span></td>
    <td colspan="2" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</span></td>
  </tr>
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">국세환급금계좌신고<br>※환급금액 2천만원 미만인 경우에만 적습니다</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">예입처</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">예금종류</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">계좌번호</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
</table>
<table width="1030" cellpadding="0" cellspacing="0">
  <tr>
    <td width="33%" height="20" align="left"><span class="style10L">&nbsp;&nbsp;</span></td>
    <td width="*" height="20" align="center"><span class="style10L">&nbsp;&nbsp;</span></td>
    <td width="33%" height="20" align="right"><span class="style12L">210㎜×297㎜(백상지80g/㎡)</span></td>
  </tr>
</table>  
</p>	
</body>
</html>
