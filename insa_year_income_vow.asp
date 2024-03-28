<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_agree = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_max = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

sql = "select * from emp_master where emp_no = '" + emp_no  + "'"
Rs.Open Sql, Dbconn, 1

agree_empno = rs("emp_no")
agree_emp_type = rs("emp_type")
agree_empname = rs("emp_name")
agree_company = rs("emp_company")
agree_org_name = rs("emp_org_name")
agree_grade = rs("emp_grade")
agree_job = rs("emp_job")
agree_position = rs("emp_position")
agree_jikmu = rs("emp_jikmu")
agree_person1 = rs("emp_person1")
agree_person2 = rs("emp_person2")
agree_sido = rs("emp_sido")
agree_gugun = rs("emp_gugun")
agree_dong = rs("emp_dong")
agree_addr = rs("emp_addr")
agree_tel_ddd = rs("emp_tel_ddd")
agree_tel_no1 = rs("emp_tel_no1")
agree_tel_no2 = rs("emp_tel_no2")

emp_in_date = mid(cstr(rs("emp_in_date")),1,10)
emp_in_year = mid(cstr(rs("emp_in_date")),1,4)
emp_in_month = mid(cstr(rs("emp_in_date")),6,2)
emp_in_day = mid(cstr(rs("emp_in_date")),9,2)

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)

year_cnt = datediff("yyyy", rs("emp_in_date"), curr_date)
mon_cnt = datediff("m", rs("emp_in_date"), curr_date)
day_cnt = datediff("d", rs("emp_in_date"), curr_date)
'rs.close()
'response.write(year_cnt)
'response.write(mon_cnt)
'response.write(day_cnt)
emp_no = "100173"

Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&curr_year&"'"
Set rs_year = DbConn.Execute(SQL)
if not rs_year.eof then
       incom_base_pay = rs_year("incom_base_pay")
       incom_overtime_pay = rs_year("incom_overtime_pay")
	   incom_meals_pay = rs_year("incom_meals_pay")
       incom_severance_pay = rs_year("incom_severance_pay")
	   incom_total_pay = rs_year("incom_total_pay")
	   incom_first3_percent = rs_year("incom_first3_percent")
   else
       incom_base_pay = 0
       incom_overtime_pay = 0
       incom_meals_pay = 0
	   incom_severance_pay = 0
       incom_total_pay = 0
	   incom_first3_percent = 0
end if
rs_year.close()

' 금액을 한글로 변환....
'amt = "21345000"
amt = incom_total_pay
Dim unit1(10)
Dim unit2(2)
Dim unit3(2)

unit1(0) = ""
unit1(1) = "일"
unit1(2) = "이"
unit1(3) = "삼"
unit1(4) = "사"
unit1(5) = "오"
unit1(6) = "육"
unit1(7) = "칠"
unit1(8) = "팔"
unit1(9) = "구"

unit2(0) = "십"
unit2(1) = "백"
unit2(2) = "천"

unit3(0) = "만"
unit3(1) = "억"
unit3(2) = "조"
 
vamt = Replace(amt, ",", "")
xchk = IsNumeric(vamt)

If xchk = True Then
    total = Len(CStr(CDbl(amt)))
    vamt = CDbl(amt)
    rt_amt = ""
    For i = 1 To total
        num = Mid(vamt, i, 1)
        temp1 = (total - i) + 1
        rt_amt = rt_amt & unit1(num)
 
        If num <> 0 And i <> total Then
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 0 Then rt_amt = rt_amt & unit2(2)
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 3 Then rt_amt = rt_amt & unit2(1)
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 2 Then rt_amt = rt_amt & unit2(0)
        End If
 
        If temp1 = 5 And Right(rt_amt, 1) <> unit3(2) And Right(rt_amt, 1) <> unit3(1) Then rt_amt = rt_amt & unit3(0)
        If temp1 = 9 And Right(rt_amt, 1) <> unit3(2) Then rt_amt = rt_amt & unit3(1)
        If temp1 = 13 Then rt_amt = rt_amt & unit3(2)
 
    Next
 
    rt_amt = rt_amt & "원"
 
    'msgbox
    'response.write  "input : " & amt & vbCr & "output : " & rt_amt
End If


seq_last = ""
agree_year = curr_year
agree_id = "연봉근로계약서"       

    sql="select max(agree_seq) as max_seq from emp_agree where agree_empno = '"&emp_no&"' and agree_year = '"&agree_year&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		seq_last = "001"
	  else
		max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		seq_last = right(max_seq,3)
	end if
    rs_max.close()

agree_seq = seq_last

main_title = cstr(agree_year) + "년 "  + " 연봉 근로계약서"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        <title>개인업무-인사</title>
        <script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript" id='dummy'></script>
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
	
    function year_income_agree(val, val2, val3) {
            if (!confirm("연봉근로계약을 동의하시겠습니까 ?")) return;
            var frm = document.frm; 
			
			alert (val);
			alert (val2);
			alert (val3);
            
			document.frm.action = "insa_year_income_agree_save.asp?emp_no=" + val;
            document.frm.submit();
			
			<%
			'var scpt= document.getElementById('dummy');
			'alert (scpt);
			'scpt.src='insa_year_income_agree_save.asp?emp_no='+val;
			'document.submit();
			%>
    }	
	
</script>
<title>개인 연봉근로계약서</title>
<style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style14BC {font-size: 14px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; text-align: center; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
		.style3 {font-size:14px;color: #666666}
-->
</style>
<style media="print"> 
.noprint     { display: none }
</style>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
    <form action="insa_year_income_vow.asp" method="post" name="frm">
    <div align=center class="noprint">
     <p>
        <% '<a href="javascript:printW();"><img src="image/b_print.gif" border="0" alt="출력하기" /></a> %>
        <a href="#" onClick="year_income_agree('<%=agree_empno%>','<%=agree_empname%>','<%=curr_year%>');return false;" style="border-width:0px;"><img src="image/b_agree2.jpg" border="0" alt="동의하기" /></a>
        <a href="#" onClick="printWindow()"><img src="image/b_print.gif" border="0" alt="출력하기" /></a>
        <a href="javascript:goBefore();"><img src="image/b_close.gif" border="0" alt="닫기" /></a> 
     </p>
    </div>
<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
</object>    
   
<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td colspan="3" align="center" class="style32BC"><%=main_title%></td>
  </tr>
  <tr>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
  </tr>
</table>
<table width="690" border="1px" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="10%" height="30" rowspan="3" align="center" bgcolor="#eaeaea"><span class="style14BC">사용자</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">사업체명</span></td>
    <td width="35%" height="30" align="center"><span class="style14C">&nbsp;&nbsp;(주)케이원정보통신</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">대표자</span></td>
    <td width="25%" height="30" align="center"><span class="style14C">&nbsp;&nbsp;김 승일</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">소재지</span></td>
    <td colspan="3" height="30" align="center"><span class="style14C">&nbsp;&nbsp;서울시 근천구 가산디지털2로 18 대륭테크노타운1차 6층</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea" style=" border-bottom:2px solid #515254;"><span class="style14BC">업종</span></td>
    <td colspan="3" height="30" align="center" style=" border-bottom:2px solid #515254;"><span class="style14C">&nbsp;&nbsp;컴퓨터 관련 운영 및 통신공사</span></td>
  </tr>
  <tr>
    <td width="10%" height="30" rowspan="2" align="center" bgcolor="#eaeaea"><span class="style14BC">근로자</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">성명</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=agree_empname%>&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">주민등록번호</span></td>
    <td width="25%" height="30" align="center"><span class="style14C"><%=agree_person1%>-*******&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">주소</span></td>
    <td colspan="3" height="30" align="center"><span class="style14C"><%=agree_sido%>&nbsp;<%=agree_gugun%>&nbsp;<%=agree_dong%>&nbsp;<%=agree_addr%></span></td>
  </tr>
</table>

<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     <br/>&nbsp;&nbsp;상기의 사용자와 근로자는 서로 동등한 지위에서 자유의사로 다음과 같이 근로계약을 체결하고 공동의<br/> 이익을 증진하기 위하여 이를 성실히 이행 할 것을 약정한다.<br/><br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="center" class="style3">
     <br/> -&nbsp;&nbsp; 다&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;음&nbsp;&nbsp; - <br/><br/></td>
  </tr>
</table>

<table width="690" border="1px" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">근로장소</span></td>
    <td width="35%" height="30" align="center"><span class="style14C">현업부서 소재지</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">근무형태</span></td>
    <td width="35%" height="30" align="center"><span class="style14C">정규직</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">직종</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=pmg_emp_no%>&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">직급</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=agree_grade%>&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">입사일</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=emp_in_year%>년&nbsp;<%=emp_in_month%>월&nbsp;<%=emp_in_day%>일</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">연락처</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=agree_tel_ddd%>-<%=agree_tel_no1%>-<%=agree_tel_no2%>&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">근로기간</span></td>
    <td colspan="3" height="30" align="center"><span class="style14C">근로기간에 대한 정함이 없음</span></td>
</table>

<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
<% if emp_in_year = curr_year then %>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>1. 임금지급 대상기간 :&nbsp;</strong><%=emp_in_year%>년&nbsp;<%=emp_in_month%>월&nbsp;<%=emp_in_day%>일 ~ </strong><%=mid(cstr(now()),1,4)%>년&nbsp;12월&nbsp;31일</td>
<%    else %>    
     <td width="100%" height="30" align="left" class="style3"><br/><strong>1. 임금지급 대상기간 :&nbsp;</strong><%=mid(cstr(now()),1,4)%>년&nbsp;01월&nbsp;01일 ~ </strong><%=mid(cstr(now()),1,4)%>년&nbsp;12월&nbsp;31일</td>
<%  end if %>    
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;상기 기간 중이라도 필요하다고 인정할 경우 상호 동의하에 년 급여액을 조정할 수 있다.<br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>2. 임금내역</strong></td>
  </tr>
</table>

<table width="690" border="1px" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">총수령액</span></td>
    <td colspan="3" height="30" align="center"><span class="style14C">일금&nbsp;:&nbsp;<%=rt_amt%>&nbsp;&nbsp;&nbsp;(\:<%=formatnumber(incom_total_pay,0)%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" rowspan="2" align="center" bgcolor="#eaeaea"><span class="style14BC">구분</span></td>
    <td width="35%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">해당금액</span></td>
    <td width="15%" height="30" rowspan="2" align="center" bgcolor="#eaeaea"><span class="style14BC">지급기준</span></td>
    <td width="35%" height="30" rowspan="2" align="center" bgcolor="#eaeaea"><span class="style14BC">비고</span></td>
  </tr>
  <tr>
    <td width="35%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">년 산정액</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">기본급</span></td>
    <td width="35%" height="30" align="right"><span class="style14C">\<%=formatnumber(incom_base_pay,0)%>&nbsp;&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">-</span></td>
    <td width="35%" height="30" align="center"><span class="style14C">&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">연장근로수당</span></td>
    <td width="35%" height="30" align="right"><span class="style14C">\<%=formatnumber(incom_overtime_pay,0)%>&nbsp;&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">포괄산정임금</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=pmg_emp_name%>&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">식대</span></td>
    <td width="35%" height="30" align="right"><span class="style14C">\<%=formatnumber(incom_meals_pay,0)%>&nbsp;&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">정규직직원</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=pmg_emp_name%>&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">퇴직금</span></td>
    <td width="35%" height="30" align="right"><span class="style14C">\<%=formatnumber(incom_severance_pay,0)%>&nbsp;&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">근무시간에준함</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=pmg_emp_name%>&nbsp;</span></td>
  </tr>
</table>

<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;① 상기의 내역 외 보직 및 업무 변경에 한하여 통신비, 주차비, 근속수당등 추가 지급한다.<br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;②“갑” 은 “을”에게 당월 초일에서 말일까지의 계산하여 당월 말일 지급한다<br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;③ 연차수당은 1년 이상 근무한 자에 한하여 지급한다.<br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>3. 수습기간</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;을의 수습기간은 신규 채용일로부터 3개월로 하며, 수습기간 중 또는 종료 후 부적절하다고 객관적으로<br/>&nbsp;&nbsp;&nbsp;&nbsp;인정된 경우에는 정식채용을 거부할 수 있다. 동기간의 급여는 제2조의 규정에도 불구하고 월 급여액의<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=incom_first3_percent%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;%로 지급한다.</td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>4. 퇴직금</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;을이 입사일로부터 1년 이상 근속 후 퇴직 시 근로기준법 및 단체협약에 따른 퇴직연금에 가입되어<br/>&nbsp;&nbsp;&nbsp;&nbsp;지급한다.</td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>5. 근로시간</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;① 근로일 : 주 5일, 1일 8시간, 주 40시간을 기준으로 한다.<br/>&nbsp;&nbsp;&nbsp;&nbsp;② 근무시간 : 오전 9시부터 오후 6시를 기준으로 하며 휴게시간은 12시부터 13시까지이며“을과”합의 후<br/>&nbsp;&nbsp;&nbsp;&nbsp;연장근무를 이행한다.<br/>&nbsp;&nbsp;&nbsp;&nbsp;③ 월 임금액 중 연장근로수당은 근로의 성질, 임금계산의 편의성 등 당사자간의 형편에 의하여 근무시<br/>&nbsp;&nbsp;&nbsp;&nbsp;당연히 발생 예상하는 법적제수당(연장근로등)이 포함된 포괄산정연봉 임금으로 한다.</td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>6. 연차는 근로기준법에 따른다</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>7. 급여공제</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;법령에서 정하는 세금,보험료등 근로자와 협정한 사항은 급여에서 공제할 수 있다.</td>
  </tr> 
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>8. 기타수당</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;직급수당, 차량유지비, 통신비, 직무보조비 등의 기타 수당은 해당업무와 직책에 따라 차등지급 될 수<br/>&nbsp;&nbsp;&nbsp;&nbsp;있으며 해당 업무에 대해 변경 적용 시 일할 계산 후 지급한다.</td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>9. 기타근로사항</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;명시되지 아니한 근로조건은 노동관계법령, 회사 제규정 및 통산관례에 따른다.</td>
  </tr>  
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>&nbsp;&nbsp;&nbsp;&nbsp;갑과 을은 상기와 같이 연봉근로계약에 합의합니다.</strong></td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="right" class="style3"><br/><%=mid(cstr(now()),1,4)%>년&nbsp;<%=mid(cstr(now()),6,2)%>월&nbsp;<%=mid(cstr(now()),9,2)%>일<br/><br/></td>
  </tr>  
</table>

<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td width="50%" height="30" align="left" class="style3"><strong>갑:&nbsp;(주)케이원정보통신</strong></td>
     <td width="50%" height="30" align="right" class="style3"><strong>대표이사&nbsp;&nbsp;&nbsp;김 승일&nbsp;&nbsp;(인)</strong></td>
  </tr>
  <tr>
     <td width="50%" height="30" align="left" class="style3"><strong>을:&nbsp;<%=agree_grade%></strong></td>
     <td width="50%" height="30" align="right" class="style3"><strong>성&nbsp;&nbsp;명&nbsp;&nbsp;&nbsp;<%=agree_empname%>&nbsp;&nbsp;(인)</strong></td>
  </tr>
</table>
</p>	

 <%         
' 		sql = "insert into emp_agree(cfm_empno,cfm_number,cfm_seq,cfm_date,cfm_type,cfm_emp_name,cfm_company,cfm_org_name,cfm_job,cfm_position,cfm_person1,cfm_person2,cfm_use,cfm_use_dept,cfm_comment) values "
'		sql = sql +	" ('"&emp_no&"','"&cfm_number&"','"&cfm_seq&"','"&curr_date&"','"&cfm_type&"','"&cfm_emp_name&"','"&cfm_company&"','"&cfm_org_name&"','"&cfm_job&"','"&cfm_position&"','"&cfm_person1&"','"&cfm_person2&"','"&cfm_use&"','"&cfm_use_dept&"','"&cfm_comment&"')"
		
'		dbconn.execute(sql)
		
 %>     
  </form>
</body>
</html>
