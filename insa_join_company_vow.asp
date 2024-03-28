<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

'	on Error resume next

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_agree = Server.CreateObject("ADODB.Recordset")
Set rs_max = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from emp_master where emp_no = '" + emp_no  + "'"
Rs.Open Sql, Dbconn, 1

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

year_cnt = datediff("yyyy", rs("emp_in_date"), curr_date)
mon_cnt = datediff("m", rs("emp_in_date"), curr_date)
day_cnt = datediff("d", rs("emp_in_date"), curr_date)

'response.write(year_cnt)
'response.write(mon_cnt)
'response.write(day_cnt)
seq_last = ""
agree_year = curr_year
agree_id = "입사서약서"       

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
'response.write(cfm_number)
'response.write(cfm_seq)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script src="/java/common.js" type="text/javascript"></script>
<script type="text/javascript">
	function printWindow(){
//		viewOff("button");   
		factory.printing.header = ""; //머리말 정의
		factory.printing.footer = ""; //꼬리말 정의
		factory.printing.portrait = true; //출력방향 설정: true - 가로, false - 세로
		factory.printing.leftMargin = 5; //외쪽 여백 설정
		factory.printing.topMargin = 15; //윗쪽 여백 설정
		factory.printing.rightMargin = 5; //오른쯕 여백 설정
		factory.printing.bottomMargin = 0; //바닦 여백 설정
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
<title>입 사 서 약 서</title>
	<style type="text/css">
    <!--
    	.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style1 {font-size:13px;color: #666666 font-family: "바탕체", "바탕체",}
    -->
    </style>
    <style media="print"> 
    .noprint     { display: none }
    </style>
    <style type="text/css" media="screen">
    .onlyprint {display:; }
    </style>

	</head>
        
    <body>
    <div align=center class="noprint">
     <p>
        <%'<a href="javascript:printW();"><img src="image/b_print.gif" border="0" /></a> %>
        <a href="javascript:goBefore();"><img src="image/b_close.gif" border="0" /></a>
        <td>
        <input type="image" name="rptCert$ctl00$btnRequest" id="rptCert_ctl00_btnRequest" src="/image/btn_career_certificate.gif" alt="입사사약서동의" onclick="s_sinchung('<%=rs("emp_no")%>','<%=rs("emp_name")%>', 'cfm_use', 'cfm_use_dept', 'cfm_comment');return false;" style="border-width:0px;" />
        </td>
     </p>
    </div>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
    </object>

	<p style='page-break-before:always'><br style='height:0; line-height:0'> 

        <table width="750" border="1" cellspacing="10" cellpadding="0" align="center" class="onlyprint" style="border:10px solid #0072BE;">
          <tr>
             <td width="100%" height="100%" bgcolor="ffffff" align="center" valign="top" style="padding-left:20px; padding-right:20px;" >
	             <table width="100%" border="0" cellspacing="0" cellpadding="0">
	               <tr>
		             <td height="120" align="center" valign="middle"><font style="font-size:22px"><strong>입 사 서 약 서</strong></td>
	               </tr>
	               <tr>
		             <td valign="middle" align="center">
		               <table width="660" cellspacing="0" cellpadding="12"  style="border:1px solid #000000;">
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;"><span class="style1">주민등록번호</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000"><span class="style1"><strong><%=rs("emp_person1")%>-<%=rs("emp_person2")%></strong></td>
                            <td rowspan="2" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">생년월일</span></td>
                            <td rowspan="2 align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=mid(cstr(rs("emp_birthday")),1,4)%>년&nbsp;<%=mid(cstr(rs("emp_birthday")),6,2)%>월&nbsp;<%=mid(cstr(rs("emp_birthday")),9,2)%>일생</strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;"><span class="style1">성&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;명</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=rs("emp_name")%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">주&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;소</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=rs("emp_sido")%>&nbsp;<%=rs("emp_gugun")%>&nbsp;<%=rs("emp_dong")%>&nbsp;<%=rs("emp_addr")%>&nbsp;&nbsp;&nbsp;&nbsp;TEL&nbsp;:&nbsp;<%=rs("emp_tel_ddd")%>-<%=rs("emp_tel_no1")%>-<%=rs("emp_tel_no2")%> </strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA;"><span class="style1">입사일</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=mid(cstr(rs("emp_in_date")),1,4)%>년&nbsp;<%=mid(cstr(rs("emp_in_date")),6,2)%>월&nbsp;<%=mid(cstr(rs("emp_in_date")),9,2)%>일</strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=rs("emp_company")%>-<%=rs("emp_org_name")%></strong></td>
                         </tr>
                </table></td>
	       </tr>
           <tr>
		      <td height="60" align="left" width="600"><font style="font-size:14px"><span class="style1"><br/>&nbsp;본인은 금번 회사에 채용되어 근무하게 되었사온 봐 하기 사항들을 엄수하여 성실하게<br/>회사의 근무에 임할 것을 서약합니다.<br/><br/>
		1) 회사의 제규정을 준수함은 물론, 회사의 제반지시사항에 복종하겠습니다.br/><br/>
        2) 회사내에서 맞은 바 직무수행에 전력을 경주하여 소관 사무 및 주어진 업무를 성실히<br/>
        처리하며 고의 또는 태만으로 명령취지에 위반됨이 없도록 하겠습니다.<br/><br/>
        3) 전임, 직무이동, 출장 등에 관한 회사명령에 대하여는 절대 불평함이 없이 순중하겠습니다.<br/><br/>
        4) 회사의 업무상 기밀에 속하는것은 재직중은 물론 퇴직 후에라고 일체 누설하지 않습니다.<br/><br/>
        5) 회사의 금품을 부당하게 개인적으로 이용하거나 업무를 빙자하여 사리를 도모하는 일이 <br/>
        없도록 하겠습니다.<br/><br/>
        6) 신의를 존중하고 품성을 도양하여 자기 인격 향상을 기함은 물론 회사직원으로서의 <br/>
        명예를 손상하게 함이 없도록 하겠습니다.<br/><br/>
        7) 만일 상기의 서약 사항을 위반하여 회사의 업무처리상 장해를 야기하게 하였거나<br/>
        회사에 손해를 끼치게 된 경우에는 여하한 처벌도 감수하겠으며 해당 손해액은 지체없이<br/>
        변상하겠습니다. </td>
	      </tr>
	       <tr>
		      <td height="60" align="center" width="600"><font style="font-size:14px"><%=mid(cstr(now()),1,4)%>년&nbsp;<%=mid(cstr(now()),6,2)%>월&nbsp;<%=mid(cstr(now()),9,2)%>일<br/><br/></td>
	       </tr>
           <tr>
		      <td height="60" align="right" width="600"><font style="font-size:14px">성명&nbsp;&nbsp;<%=rs("emp_name")%>&nbsp;&nbsp;인<br/><br/></td>
	       </tr>
           <tr>
		      <td height="60" align="center"><font style="font-size:18px"><strong>(주)케이원정보통신 귀중</td>
	       </tr>
          
 <%         
' 		sql = "insert into emp_confirm(cfm_empno,cfm_number,cfm_seq,cfm_date,cfm_type,cfm_emp_name,cfm_company,cfm_org_name,cfm_job,cfm_position,cfm_person1,cfm_person2,cfm_use,cfm_use_dept,cfm_comment) values "
'		sql = sql +	" ('"&emp_no&"','"&cfm_number&"','"&cfm_seq&"','"&curr_date&"','"&cfm_type&"','"&cfm_emp_name&"','"&cfm_company&"','"&cfm_org_name&"','"&cfm_job&"','"&cfm_position&"','"&cfm_person1&"','"&cfm_person2&"','"&cfm_use&"','"&cfm_use_dept&"','"&cfm_comment&"')"
		
'		dbconn.execute(sql)
		
 %>         

	   </table>
	<br><br><br>
	
		
	   </td>
    </tr>
    </table>
 </p> 

    </body>
    </html>
