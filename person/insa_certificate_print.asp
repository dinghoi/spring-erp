<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim curr_date, curr_year, curr_month, curr_day
Dim emp_name, cfm_use, cfm_use_dept, cfm_comment
Dim rsCert, rsMax
Dim companyAddr

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

emp_no = Request.Form("in_empno")
emp_name = Request.Form("in_name")

cfm_use = Request.Form("cfm_use")
cfm_use_dept = Request.Form("cfm_use_dept")
cfm_comment = Request.Form("cfm_comment")

objBuilder.Append "SELECT emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "
objBuilder.Append "	emtt.emp_org_name, emtt.emp_name, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_person1, emtt.emp_person2, emtt.emp_in_date, emtt.emp_sido, "
objBuilder.Append "	emtt.emp_gugun, emtt.emp_dong, emtt.emp_addr, "
objbuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emp_no = '" & emp_no  & "' "

Set rsCert = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If not rsCert.eof Then
	'emp_company = rsCert("emp_company")
	emp_company = rsCert("org_company")
Else
	emp_company = ""
End If

Select Case emp_company
	Case "케이원" : emp_company = "(주)" & "케이원"
	Case "케이네트웍스" : emp_company = "(주)" & "케이네트웍스"
	Case "케이시스템" : emp_company = "(주)" & "케이시스템"
	Case "에스유에이치" : emp_company = "(주)" & "에스유에이치"
	Case "휴디스" : emp_company = "(주)" & "휴디스"
End Select

Dim cfm_company, cfm_emp_name, cfm_org_name, cfm_job, cfm_position
Dim cfm_person1, cfm_person2, emp_in_date, emp_in_year
Dim emp_in_month, emp_in_day, year_cnt, mon_cnt, day_cnt, seq_last
Dim cfm_number, cfm_type, max_seq, cfm_seq, emp_person2

cfm_company = rsCert("org_company")
cfm_emp_name = rsCert("emp_name")
cfm_org_name = rsCert("org_name")
cfm_job = rsCert("emp_job")
cfm_position = rsCert("emp_position")
cfm_person1 = rsCert("emp_person1")
cfm_person2 = rsCert("emp_person2")

emp_in_date = mid(cstr(rsCert("emp_in_date")),1,10)
emp_in_year = mid(cstr(rsCert("emp_in_date")),1,4)
emp_in_month = mid(cstr(rsCert("emp_in_date")),6,2)
emp_in_day = mid(cstr(rsCert("emp_in_date")),9,2)

year_cnt = datediff("yyyy", rsCert("emp_in_date"), curr_date)
mon_cnt = datediff("m", rsCert("emp_in_date"), curr_date)
day_cnt = datediff("d", rsCert("emp_in_date"), curr_date)

seq_last = ""
cfm_number = curr_year
cfm_type = "재직증명서"

objBuilder.Append "SELECT MAX(cfm_seq) AS max_seq FROM emp_confirm "
objBuilder.Append "WHERE cfm_type = '"&cfm_type&"' AND cfm_number = '"&curr_year&"' "

set rsMax = dbconn.execute(objBuilder.ToString())
objBuilder.Clear()

If isnull(rsMax("max_seq")) then
	seq_last = "0001"
else
	max_seq = "000" + cstr((int(rsMax("max_seq")) + 1))
	seq_last = right(max_seq,4)
end if
rsMax.close() : Set rsMax = Nothing

cfm_seq = seq_last
emp_person2 = cfm_person2
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript">
	//ActiveX 사용으로 IE11 외 브라우저에서 오류 발생(scriptX 미사용 처리)[허정호_20220204]
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

		//document.frm.action = "insa_certificate_print.asp"; 출력후 발행내역 DB저장하는것 정리할것
	}

	function printW(){
        window.print();
    }

	function goBefore () {
		location.href = "/person/insa_individual_confirm.asp";
	}

	//프린트 함수 신규 작성[허정호_20220204]
	var printArea;
	var initBody;

	function fnPrint(id){
		printArea = document.getElementById(id);

		window.onbeforeprint = beforePrint;
		window.onafterprint = afterPrint;

		window.print();
	}

	function beforePrint(){
		initBody = document.body.innerHTML;
		document.body.innerHTML = printArea.innerHTML;
	}

	function afterPrint(){
		document.body.innerHTML = initBody;
	}
</script>
<title>재직증명서 출력</title>
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
		.style1 {font-size:16px;color: #666666}
		.style2 {font-size:14px;color: #666666}
    -->
    </style>
    <style media="print">
    .noprint     { display: none }
    </style>
    <style type="text/css" media="screen">
    .onlyprint {display:; }
    </style>

	</head>

    <body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
    <div align="center" class="noprint">
     <p>
        <a href="javascript:fnPrint('print_pg');"><img src="/image/b_print.gif" border="0" /></a>
        <a href="javascript:goBefore();"><img src="/image/b_close.gif" border="0" /></a>
     </p>
    </div>
    <!--<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="http://intra.k-won.co.kr/smsx.cab#Version=7.0.0.8">
    </object>-->
	<div id="print_pg">
        <table width="750" border="1" cellspacing="10" cellpadding="1" align="center" class="onlyprint" style="border:10px solid #0072BE;">
          <tr>
             <td width="100%" height="100%" bgcolor="ffffff" align="center" valign="top" style="padding-left:20px; padding-right:20px;" >
	             <table width="100%" border="0" cellspacing="0" cellpadding="0">
	               <tr>
		             <td align="left" height="60" valign="middle" style="padding-left:20px;" >제<%=cfm_number%>―<%=cfm_seq%>&nbsp;호</td>
	               </tr>
	               <tr>
		             <td height="130" align="center" valign="middle"><strong class="style32BC">재 직 증 명 서</strong></td>
	               </tr>
	               <tr>
		             <td valign="middle" align="center">
		               <table width="560" cellspacing="1" cellpadding="12" style="border:1px solid #000000;">
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;"><span class="style2">성&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;명</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style2"><strong><%=rsCert("emp_name")%></strong></td>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style2">주민등록번호</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=rsCert("emp_person1")%>-<%=emp_person2%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style2">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style2"><strong><%=emp_company%>&nbsp;</strong></td>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style2">직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;위 </span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=rsCert("emp_job")%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style2">주&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;소</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=rsCert("emp_sido")%>&nbsp;<%=rsCert("emp_gugun")%>&nbsp;<%=rsCert("emp_dong")%>&nbsp;<%=rsCert("emp_addr")%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style2">용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;도</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=cfm_use_dept%>&nbsp;<%=cfm_use%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA;"><span class="style2">기&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;간</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=mid(cstr(rsCert("emp_in_date")),1,4)%>년&nbsp;<%=mid(cstr(rsCert("emp_in_date")),6,2)%>월&nbsp;<%=mid(cstr(rsCert("emp_in_date")),9,2)%>일부터&nbsp;&nbsp;현재까지</strong></td>
                        </tr>
                        <tr>
                           <td height="30" align="center" valign="middle" style="border-right:1px solid #000000; background-color:#EAEAEA;"><span class="style2">비&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;고</span></td>
                           <td colspan="3"><span class="style2"><strong><%=cfm_comment%></strong></td>
                       </tr>
                </table></td>
	       </tr>
	       <tr>
		      <td height="280" align="center"><font style="font-size:18px"><strong>위 내용이 사실임을 증명함</td>
	       </tr>
	       <tr>
			<%
			Select Case cfm_company
				Case "케이원" : companyAddr = "서울시 금천구 가산디지털2로 14, 대륭테크노타운12차 1405호"
				Case "케이네트웍스" : companyAddr = "서울시 금천구 가산디지털2로 18, 대륭테크노타운1차 605호"
				Case "케이시스템" : companyAddr = "서울시 금천구 가산디지털2로 18, 대륭테크노타운1차 406호"
				Case Else
					companyAddr = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
			End Select
			%>
			<%' if cfm_company = "케이원" Or cfm_company = "케이네트웍스" then %>
				<td height="60" align="right" width="600"><font style="font-size:14px"><%=Mid(CStr(Now()), 1, 4)%>년&nbsp;<%=Mid(CStr(Now()), 6, 2)%>월&nbsp;<%=Mid(CStr(Now()), 9, 2)%>일<br/><br/>
				<%=companyAddr%>
				</td>
			<%'  else %>
				<!--<td height="60" align="right" width="600"><font style="font-size:14px"><%=mid(cstr(now()),1,4)%>년&nbsp;<%=mid(cstr(now()),6,2)%>월&nbsp;<%=mid(cstr(now()),9,2)%>일<br/><br/>
				서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)
				</td>-->
			<%' end if %>
	      </tr>
	      <tr>
          <%
		  if cfm_company = "케이원" then %>
	         <!--<td height="60" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>주식회사 케이원정보통신<br />-->
			 <td height="60" align="right" valign="bottom" width="100%"><img src="/image/stamp/k_one_2021_001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>주식회사 케이원<br />
			<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>김승일</b></font></td>
          <% end if %>
          <% if cfm_company = "휴디스" then %>
	        <td height="60" align="right" valign="bottom" width="100%"><img src="/image/k_hudis001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>주식회사 휴디스<br />
			<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>박영진</b></font></td>
          <% end if %>
          <% if cfm_company = "케이네트웍스" then %>
	        <!--<td height="60" align="right" valign="middle" width="100%"><img src="image/k_net001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>케이네트웍스 주식회사<br />-->
			<td height="60" align="right" valign="bottom" width="100%"><img src="/image/stamp/k_net_2021_001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>주식회사 케이네트웍스<br />
			<!--<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>김승일</b></font><br/>-->
			<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>이동규</b></font></td>
          <% end if %>
          <% if cfm_company = "에스유에이치" then %>
	        <td height="60" align="right" valign="bottom" width="100%"><img src="/image/stamp/k_one_2021_001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>주식회사 에스유에이치<br />
			<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>박미애</b></font></td>
          <% end if %>
          <% if cfm_company = "케이시스템" then %>
	        <!--<td height="60" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>코리아디엔씨 주식회사<br />-->
			<td height="60" align="right" valign="bottom" width="100%"><img src="/image/stamp/k_sys_2021_001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>주식회사 케이시스템<br />
			<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>송관섭</b></font></td>
          <% end if %>
	      </tr>
       </table>
	<br><br><br>

	   </td>
    </tr>
 <%
 		'sql = "insert into emp_confirm(cfm_empno,cfm_number,cfm_seq,cfm_date,cfm_type,cfm_emp_name,cfm_company,cfm_org_name,cfm_job,cfm_position,cfm_person1,cfm_person2,cfm_use,cfm_use_dept,cfm_comment,cfm_reg_date,cfm_reg_user) values "
		'sql = sql +	" ('"&emp_no&"','"&cfm_number&"','"&cfm_seq&"','"&curr_date&"','"&cfm_type&"','"&cfm_emp_name&"','"&cfm_company&"','"&cfm_org_name&"','"&cfm_job&"','"&cfm_position&"','"&cfm_person1&"','"&cfm_person2&"','"&cfm_use&"','"&cfm_use_dept&"','"&cfm_comment&"',now(),'"&emp_user&"')"

		objBuilder.Append "INSERT INTO emp_confirm(cfm_empno,cfm_number,cfm_seq,cfm_date,cfm_type, "
		objBuilder.Append "	cfm_emp_name,cfm_company,cfm_org_name,cfm_job,cfm_position, "
		objBuilder.Append "	cfm_person1,cfm_person2,cfm_use,cfm_use_dept,cfm_comment, "
		objBuilder.Append "	cfm_reg_date,cfm_reg_user)VALUES("
		objBuilder.Append "'"&emp_no&"','"&cfm_number&"','"&cfm_seq&"','"&curr_date&"','"&cfm_type&"',"
		objBuilder.Append "'"&cfm_emp_name&"','"&cfm_company&"','"&cfm_org_name&"','"&cfm_job&"','"&cfm_position&"',"
		objBuilder.Append "'"&cfm_person1&"','"&cfm_person2&"','"&cfm_use&"','"&cfm_use_dept&"','"&cfm_comment&"',"
		objBuilder.Append "NOW(),'"&user_name&"') "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

'		dbconn.CommitTrans
		DBConn.Close() : Set DBConn = Nothing
 %>
	</table>
	</div>
</body>
</html>
