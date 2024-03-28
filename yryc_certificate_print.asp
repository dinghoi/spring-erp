<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

'	on Error resume next

emp_user		= request.cookies("nkpmg_user")("coo_user_name")

curr_date		= mid(cstr(now()),1,10)
curr_year		= mid(cstr(now()),1,4)
curr_month	= mid(cstr(now()),6,2)
curr_day		= mid(cstr(now()),9,2)

'//전달 변수
reqYrycSn		= Request("yryc_sn")		'//연차 순번


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_max = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'//
sql = "select a.* "
sql = sql & " ,DATE_ADD(emp_end_date, INTERVAL -2 week) as writeDate "
sql = sql & " from emp_use_yryc a "
sql = sql & " where yryc_sn = " & reqYrycSn
Rs.Open Sql, Dbconn, 1

If rs.bof Or rs.eof Then
	Set rs = Nothing
	dbconn.Close()
	Set dbconn = Nothing
	Response.write "정보가 정확하지 않습니다."
	Response.end
End If


'//
yrycSn						= rs("yryc_sn")
empName					= rs("emp_name")
empPerson1				= rs("emp_person1")
empPerson2				= rs("emp_person2")
empFirstDate			= rs("emp_first_date")
empEndDate		= rs("emp_end_date")
yrycDaycnt				= rs("yryc_daycnt")
yrycUseDaycnt		= rs("yryc_use_daycnt")
regDate					= rs("reg_date")
prntngCo				= rs("prntng_co")
writeDate				= rs("writeDate")


'//연차 사용 일수 구하기
yrycUseDaycntExtra = CDbl(toString(yrycDaycnt,"0")) - CDbl(toString(yrycUseDaycnt,"0"))
if CDbl(yrycUseDaycntExtra)<0 Then yrycUseDaycntExtra = 0 End If

'emp_in_date = mid(cstr(rs("emp_in_date")),1,10)
'emp_in_year = mid(cstr(rs("emp_in_date")),1,4)
'emp_in_month = mid(cstr(rs("emp_in_date")),6,2)
'emp_in_day = mid(cstr(rs("emp_in_date")),9,2)

'year_cnt = datediff("yyyy", rs("emp_in_date"), curr_date)
'mon_cnt = datediff("m", rs("emp_in_date"), curr_date)
'day_cnt = datediff("d", rs("emp_in_date"), curr_date)

'response.write(year_cnt)
'response.write(mon_cnt)
'response.write(day_cnt)
seq_last = ""
cfm_number = curr_year
cfm_type = "재직증명서"       

'    sql="select max(cfm_seq) as max_seq from emp_confirm where cfm_type = '"&cfm_type&"' and cfm_number = '"&curr_year&"'"
'	set rs_max=dbconn.execute(sql)
	
'	if	isnull(rs_max("max_seq"))  then
'		seq_last = "0001"
'	  else
'		max_seq = "000" + cstr((int(rs_max("max_seq")) + 1))
'		seq_last = right(max_seq,4)
'	end if
 '   rs_max.close()

cfm_seq = seq_last
'response.write(cfm_number)
'response.write(cfm_seq)
emp_person2 = "*******"

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
		<% 'document.frm.action = "insa_certificate_print.asp"; 출력후 발행내역 DB저장하는것 정리할것%>
	}
	function printW() {
        window.print();
    }
	function goBefore () {
		history.back() ;
	}
</script>
<title>연차유급휴가 사용일수 확인 출력</title>
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
    <div align=center class="noprint">
     <p> 
        <a href="javascript:printWindow();"><img src="image/b_print.gif" border="0" /></a>
        <a href="javascript:goBefore();"><img src="image/b_close.gif" border="0" /></a>
     </p>
    </div>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
    </object>
        <table width="750" border="1" cellspacing="10" cellpadding="0" align="center" class="onlyprint" style="border:10px solid #0072BE;">
          <tr>
             <td width="100%" height="100%" bgcolor="ffffff" align="center" valign="top" style="padding-left:20px; padding-right:20px;" >
	             <table width="100%" border="0" cellspacing="0" cellpadding="0">
	               <tr>
		             <td align="left" height="20" valign="middle" style="padding-left:20px;" ><!-- 제<%=cfm_number%>―<%=cfm_seq%>&nbsp;호--></td>
	               </tr>
	               <tr>
		             <td height="150" align="center" valign="middle"><strong class="style32BC">연차유급휴가 사용일수 확인</strong></td>
	               </tr>
	               <tr>
		             <td valign="middle" align="center">
		               <table width="560" cellspacing="0" cellpadding="12"  style="border:1px solid #000000;">
                         <tr>
                            <td width="150px" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;"><span class="style2">성&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;명</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=empName%></strong></td>
                         </tr>
						 <tr>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style2">주민등록번호</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=empPerson1%>-*******<%'=emp_person2%></strong></td>
						 </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style2">입&nbsp;&nbsp;&nbsp;&nbsp;사&nbsp;&nbsp;&nbsp;&nbsp;일</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=empFirstDate%></strong></td>
                         </tr>
						 <tr>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style2">퇴&nbsp;&nbsp;&nbsp;&nbsp;사&nbsp;&nbsp;&nbsp;&nbsp;일 </span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style2"><strong><%=empEndDate%></strong></td>
						</tr>


                        <tr>
                           <td height="30" align="center" valign="middle" style="border-right:1px solid #000000; background-color:#EAEAEA;"><span class="style2">연차휴가발생일수</span></td>
                           <td colspan="3"><span class="style2"><strong><%=yrycDaycnt%></strong></td>
                       </tr>
                </table></td>
	       </tr>
	       <tr>
		      <td height="380px" align="center"><span style="font-size:18px;"><strong>상기 본인은 연차휴가 발생일수 ( <%=yrycDaycnt%> ) 일 중
				<br /><br />( <%=yrycUseDaycnt%> )일을 사용하고, 잔여휴가일수 ( <%=yrycUseDaycntExtra%> ) 일을 
				<br/><br />퇴사시 모두 소진하였음을 확인합니다.
				</strong>
				</span>
			  </td>
	       </tr>

	       <tr>
              <td height="30" align="center" width="600"><font style="font-size:14px"><%=mid(writeDate,1,4)%>년&nbsp;<%=Cint(mid(writeDate,6,2))%>월&nbsp;<%=Cint(mid(writeDate,9,2))%>일</td>
		  </tr>
	       <tr>
              <td height="30" align="center" width="600"><font style="font-size:14px">확인자 : <%=empName%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(인)
				</font></td>
	      </tr>
	      <tr>  
	         <td height="60" align="center" valign="middle" width="100%"><font style="font-size:14px"><strong>(주)케이원정보통신 귀중</strong></font></td>
	      </tr>
	      <tr>  
	         <td height="50" align="left" valign="bottom" width="100%"><font style="font-size:14px">(주)케이원정보통신 그룹웨어를 통한 신청문서입니다.</font></td>
	      </tr>
       </table>
	<br><br><br>
	
		
	   </td>
    </tr> 
 <%         
		If Trim(yrycSn&"") <> "" Then
			sql = "update  emp_use_yryc set prntng_co= prntng_co+1 where yryc_sn = " & yrycSn
			dbconn.execute(sql)
		End IF
		
'		dbconn.CommitTrans
'		dbconn.Close()
'	    Set dbconn = Nothing
		
 %>         
    </table>
    </body>
    </html>