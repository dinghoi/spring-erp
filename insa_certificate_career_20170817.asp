<%@LANGUAGE="VBSCRIPT"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim from_date
Dim to_date

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

emp_no=Request.form("in_empno")
emp_name=Request.form("in_name")

cfm_use=Request.form("cfm_use")
cfm_use_dept=Request.form("cfm_use_dept")
cfm_comment=Request.form("cfm_comment")

'response.write(cfm_use)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from emp_master where emp_no = '" & emp_no  & "'"
Rs.Open Sql, Dbconn, 1

if not Rs.eof then
		emp_company = Rs("emp_company")
   else
   	    emp_company = ""
end if

if emp_company = "케이원정보통신" then
      emp_company = "(주)" + "케이원정보통신"
   elseif emp_company = "휴디스" then
              emp_company = "(주)" + "휴디스"
		  elseif emp_company = "케이네트웍스" then
                     emp_company = "케이네트웍스" + "(주)"
				 elseif emp_company = "에스유에이치" then
                            emp_company = "(주)" + "에스유에이치"	
						elseif emp_company = "코리아디엔씨" then
                                   emp_company = "코리아디엔씨" + "(주)"	
end if 

cfm_company = rs("emp_company")
cfm_emp_name = rs("emp_name")
cfm_org_name = rs("emp_org_name")
cfm_job = rs("emp_job")
cfm_position = rs("emp_position")
cfm_person1 = rs("emp_person1")
cfm_person2 = rs("emp_person2")

emp_in_date = mid(cstr(rs("emp_in_date")),1,10)
emp_in_year = mid(cstr(rs("emp_in_date")),1,4)
emp_in_month = mid(cstr(rs("emp_in_date")),6,2)
emp_in_day = mid(cstr(rs("emp_in_date")),9,2)

if rs("emp_end_date") = "1900-01-01" or isNull(rs("emp_end_date")) then
   emp_end_date = ""
   target_date = curr_date
   else 
   emp_end_date = rs("emp_end_date")
   target_date = rs("emp_end_date")
end if

year_cnt = datediff("yyyy", rs("emp_in_date"), target_date)
mon_cnt = datediff("m", rs("emp_in_date"), target_date)
day_cnt = datediff("d", rs("emp_in_date"), target_date)

'response.write(year_cnt)

year_cnt = int(year_cnt) + 1
mon_cnt = int(mon_cnt) + 1
day_cnt = int(day_cnt) + 1
y_cnt = int(mon_cnt / 12)
m_cnt = mon_cnt - (y_cnt * 12)

if y_cnt > 0 and m_cnt > 0 then 
       stay_tit = cstr(y_cnt) + "년 " + cstr(m_cnt) + "개월"
   elseif y_cnt > 0 and m_cnt = 0 then 
              stay_tit = cstr(y_cnt) + "년 "
		  elseif y_cnt = 0 and m_cnt > 0 then 
		             stay_tit = cstr(m_cnt) + "개월 "
			     elseif y_cnt = 0 and m_cnt = 0 then 
				            stay_tit = cstr(m_cnt) + "개월 "
end if

seq_last = ""
cfm_number = curr_year
cfm_type = "경력증명서"       

    sql="select max(cfm_seq) as max_seq from emp_confirm where cfm_type = '"&cfm_type&"' and cfm_number = '"&curr_year&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		seq_last = "0001"
	  else
		max_seq = "000" + cstr((int(rs_max("max_seq")) + 1))
		seq_last = right(max_seq,4)
	end if
    rs_max.close()

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
	}
	function printW() {
        window.print();
    }
	function goBefore () {
		history.back() ;
	}
</script>
<title>경력증명서 출력</title>
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
		.style1 {font-size:14px;color: #666666}
		.style2 {font-size:12px;color: #666666}
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
		             <td align="left" height="60" valign="middle" style="padding-left:20px;" >제<%=cfm_number%>―<%=cfm_seq%>&nbsp;호</td>
	               </tr>
	               <tr>
		             <td height="130" align="center" valign="middle"><strong class="style32BC">경 력 증 명 서</strong></td>
	               </tr>
	               <tr>
		             <td valign="middle" align="center">
		               <table width="600" cellspacing="0" cellpadding="12"  style="border:1px solid #000000;">
                         <tr>
                            <td width="22%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;"><span class="style1">성&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;명</span></td>
                            <td width="28%" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=rs("emp_name")%></strong></td>
                            <td width="22%" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">주민등록번호</span></td>
                            <td width="28%" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=rs("emp_person1")%>-<%=emp_person2%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=rs("emp_org_name")%></strong></td>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;위 </span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=rs("emp_job")%></strong></td>
                         </tr>
                         <tr>
                         <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">현근무지입사일</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=mid(cstr(rs("emp_in_date")),1,4)%>년&nbsp;<%=mid(cstr(rs("emp_in_date")),6,2)%>월&nbsp;<%=mid(cstr(rs("emp_in_date")),9,2)%>일</strong></td>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">퇴&nbsp;&nbsp;&nbsp;&nbsp;사&nbsp;&nbsp;&nbsp;&nbsp;일 </span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=mid(cstr(emp_end_date),1,4)%>년&nbsp;<%=mid(cstr(emp_end_date),6,2)%>월&nbsp;<%=mid(cstr(emp_end_date),9,2)%>일</strong></td>
                         </tr>
                         <tr>
                         <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">근&nbsp;&nbsp;무&nbsp;&nbsp;월&nbsp;&nbsp;수</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=stay_tit%></strong></td>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;도</span>&nbsp;</td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=cfm_use%>&nbsp;-<%=cfm_use_dept%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">주&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;소</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=rs("emp_sido")%>&nbsp;<%=rs("emp_gugun")%>&nbsp;<%=rs("emp_dong")%>&nbsp;<%=rs("emp_addr")%></strong></td>
                         </tr>
                        <tr>
                           <td height="30" align="center" valign="middle" style="border-right:1px solid #000000; background-color:#EAEAEA;"><span class="style1">비&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;고</span></td>
                           <td colspan="3"><span class="style1"><strong><%=cfm_comment%></strong></td>
                       </tr>
                </table></td>
	       </tr>
	       <tr>
		      <td height="280" align="center"><font style="font-size:16px"><strong>위 내용이 사실임을 증명함</td>
	       </tr>
	       <tr>
          <% if cfm_company = "케이네트웍스" then %>
		      <td height="60" align="right" width="600"><font style="font-size:14px"><%=mid(cstr(now()),1,4)%>년&nbsp;<%=mid(cstr(now()),6,2)%>월&nbsp;<%=mid(cstr(now()),9,2)%>일<br/><br/>
		서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 4층)</td>
          <%    else %>
              <td height="60" align="right" width="600"><font style="font-size:14px"><%=mid(cstr(now()),1,4)%>년&nbsp;<%=mid(cstr(now()),6,2)%>월&nbsp;<%=mid(cstr(now()),9,2)%>일<br/><br/>
		서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)</td>
          <% end if %>                  
	      </tr>
	      <tr>
	         <% if cfm_company = "케이원정보통신" then %>
	         <td height="60" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>주식회사 케이원정보통신<br />
		<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>김승일</b></font></td>
          <% end if %>
          <% if cfm_company = "휴디스" then %>
	         <td height="60" align="right" valign="middle" width="100%"><img src="image/k_hudis001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>주식회사 휴디스<br />
		<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>빅영진</b></font></td>
          <% end if %>
          <% if cfm_company = "케이네트웍스" then %>
	         <td height="60" align="right" valign="middle" width="100%"><img src="image/k_net001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>케이네트웍스 주식회사<br />
		<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>김승일</b></font></td>
          <% end if %>
          <% if cfm_company = "에스유에이치" then %>
	         <td height="60" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>주식회사 에스유에이치<br />
		<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>박미애</b></font></td>
          <% end if %>
          <% if cfm_company = "코리아디엔씨" then %>
	         <td height="60" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>코리아디엔씨 주식회사<br />
		<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>송관섭</b></font></td>
          <% end if %>
	      </tr>
	   </table>
	<br><br><br>
	
		
	   </td>
    </tr>
    
 <%         
 		sql = "insert into emp_confirm(cfm_empno,cfm_number,cfm_seq,cfm_date,cfm_type,cfm_emp_name,cfm_company,cfm_org_name,cfm_job,cfm_position,cfm_person1,cfm_person2,cfm_use,cfm_use_dept,cfm_comment,cfm_reg_date,cfm_reg_user) values "
		sql = sql +	" ('"&emp_no&"','"&cfm_number&"','"&cfm_seq&"','"&curr_date&"','"&cfm_type&"','"&cfm_emp_name&"','"&cfm_company&"','"&cfm_org_name&"','"&cfm_job&"','"&cfm_position&"','"&cfm_person1&"','"&cfm_person2&"','"&cfm_use&"','"&cfm_use_dept&"','"&cfm_comment&"',now(),'"&emp_user&"')"
		
		dbconn.execute(sql)
		
'		dbconn.CommitTrans
'		dbconn.Close()
'	    Set dbconn = Nothing		
	
 %>         
    
    </table>
     </body>
    </html>
