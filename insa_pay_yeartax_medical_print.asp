<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

emp_no=Request("emp_no")
emp_name=Request("emp_name")
inc_yyyy=Request("inc_yyyy")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")
emp_person = cstr(rs_emp("emp_person1")) + "-" + cstr(rs_emp("emp_person2"))	
rs_emp.close()	

if emp_company = "케이원정보통신" then
      company_name = "(주)" + "케이원정보통신"
	  owner_name = "김승일"
	  addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	  trade_no = "107-81-54150"
	  tel_no = "02) 853-5250"
	  e_mail = "js10547@k-won.co.kr"
   elseif emp_company = "휴디스" then
              company_name = "(주)" + "휴디스"
			  owner_name = "김한종"
	          addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	          trade_no = "107-81-54150"
	          tel_no = "02) 853-5250"
	          e_mail = "js10547@k-won.co.kr"
		  elseif emp_company = "케이네트웍스" then
                     company_name = "케이네트웍스" + "(주)"
					 owner_name = "이중원"
	                 addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                 trade_no = "107-81-54150"
	                 tel_no = "02) 853-5250"
	                 e_mail = "js10547@k-won.co.kr"
				 elseif emp_company = "에스유에이치" then
                        company_name = "(주)" + "에스유에이치"	
						owner_name = "박미애"
	                    addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                    trade_no = "119-86-78709"
	                    tel_no = "02) 6116-8248"
	                    e_mail = "pshwork27@k-won.co.kr"
end if 

tot_cnt = 0
tot_amt = 0

sql = "select * from pay_yeartax_medical where m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' ORDER BY m_emp_no,m_person_no,m_seq ASC"
rs_medi.Open Sql, Dbconn, 1
'Set rs_medi = DbConn.Execute(SQL)
do until rs_medi.eof
         tot_cnt = tot_cnt + int(rs_medi("m_cnt"))	
		 tot_amt = tot_amt + int(rs_medi("m_amt"))
	rs_medi.MoveNext()
loop
rs_medi.close()	

sql = "select * from pay_yeartax_medical where m_year = '"&inc_yyyy&"' and m_emp_no = '"&emp_no&"' ORDER BY m_emp_no,m_person_no,m_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "의료비 지급명세서"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
		  		 window.close () ;
			}
			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = true; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
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
        </script>
    <style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
-->
    </style>
	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="wrap">			
			<div id="container">
				<form action="insa_pay_yeartax_medical_print.asp" method="post" name="frm">
				<div class="gView">
				<table width="1150" cellpadding="0" cellspacing="0">
                   <tr>
                      <td class="style20C"><%=title_line%></td>
                   </tr>
                   <tr>
                      <td height="20" class="style20C">&nbsp;</td>
                   </tr>
                </table>
                <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
				        <colgroup>
							<col height="30px" width="20%" >
							<col height="30px" width="30%" >
							<col height="30px" width="20%" >
							<col height="30px" width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td colspan="4" height="30" align="center" class="style12C">소득자 인적 사항</td>
						    </tr>
                            <tr>
							  <th height="30" align="left" style=" border-top:1px solid #e3e3e3;">①성명</th>
                              <td height="30" align="center"><%=emp_name%></td>
                              <th height="30" align="left" style=" border-top:1px solid #e3e3e3;">②주민등록번호(또는 외국인등록번호)</th>
                              <td height="30" align="center"><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th height="30" align="left">③법인명</th>
                              <td height="30" align="center"><%=company_name%></td>
                              <th height="30" align="left">④업체명</th>
                              <td height="30" align="center">&nbsp;</td>
						    </tr>
                            <tr>
                              <td colspan="4" height="30" align="center" class="style12C">(<%=inc_yyyy%>) 년 의료비 지급명세</td>
						    </tr>
						</thead>
				  </table>
					<table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="8%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="12%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="12%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <th colspan="4" height="30" align="center" scope="col" style=" border-bottom:1px solid #e3e3e3;">대상자</th>
                                <th colspan="3" scope="col" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;">지급처</th>
                                <th colspan="2" scope="col" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;">지급명세</th>
                              </tr>
                              <tr>
                                <th class="first" scope="col">관계코드</th>
                                <th scope="col">내외</th>
                                <th scope="col">⑤주민등록번호</th>
                                <th scope="col">⑥본인등<br>해당여부</th>
                                <th scope="col">⑦사업자등록번호</th>
                                <th scope="col">⑧상호</th>
                                <th scope="col">⑨의료증빙코드</th>
                                <th scope="col">⑩건수</th>
                                <th scope="col">⑪금액</th>
                              </tr>
                        </thead>
						<tbody>
                              <tr>
                                <td colspan="7" class="first" height="30" align="center">합계</td>
                                <td align="right"><%=formatnumber(tot_cnt,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(tot_amt,0)%>&nbsp;</td>
							  </tr>                        
				     <%
						do until rs.eof
                             m_rel = ""
							 if rs("m_rel") = "본인" then 
							          m_rel = "0"
							    elseif rs("m_rel") = "부" or rs("m_rel") = "모" or rs("m_rel") = "조부" or rs("m_rel") = "조모" then 
							                 m_rel = "1"
									   elseif rs("m_rel") = "장인" or rs("m_rel") = "장모" then 
							                        m_rel = "2"
											  elseif rs("m_rel") = "남편" or rs("m_rel") = "아내" then 
							                               m_rel = "3"
													 elseif rs("m_rel") = "아들" or rs("m_rel") = "딸" then 
							                                      m_rel = "4"
														    elseif rs("m_rel") = "손자" or rs("m_rel") = "손녀" then 
							                                             m_rel = "5"
																   elseif rs("m_rel") = "형(형제자매)" or rs("m_rel") = "제(형제자매)" or rs("m_rel") = "매(형제자매)" or rs("m_rel") = "자(형제자매)" then 
							                                                    m_rel = "6"
																		  elseif rs("m_witak") = "Y" then
																		               m_rel = "7"
																				 elseif rs("m_pensioner") = "Y" then
																		               m_rel = "8"
							 end if
							 m_bon = ""
							 if rs("m_rel") = "본인" or rs("m_disab") = "Y" or rs("m_age65") = "Y" then 	
							          m_bon = "1"
								else  
								      m_bon = "2"
							 end if	
							 m_data_gubun = ""
							 if rs("m_data_gubun") = "국세청"	then
							 		  m_data_gubun	= "1"
								elseif rs("m_data_gubun") = "국민건강보험공단"	then
							 		         m_data_gubun	= "2"
									   elseif rs("m_data_gubun") = "진료비/약제비"	then
							 		                m_data_gubun	= "3"
											  elseif rs("m_data_gubun") = "장기요양급여"	then
							 		                       m_data_gubun	= "4"
											         elseif rs("m_data_gubun") = "기타의료비영수증"	then
							 		                              m_data_gubun	= "5"
																  
							 end if		  						   
	           			%>
							<tr>
                                <td class="first" height="30" align="center"><%=m_rel%>&nbsp;</td>
                                <td align="center"><%=rs("m_national")%>&nbsp;</td>
                                <td align="center"><%=rs("m_person_no")%>&nbsp;</td>
                                <td align="center"><%=m_bon%>&nbsp;</td>
                                <td align="center"><%=rs("m_trade_no")%>&nbsp;</td>
                                <td align="center"><%=rs("m_trade_name")%>&nbsp;</td>
                                <td align="center"><%=m_data_gubun%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("m_cnt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("m_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<table width="1150" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<br>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				    <br>                 
                    </td>
			      </tr>
				</table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

