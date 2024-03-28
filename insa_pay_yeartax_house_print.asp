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
Set rs_dona = Server.CreateObject("ADODB.Recordset")
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

s_id = "연금저축"

sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산-월세액·거주가 간 주택임대차임금 원리금 산환액 소득공제 명세서"
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
				<form action="insa_pay_yeartax_house_print.asp" method="post" name="frm">
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
							<col height="30px" width="10%" >
                            <col height="30px" width="10%" >
							<col height="30px" width="30%" >
							<col height="30px" width="20%" >
							<col height="30px" width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td rowspan="4" height="30" align="left">1. 인적 사항</td>
                              <th height="30" align="left">①법인명</th>
                              <td height="30" align="center"><%=company_name%></td>
                              <th height="30" align="left">②업체명</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
							  <th height="30" align="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">③성명</th>
                              <td height="30" align="center"><%=emp_name%></td>
                              <th height="30" align="left" style=" border-top:1px solid #e3e3e3;">④주민등록번호(또는 외국인등록번호)</th>
                              <td height="30" align="center"><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th height="30" align="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">⑤주소</th>
                              <td colspan="3" height="30" align="left"><%=addr_name%><br>(전화번호:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
                            </tr>
                            <tr>
                              <th height="30" align="left" style="border-left:1px solid #e3e3e3;">⑥사업장 소재지</th>
                              <td colspan="3" height="30" align="left"><%=addr_name%><br>(전화번호:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
						    </tr>
						</thead>
				  </table>
					<table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="10%" >
                              <col width="12%" >
							  <col width="8%" >
							  <col width="8%" >
							  <col width="*" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="12%" >
                              <col width="12%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <td colspan="9" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">2. 월세액 소득공제 명세</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col" height="30" align="center">임대인성명<br>(상호)</th>
                                <th rowspan="2" scope="col" height="30" align="center">주민등록번호<br>(사업자번호)</th>
                                <th rowspan="2" scope="col" height="30" align="center">주택유형</th>
                                <th rowspan="2" scope="col" height="30" align="center">주택계약<br>면적(㎡)</th>
                                <th rowspan="2"scope="col" height="30" align="center">임대차계약서 상 주소지</th>
                                <th colspan="2" scope="col" height="30" align="center" style="border-bottom:1px solid #e3e3e3;">계약서상<br>임대차 계약기간</th>
                                <th rowspan="2" scope="col" height="30" align="center">연간 월세액(원)</th>
                                <th rowspan="2" scope="col" height="30" align="center">공제금액(원)</th>
                              </tr>
                              <tr>
                                <th scope="col" height="30" align="center" style="border-left:1px solid #e3e3e3;">개시일</th>
                                <th scope="col" height="30" align="center">종료일</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						do until rs.eof
                             
							 if rs("s_type") = "퇴직연금소득공제" then
	           		 %>
							<tr>
                                <td height="30" align="center"><%=rs("s_type")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_name")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td align="left"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
					<%
							end if
							rs.movenext()
						loop
						rs.close()
						
					%>
                            <tr>
                                <td colspan="9" height="30" align="left" scope="col" style="border-bottom:1px solid #e3e3e3;">※ 주택유형 구분코드 - 단독주택:1, 다가구:2, 다세대주택:3, 연립주택:4, 아파트:5, 오피스텔:6, 기타:7<br><br>
                                ※ 계약서상 임대차계약기간 - 게시일과 종료일은 예시와 같이 기재(예시) 2013.01.01.</td>
                            </tr>
						</tbody>
					</table>
                    
                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="10%" >
                              <col width="12%" >
							  <col width="*" >
							  <col width="8%" >
                              <col width="12%" >
                              <col width="12%" >
                              <col width="10%" >
                              <col width="14%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <td colspan="8" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">3 거주자 간 주택임차차입금 원리금 산환액 소득공제 명세</td>
                              </tr>
                              <tr>
                                <td colspan="8" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">1) 금전소비대차 계약내용</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col" height="30" align="center">대주</th>
                                <th rowspan="2" scope="col" height="30" align="center">주민등록번호</th>
                                <th rowspan="2" scope="col" height="30" align="center">금전소비대차<br>계약기간</th>
                                <th rowspan="2" scope="col" height="30" align="center">차입금<br>이자율</th>
                                <th colspan="3" scope="col" height="30" align="center" style="border-bottom:1px solid #e3e3e3;">원리금 산환액</th>
                                <th rowspan="2" scope="col" height="30" align="center">공제금액(원)</th>
                              </tr>
                              <tr>
                                <th scope="col" height="30" align="center" style="border-left:1px solid #e3e3e3;">계</th>
                                <th scope="col" height="30" align="center">원금</th>
                                <th scope="col" height="30" align="center">이자</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						s_id = "연금저축"
						sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
                        Rs.Open Sql, Dbconn, 1
						
						do until rs.eof
                             
							 if rs("s_type") <> "퇴직연금소득공제" then
	           			%>
							<tr>
                                <td height="30" align="center"><%=rs("s_type")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							end if
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>                    

                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="10%" >
                              <col width="12%" >
                              <col width="8%" >
                              <col width="8%" >
							  <col width="*" >
                              <col width="10%" >
                              <col width="10%" >
                              <col width="14%" >
                        </colgroup>
						 <thead>
                              <tr>
                                <td colspan="8" height="30" align="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">2) 임대차 계약내용</td>
                              </tr>
                              <tr>
                                <th rowspan="2" class="first" scope="col" height="30" align="center">임대인성명<br>(상호)</th>
                                <th rowspan="2" scope="col" height="30" align="center">주민등록번호<br>(사업자번호)</th>
                                <th rowspan="2" scope="col" height="30" align="center">주택유형</th>
                                <th rowspan="2" scope="col" height="30" align="center">주택계약<br>면적(㎡)</th>
                                <th rowspan="2"scope="col" height="30" align="center">임대차계약서 상 주소지</th>
                                <th colspan="2" scope="col" height="30" align="center" style="border-bottom:1px solid #e3e3e3;">계약서상<br>임대차 계약기간</th>
                                <th rowspan="2" scope="col" height="30" align="center">전세보증금(원)</th>
                              </tr>
                              <tr>
                                <th scope="col" height="30" align="center" style="border-left:1px solid #e3e3e3;">개시일</th>
                                <th scope="col" height="30" align="center">종료일</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						s_id = "주택마련저축"
						sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
                        Rs.Open Sql, Dbconn, 1
						
						do until rs.eof
                             
	           			%>
							<tr>
                                <td height="30" align="center"><%=rs("s_type")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_name")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td height="30" align="left"><%=rs("s_account_no")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td height="30" align="center"><%=rs("s_bank_code")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
                            <tr>
                                <td colspan="8" height="30" align="left" scope="col" style="border-bottom:1px solid #e3e3e3;">※ 주택유형 구분코드 - 단독주택:1, 다가구:2, 다세대주택:3, 연립주택:4, 아파트:5, 오피스텔:6, 기타:7<br><br>
                                ※ 계약서상 임대차계약기간 - 게시일과 종료일은 예시와 같이 기재(예시) 2013.01.01.</td>
                            </tr>
						</tbody>
					</table>   

                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="100%" >
                        </colgroup>
						 <thead>
                              <tr>
								<td scope="col" height="30" align="center" style=" border-bottom:2px solid #515254;">작 성 방 법</td>
							  </tr>
                              <tr>
								<td scope="col" height="30" align="left" >
                                1. 월세액 공제나 거주자 간 주택임차자금 차임금 원리금 상환액 공제를 받는 근로소득자에 대해서는 해당 소득공제에 대한 명세를 작성해야 합니다.<br><br>
                                2. 해당 임대차 계약별로 연간 합계인 월세액·원리금상환액과 공제금액을 적으며, 공제금액이 0인경우에는 적지 않습니다.<br><br>
                                3. 주택유형은 단독주택, 다가구, 다세대주택, 연립주택, 아파트, 오피스텔, 기타 중에서 해당되는 주택유형 구분코드를 적습니다.<br><br>
                                4. 전세보증금은 과세기간 종료일(12.31.) 현재의 전세보증금을 적습니다.</td>
                             </tr>
                        </thead>
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

