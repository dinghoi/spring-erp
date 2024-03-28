<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

inc_yyyy = cint(mid(now(),1,4)) - 1

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

s_id = "연금저축"

sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
Rs.Open Sql, Dbconn, 1


title_line = "연말정산-연금·저축 등 소득·세액 공제명세서"
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
			function getPageCode(){
				return "3 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.inc_yyyy.value == "") {
					alert ("귀속년도를 입력하세요.");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_saving_report.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
                            <col width="10%" >
							<col width="30%" >
							<col width="20%" >
							<col width="30%" >
						</colgroup>
						<thead>
                            <tr>
                              <td rowspan="4">1. 인적 사항</td>
                              <th class="left">①법인명</th>
                              <td><%=company_name%></td>
                              <th class="left">②업체명</th>
                              <td>&nbsp;</td>
						    </tr>
                            <tr>
							  <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">③성명</th>
                              <td><%=emp_name%></td>
                              <th class="left" style=" border-top:1px solid #e3e3e3;">④주민등록번호(또는 외국인등록번호)</th>
                              <td><%=emp_person%></td>
						    </tr>
                            <tr>
							  <th class="left" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">⑤주소</th>
                              <td colspan="3" class="left"><%=addr_name%><br>(전화번호:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
                            </tr>
                            <tr>
                              <th class="left" style="border-left:1px solid #e3e3e3;">⑥사업장 소재지</th>
                              <td colspan="3" class="left"><%=addr_name%><br>(전화번호:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
						    </tr>
						</thead>
					</table>

                    <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="10%" >
                              <col width="20%" >
							  <col width="20%" >
							  <col width="20%" >
							  <col width="20%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <td colspan="5" class="left" scope="col">2. 연금계좌 세액공제</td>
                              </tr>
                              <tr>
                                <td colspan="5" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">1) 퇴직연금계좌<br>* 퇴직연금계좌에 대한 명세를 작성합니다.</td>
                              </tr>
                              <tr>
                                <th class="first" scope="col">퇴직연금 구분</th>
                                <th scope="col">금융회사 등</th>
                                <th scope="col">계좌번호(또는 증권번호)</th>
                                <th scope="col">납입금액</th>
                                <th scope="col">세액공제금액</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
                             if rs("s_type") = "퇴직연금소득공제" then
	           			%>
							<tr>
                                <td><%=rs("s_type")%>&nbsp;</td>
                                <td><%=rs("s_bank_name")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
						    end if
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
                </table>      

                <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="10%" >
                              <col width="20%" >
							  <col width="20%" >
							  <col width="20%" >
							  <col width="20%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <td colspan="5" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">2) 연금저축계좌<br>* 연금저축계좌에 대한 명세를 작성합니다.</td>
                              </tr>
                              <tr>
                                <th class="first" scope="col">연금저축 구분</th>
                                <th scope="col">금융회사 등</th>
                                <th scope="col">계좌번호(또는 증권번호)</th>
                                <th scope="col">납입금액</th>
                                <th scope="col">세액공제금액</th>
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
                                <td><%=rs("s_type")%>&nbsp;</td>
                                <td><%=rs("s_bank_name")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							end if
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
                </table>   
                
                <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="10%" >
                              <col width="20%" >
							  <col width="20%" >
							  <col width="20%" >
							  <col width="20%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <td colspan="5" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">3. 주택마련저축 소득공제<br>* 주택마련저축 소득공제에 대한 명세를 작성합니다.</td>
                              </tr>
                              <tr>
                                <th class="first" scope="col">저축 구분</th>
                                <th scope="col">금융회사 등</th>
                                <th scope="col">계좌번호(또는 증권번호)</th>
                                <th scope="col">납입금액</th>
                                <th scope="col">세액공제금액</th>
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
                                <td><%=rs("s_type")%>&nbsp;</td>
                                <td><%=rs("s_bank_name")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
              </table>         
                
              <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="10%" >
                              <col width="20%" >
							  <col width="20%" >
							  <col width="20%" >
							  <col width="20%" >
                           </colgroup>
                           <thead>
                              <tr>
                                <td colspan="5" class="left" scope="col" style=" border-bottom:1px solid #e3e3e3;">4. 장기집합투자증권 소득공제<br>* 장기집합투자증권 소득공제에 대한 명세를 작성합니다.</td>
                              </tr>
                              <tr>
                                <th colspan="2" scope="col">금융회사 등</th>
                                <th scope="col">계좌번호(또는 증권번호)</th>
                                <th scope="col">납입금액</th>
                                <th scope="col">세액공제금액</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						s_id = "장기집합투자증권저축"
						sql = "select * from pay_yeartax_saving where s_year = '"&inc_yyyy&"' and s_emp_no = '"&emp_no&"' and s_id = '"&s_id&"' ORDER BY s_emp_no,s_id,s_seq ASC"
                        Rs.Open Sql, Dbconn, 1
						
						do until rs.eof
                             
	           			%>
							<tr>
                                <td colspan="2"><%=rs("s_bank_name")%>&nbsp;</td>
                                <td><%=rs("s_account_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("s_amt"),0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
                </table>                               

                <table cellpadding="0" cellspacing="0" class="tableList">
                        <colgroup>
							   <col width="100%" >
                        </colgroup>
                        <thead>
                            <tr>
								<td scope="col" style=" border-bottom:2px solid #515254;">작 성 방 법</td>
							</tr>
                            <tr>
								<td scope="col" class="left" >
                                1. 연금계좌 세액공제, 주택마렴저축, 장기집합투자증권저축 소득공제를 받는 소득자애 대해서는 해당 소득.세액 공제에 대한 명세를 작성해야 합니다.<br>해당 계좌별로 불입금액과 소득.세액공제금액을 적고, 공제금액이 0인 경우에는 적지 않습니다<br><br>
                                2. 퇴직연금계좌에서 퇴직연금구분란은 퇴직연금(DC, IRP).과학기술인공제회로 구분하여 적습니다.<br><br>
                                3. 연금저축계좌에서 연금저축구분란은 개인연금저축과 연금저축으로 구분하여 적습니다.<br><br>
                                4. 주택마련저축 공제의 자축구분란은 청약저축, 주택청약종합저축 및 근로자주택마련저축으로 구분하여 적습니다.<br><br>
                                5. 공제금액란은 근로소득자가 적지 않을 수 있습니다.</td>
                            </tr>
                        </thead>
                </table>            
                                               
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="insa_pay_yeartax_medical_report.asp" class="btnType04">의료비지급명세서</a>
                    <a href="insa_pay_yeartax_donation_report.asp" class="btnType04">기부금명세서</a>
                    <a href="insa_pay_yeartax_credit_report.asp" class="btnType04">신용카드등 명세서</a>
                    <a href="insa_pay_yeartax_tax_report.asp" class="btnType04">소득공제신고서</a>
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_saving_print.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&inc_yyyy=<%=inc_yyyy%>','yeartax_donation_print_popup','scrollbars=yes,width=1250,height=500')" class="btnType04">연금.저축 등 소득세액 공제명세서 출력</a>
                    <a href="insa_pay_yeartax_house_report.asp" class="btnType04">주택임차차임금원리금상환 명세서</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

