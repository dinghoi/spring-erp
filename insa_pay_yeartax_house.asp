<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim family_tab(10,3)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_house.asp"

y_final=Request("y_final")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

inc_yyyy = cint(mid(now(),1,4)) - 1

for i = 1 to 10
    family_tab(i,1) = ""
	family_tab(i,2) = ""
	family_tab(i,3) = ""
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_hous = Server.CreateObject("ADODB.Recordset")
Set rs_houm = Server.CreateObject("ADODB.Recordset")
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
emp_person = cstr(rs_emp("emp_person1")) + cstr(rs_emp("emp_person2"))	
rs_emp.close()	

Sql = "select * from pay_yeartax_house where h_year = '"&inc_yyyy&"' and h_emp_no = '"&emp_no&"'"
rs_hous.Open Sql, Dbconn, 1
Set rs_hous = DbConn.Execute(SQL)
if not rs_hous.eof then
       u_type = "U"
       h_lender_amt = rs_hous("h_lender_amt")
	   h_person_amt = rs_hous("h_person_amt")
	   h_long15_amt = rs_hous("h_long15_amt")
	   h_long29_amt = rs_hous("h_long29_amt")
	   h_long30_amt = rs_hous("h_long30_amt")
	   h_fixed_amt = rs_hous("h_fixed_amt")
	   h_other_amt = rs_hous("h_other_amt")
   else
       u_type = ""
       h_lender_amt = 0
	   h_person_amt = 0
	   h_long15_amt = 0
	   h_long29_amt = 0
	   h_long30_amt = 0
	   h_fixed_amt = 0
	   h_other_amt = 0
end if
rs_hous.close()	

h_month_amt = 0
Sql = "select * from pay_yeartax_house_m where hm_year = '"&inc_yyyy&"' and hm_emp_no = '"&emp_no&"' ORDER BY hm_emp_no,hm_seq ASC"
rs_houm.Open Sql, Dbconn, 1
Set rs_houm = DbConn.Execute(SQL)
do until rs_houm.eof
       h_month_amt = h_month_amt + rs_houm("hm_month_amt")
	rs_houm.MoveNext()
loop
rs_houm.close()

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
Set rs_fami = DbConn.Execute(SQL)
i = 0
do until rs_fami.eof
   if rs_fami("f_rel") = "본인" or rs_fami("f_wife") = "Y" or rs_fami("f_age20") = "Y" or rs_fami("f_age60") = "Y" or rs_fami("f_old") = "Y" then
		  i = i + 1
		  family_tab(i,1) = rs_fami("f_rel")
	      family_tab(i,2) = rs_fami("f_family_name")
	      family_tab(i,3) = rs_fami("f_person_no")
	end if
	rs_fami.MoveNext()
loop
rs_fami.close()

sql = "select * from pay_yeartax_house_m where hm_year = '"&inc_yyyy&"' and hm_emp_no = '"&emp_no&"' ORDER BY hm_emp_no,hm_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산 - 특별공제(주택자금) "
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}		
			
			function chkfrm() {
//				if(document.frm.emp_ename.value =="") {
//					alert('영문성명을 입력하세요');
//					frm.emp_ename.focus();
//					return false;}
					
				a=confirm('등록하시겠습니까?');
				if (a==true) {
					return true;
				}
				return false;
			} 
			
			function num_chk(txtObj){
				lender_amt = parseInt(document.frm.h_lender_amt.value.replace(/,/g,""));	
				person_amt = parseInt(document.frm.h_person_amt.value.replace(/,/g,""));	
				long15_amt = parseInt(document.frm.h_long15_amt.value.replace(/,/g,""));	
				long29_amt = parseInt(document.frm.h_long29_amt.value.replace(/,/g,""));	
				long30_amt = parseInt(document.frm.h_long30_amt.value.replace(/,/g,""));	
				fixed_amt = parseInt(document.frm.h_fixed_amt.value.replace(/,/g,""));	
				other_amt = parseInt(document.frm.h_other_amt.value.replace(/,/g,""));	
		
				lender_amt = String(lender_amt);
				num_len = lender_amt.length;
				sil_len = num_len;
				lender_amt = String(lender_amt);
				if (lender_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) lender_amt = lender_amt.substr(0,num_len -3) + "," + lender_amt.substr(num_len -3,3);
				if (sil_len > 6) lender_amt = lender_amt.substr(0,num_len -6) + "," + lender_amt.substr(num_len -6,3) + "," + lender_amt.substr(num_len -2,3);
				document.frm.h_lender_amt.value = lender_amt;
				
				person_amt = String(person_amt);
				num_len = person_amt.length;
				sil_len = num_len;
				person_amt = String(person_amt);
				if (person_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) person_amt = person_amt.substr(0,num_len -3) + "," + person_amt.substr(num_len -3,3);
				if (sil_len > 6) person_amt = person_amt.substr(0,num_len -6) + "," + person_amt.substr(num_len -6,3) + "," + person_amt.substr(num_len -2,3);
				document.frm.h_person_amt.value = person_amt;
				
				long15_amt = String(long15_amt);
				num_len = long15_amt.length;
				sil_len = num_len;
				long15_amt = String(long15_amt);
				if (long15_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) long15_amt = long15_amt.substr(0,num_len -3) + "," + long15_amt.substr(num_len -3,3);
				if (sil_len > 6) long15_amt = long15_amt.substr(0,num_len -6) + "," + long15_amt.substr(num_len -6,3) + "," + long15_amt.substr(num_len -2,3);
				document.frm.h_long15_amt.value = long15_amt;
				
				long29_amt = String(long29_amt);
				num_len = long29_amt.length;
				sil_len = num_len;
				long29_amt = String(long29_amt);
				if (long29_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) long29_amt = long29_amt.substr(0,num_len -3) + "," + long29_amt.substr(num_len -3,3);
				if (sil_len > 6) long29_amt = long29_amt.substr(0,num_len -6) + "," + long29_amt.substr(num_len -6,3) + "," + long29_amt.substr(num_len -2,3);
				document.frm.h_long29_amt.value = long29_amt;
				
				long30_amt = String(long30_amt);
				num_len = long30_amt.length;
				sil_len = num_len;
				long30_amt = String(long30_amt);
				if (long30_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) long30_amt = long30_amt.substr(0,num_len -3) + "," + long30_amt.substr(num_len -3,3);
				if (sil_len > 6) long30_amt = long30_amt.substr(0,num_len -6) + "," + long30_amt.substr(num_len -6,3) + "," + long30_amt.substr(num_len -2,3);
				document.frm.h_long30_amt.value = long30_amt;
				
				fixed_amt = String(fixed_amt);
				num_len = fixed_amt.length;
				sil_len = num_len;
				fixed_amt = String(fixed_amt);
				if (fixed_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) fixed_amt = fixed_amt.substr(0,num_len -3) + "," + fixed_amt.substr(num_len -3,3);
				if (sil_len > 6) fixed_amt = fixed_amt.substr(0,num_len -6) + "," + fixed_amt.substr(num_len -6,3) + "," + fixed_amt.substr(num_len -2,3);
				document.frm.h_fixed_amt.value = fixed_amt;
				
				other_amt = String(other_amt);
				num_len = other_amt.length;
				sil_len = num_len;
				other_amt = String(other_amt);
				if (other_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) other_amt = other_amt.substr(0,num_len -3) + "," + other_amt.substr(num_len -3,3);
				if (sil_len > 6) other_amt = other_amt.substr(0,num_len -6) + "," + other_amt.substr(num_len -6,3) + "," + other_amt.substr(num_len -2,3);
				document.frm.h_other_amt.value = other_amt;
			}		
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_house_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="10%" >
							<col width="8%" >
                            <col width="8%" >
							<col width="8%" >
                            <col width="9%" >
                            <col width="24%" >
                            <col width="25%" >
						</colgroup>
						<thead>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">성명(<%=emp_no%><input name="emp_no" type="hidden" value="<%=emp_no%>" style="width:40px" readonly="true">)</th>
							  <td colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_name%>
                                <input name="emp_name" type="hidden" value="<%=emp_name%>" style="width:50px" readonly="true">
                                (입사일:<%=emp_in_date%>
                                <input name="emp_in_date" type="hidden" value="<%=emp_in_date%>" style="width:70px" readonly="true">)
                              </td>
							  <th style=" border-bottom:1px solid #e3e3e3;">소속<input name="emp_company" type="hidden" value="<%=emp_company%>" style="width:90px" readonly="true"></th>
							  <td colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_company%> - <%=emp_org_name%>
                                <input name="emp_org_name" type="hidden" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                - <%=emp_grade%>
                                <input name="emp_grade" type="hidden" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                - <%=emp_position%>
                                <input name="emp_position" type="hidden" value="<%=emp_position%>" style="width:70px" readonly="true">
                                (귀속년도:
                                <input name="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:40px; text-align:center" readonly="true">)
                              </td>
						    </tr>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">구분</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3;">지출명세</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">지출구분</th>
                              <th>금액</th>
                              <th colspan="2">공제요건</th>
						    </tr>
                            <tr>
							  <th rowspan="8">주택자금</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3;">주택임차차입금</th>
                              <th colspan="2" style=" border-bottom:1px solid #e3e3e3;">대출기관차입</th>
                              <th rowspan="2" style=" border-bottom:1px solid #e3e3e3;">원리금상환액</th>
                              <td class="right"><input name="h_lender_amt" type="text" id="h_lender_amt" style="width:90px;text-align:right" value="<%=formatnumber(h_lender_amt,0)%>" onKeyUp="num_chk(this);"></td>
                              <td rowspan="2" colspan="2" class="left">※ 금융기관 또는 개인으로부터 국민주택규모의 주택(85㎡이하)을 임차하기 위해 주택전세금 또는 보증금을 차입(3개월이내)하여 원리금을 상환하는 자로서 연말현재 무주택세대의 세대주인 근로소득이어야 함.<br>
                ※ 개인간 차입의 경우에도 총급여액 5천만원 이하만 적용 가능.
                              </td>
						    </tr>
                            <tr>
                              <th colspan="2" style="background:#f8f8f8; border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">개인간차입</th>
                              <td class="right"><input name="h_person_amt" type="text" id="h_person_amt" style="width:90px;text-align:right" value="<%=formatnumber(h_person_amt,0)%>" onKeyUp="num_chk(this);"></td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">월세액</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">지출액</th>
                              <td class="right"><%=formatnumber(h_month_amt,0)%>&nbsp;
                              <input name="h_month_amt" type="hidden" id="h_month_amt" style="width:90px;text-align:right" value="<%=formatnumber(h_month_amt,0)%>" readonly="true">
                              </td>
                              <td colspan="2" class="left">※ 연말 현재 무주택세대의 세대주이어야 함.<br>
                ※ 총급여액이 5천만원 이하면서 세대주의 범위내에는 배우자나 부양가족이 없는 단독세대주도 포함.
                              </td>
						    </tr>
                            <tr>
                              <th rowspan="5" style=" border-left:1px solid #e3e3e3;">장기주택저당차입금</th>
                              <th rowspan="3" style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">2011년 이전<br>차입분</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">15년미만</th>
                              <th rowspan="5" style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">이자상환액</th>
                              <td class="right"><input name="h_long15_amt" type="text" id="h_long15_amt" style="width:90px;text-align:right" value="<%=formatnumber(h_long15_amt,0)%>" onKeyUp="num_chk(this);"></td>
                              <td rowspan="5" colspan="2" class="left">※ 무주택자인 근로자가 국민주택규모 주택을 취득하기 위하여 해당 주택에 저당권을 설정하고 차입한 장기주택저당차입금의 이자를 상환하는 자.
                </td>
						    </tr>
                            <tr>
                              <th style="background:#f8f8f8; border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">15년 ~ 29년</th>
                              <td class="right"><input name="h_long29_amt" type="text" id="h_long29_amt" style="width:90px;text-align:right" value="<%=formatnumber(h_long29_amt,0)%>" onKeyUp="num_chk(this);"></td>
						    </tr>
                            <tr>
                              <th style="background:#f8f8f8; border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">30년</th>
                              <td class="right"><input name="h_long30_amt" type="text" id="h_long30_amt" style="width:90px;text-align:right" value="<%=formatnumber(h_long30_amt,0)%>" onKeyUp="num_chk(this);"></td>
						    </tr>
                            <tr>
                              <th rowspan="2" style="background:#f8f8f8; border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">2012년 이후<br>차입분<br>(15년이상)</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">고정금리.비거치상환대출</th>
                              <td class="right"><input name="h_fixed_amt" type="text" id="h_fixed_amt" style="width:90px;text-align:right" value="<%=formatnumber(h_fixed_amt,0)%>" onKeyUp="num_chk(this);"></td>
						    </tr>
                            <tr>
                              <th style="background:#f8f8f8; border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">기타대출<br>(변동금리.거치식 상황대출)</th>
                              <td class="right"><input name="h_other_amt" type="text" id="h_other_amt" style="width:90px;text-align:right" value="<%=formatnumber(h_other_amt,0)%>" onKeyUp="num_chk(this);"></td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">&nbsp;</h3>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="69%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                           <colgroup>
                              <col width="4%" >
                              <col width="20%" >
                              <col width="20%" >
                              <col width="20%" >
                              <col width="32%" >
                              <col width="4%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">선택</th>
                                <th scope="col">자료유형</th>
                                <th scope="col">임대계약시작일</th>
                                <th scope="col">임대계약종료일</th>
                                <th scope="col">월세액(총액이 아닌 한달치 월세액)</th>
                                <th scope="col">비고</th>
                              </tr>
                            </thead>
                            <tbody>
						<%
						do until rs.eof
	           			%>
							<tr>
                                <td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="Y"></td>
                                <td><%=rs("hm_data_gubun")%>&nbsp;</td>
                                <td><%=rs("hm_from_date")%>&nbsp;</td>
                                <td><%=rs("hm_to_date")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("hm_month_amt"),0)%>&nbsp;</td>
                        <% if y_final <> "Y" then  %>                                  
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_housem_add.asp?hm_year=<%=rs("hm_year")%>&hm_emp_no=<%=rs("hm_emp_no")%>&hm_seq=<%=rs("hm_seq")%>&hm_emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_pay_yeartax_house_add_pop','scrollbars=yes,width=850,height=450')">수정</a></td>
                        <%    else  %>
                                <td>&nbsp;</td>
                        <% end if  %>										
                            </tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  <td width="50%">
                    <div align=center>
                <% if y_final <> "Y" then  %>                    
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                <% end if  %>				    
                    </div>
				  </td>	
                  <td width="50%">
                    <div class="btnRight">
                    <a href="insa_pay_yeartax_insurance.asp" class="btnType04">보험료등록</a>
                    <a href="insa_pay_yeartax_medical.asp" class="btnType04">의료비등록</a>
                    <a href="insa_pay_yeartax_edu.asp" class="btnType04">교육비등록</a>
              <% if y_final <> "Y" then  %>                        
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_housem_add.asp?hm_year=<%=inc_yyyy%>&hm_emp_no=<%=emp_no%>&hm_emp_name=<%=emp_name%>&u_type=<%=""%>','insa_pay_yeartax_house_add_pop','scrollbars=yes,width=850,height=450')" class="btnType04">월세액추가등록</a>
              <%   else  %>
                    <a href="insa_pay_yeartax_house.asp" class="btnType04">주택자금등록</a>
			  <%   end if  %>                           
                    <a href="insa_pay_yeartax_donation.asp" class="btnType04">기부금등록</a>
					</div> 
                  </td>                 
                  </tr>
				</table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="emp_person" value="<%=emp_person%>" ID="Hidden1">                 
			</form>
		</div>				
	</div>        				
	</body>
</html>

