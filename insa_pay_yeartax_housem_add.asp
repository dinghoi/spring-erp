<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim family_tab(10,5)

u_type = request("u_type")
hm_year = request("hm_year")
hm_emp_no = request("hm_emp_no")
hm_person_no = request("hm_person_no")
hm_emp_name = request("hm_emp_name")
hm_seq = request("hm_seq")

hm_person_no1 = mid(cstr(hm_person_no),1,6)
hm_person_no2 = mid(cstr(hm_person_no),7,7)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

for i = 1 to 10
    family_tab(i,1) = ""
	family_tab(i,2) = ""
	family_tab(i,3) = ""
	family_tab(i,4) = ""
	family_tab(i,5) = ""
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from pay_yeartax_family where f_year = '"&e_year&"' and f_emp_no = '"&e_emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
Set rs_fami = DbConn.Execute(SQL)
i = 0
do until rs_fami.eof
   if rs_fami("f_rel") = "본인" or rs_fami("f_wife") = "Y" or rs_fami("f_age20") = "Y" or rs_fami("f_age60") = "Y" or rs_fami("f_old") = "Y" then
		  i = i + 1
		  family_tab(i,1) = rs_fami("f_rel")
	      family_tab(i,2) = rs_fami("f_family_name")
	      family_tab(i,3) = rs_fami("f_person_no")
		  family_tab(i,4) = rs_fami("f_disab")
		  f_birthday = rs_fami("f_birthday")
		  if f_birthday < "1949-12-31" then     
				  family_tab(i,5) = "Y"
			 else
			      family_tab(i,5) = ""	  
		  end if 
	end if
	rs_fami.MoveNext()
loop
rs_fami.close()

title_line = " 월세액 또는 거주자간 주택임차차임금 세부항목 입력 "
if u_type = "U" then

	Sql="select * from pay_yeartax_house_m where hm_year = '"&hm_year&"' and hm_emp_no = '"&hm_emp_no&"' and hm_seq = '"&hm_seq&"'"
	Set rs=DbConn.Execute(Sql)

	hm_from_date = rs("hm_from_date")
    hm_to_date = rs("hm_to_date")
    hm_month_amt = rs("hm_month_amt")
	hm_data_gubun = rs("hm_data_gubun")
	hm_trade_name = rs("hm_trade_name")
    hm_trade_no = rs("hm_trade_no")
    hm_house_type = rs("hm_house_type")
    hm_size = rs("hm_size")
    hm_addr = rs("hm_addr")
    hm_lender = rs("hm_lender")
	hm_lender_person = rs("hm_lender_person")
	hm_lender_from = rs("hm_lender_from")
	hm_lender_to = rs("hm_lender_to")
	hm_lender_rate = rs("hm_lender_rate")
	hm_lender_amt = rs("hm_lender_amt")
	hm_lender_rate_amt = rs("hm_lender_rate_amt")
	hm_hap_amt = hm_lender_amt + hm_lender_rate_amt
	if hm_lender_from = "1900-01-01" then
	   hm_lender_from = ""
	end if
	if hm_lender_to = "1900-01-01" then
	   hm_lender_to = ""
	end if

	rs.close()

	title_line = " 월세액 또는 거주자간 주택임차차임금 세부항목 변경  "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=hm_from_date%>" );
			});	
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=hm_to_date%>" );
			});	
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=hm_lender_from%>" );
			});	
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=hm_lender_to%>" );
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
				if(document.frm.hm_from_date.value =="") {
					alert('계약시작일을 선택하세요');
					frm.hm_from_date.focus();
					return false;}
				if(document.frm.hm_to_date.value =="") {
					alert('계약종료일을 선택하세요');
					frm.hm_to_date.focus();
					return false;}
				if(document.frm.hm_month_amt.value =="") {
					alert('월세액을 입력하세요');
					frm.hm_month_amt.focus();
					return false;}
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			} 
			
			function num_chk(txtObj){
				month_amt = parseInt(document.frm.hm_month_amt.value.replace(/,/g,""));
				lender_amt = parseInt(document.frm.hm_lender_amt.value.replace(/,/g,""));	
				rate_amt = parseInt(document.frm.hm_lender_rate_amt.value.replace(/,/g,""));	
		
				hap_amt = lender_amt + rate_amt;
				
				month_amt = String(month_amt);
				num_len = month_amt.length;
				sil_len = num_len;
				month_amt = String(month_amt);
				if (month_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) month_amt = month_amt.substr(0,num_len -3) + "," + month_amt.substr(num_len -3,3);
				if (sil_len > 6) month_amt = month_amt.substr(0,num_len -6) + "," + month_amt.substr(num_len -6,3) + "," + month_amt.substr(num_len -2,3);
				document.frm.hm_month_amt.value = month_amt;
				
				lender_amt = String(lender_amt);
				num_len = lender_amt.length;
				sil_len = num_len;
				lender_amt = String(lender_amt);
				if (lender_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) lender_amt = lender_amt.substr(0,num_len -3) + "," + lender_amt.substr(num_len -3,3);
				if (sil_len > 6) lender_amt = lender_amt.substr(0,num_len -6) + "," + lender_amt.substr(num_len -6,3) + "," + lender_amt.substr(num_len -2,3);
				document.frm.hm_lender_amt.value = lender_amt;
				
				rate_amt = String(rate_amt);
				num_len = rate_amt.length;
				sil_len = num_len;
				rate_amt = String(rate_amt);
				if (rate_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) rate_amt = rate_amt.substr(0,num_len -3) + "," + rate_amt.substr(num_len -3,3);
				if (sil_len > 6) rate_amt = rate_amt.substr(0,num_len -6) + "," + rate_amt.substr(num_len -6,3) + "," + rate_amt.substr(num_len -2,3);
				document.frm.hm_lender_rate_amt.value = rate_amt;
				
				hap_amt = String(hap_amt);
				num_len = hap_amt.length;
				sil_len = num_len;
				hap_amt = String(hap_amt);
				if (hap_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) hap_amt = hap_amt.substr(0,num_len -3) + "," + hap_amt.substr(num_len -3,3);
				if (sil_len > 6) hap_amt = hap_amt.substr(0,num_len -6) + "," + hap_amt.substr(num_len -6,3) + "," + hap_amt.substr(num_len -2,3);
				document.frm.hm_hap_amt.value = hap_amt;
				
			}		
			
			 function setaddr() {
			 var srt = document.frm.e_family.value;
//			 alert(srt);
			 var arr = srt.split(','); 
			 var sub_string = arr[arr.length-3]; 
			 var sub_temp1 = sub_string.substring(0,6); 
			 var sub_temp2 = sub_string.substring(6,13); 
//             alert(sub_temp1);
//			 alert(sub_temp2);
			 document.frm.e_person_no.value = arr[arr.length-3];
			 document.frm.e_person_no1.value = sub_temp1;
			 document.frm.e_person_no2.value = sub_temp2;
			 document.frm.e_name.value = arr[arr.length-4];
			 document.frm.e_rel.value = arr[arr.length-5];
//			 alert(arr[arr.length-2]);
			 document.frm.e_disab.value = arr[arr.length-2];
//			 document.frm.e_age65.value = arr[arr.length-1];
             }

			
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_housem_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="12%" >
						<col width="13%" >
						<col width="12%" >
						<col width="13%" >
                        <col width="12%" >
						<col width="13%" >
                        <col width="12%" >
						<col width="13%" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="hm_emp_no" type="text" id="hm_emp_no" size="10" value="<%=hm_emp_no%>" readonly="true">
                      <input type="hidden" name="hm_year" value="<%=hm_year%>" ID="m_year">
                      <input type="hidden" name="hm_seq" value="<%=hm_seq%>" ID="m_seq"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="hm_emp_name" type="text" id="hm_emp_name" size="10" value="<%=hm_emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th>자료구분</th>
					  <td colspan="7" class="left">
					  <select name="hm_data_gubun" id="hm_data_gubun" value="<%=hm_data_gubun%>" style="width:120px">
				          <option value="" <% if hm_data_gubun = "" then %>selected<% end if %>>선택</option>
				          <option value='월세액소득공제액' <%If hm_data_gubun = "월세액소득공제액" then %>selected<% end if %>>월세액소득공제액</option>
				          <option value='거주자간주택임차차임금' <%If hm_data_gubun = "거주자간주택임차차임금" then %>selected<% end if %>>거주자간주택임차차임금</option>
                      </select>
                      </td>
                    </tr>
                    <tr>
                      <th>임대인 성명<br>(상호)</th>
					  <td class="left"><input name="hm_trade_name" type="text" id="hm_trade_name" value="<%=hm_trade_name%>" style="width:90px" ></td>
                      <th>주민등록번호<br>(사업자번호)</th>
					  <td class="left"><input name="hm_trade_no" type="text" id="hm_trade_no" value="<%=hm_trade_no%>" style="width:90px" ></td>
                      <th>주택유형</th>
					  <td class="left"><input name="hm_house_type" type="text" id="hm_house_type" value="<%=hm_house_type%>" style="width:90px" ></td>
                      <th>주택계약<br>면적(㎡)</th>
					  <td class="left"><input name="hm_size" type="text" id="hm_size" value="<%=hm_size%>" style="width:70px" ></td>
                    </tr>
                    <tr>
                      <th>주소</th>
					  <td colspan="7" class="left"><input name="hm_addr" type="text" id="hm_addr" value="<%=hm_addr%>" style="width:700px" ></td>
                    </tr>
                    <tr>
                      <th>임대계약<br>시작일</th>
					  <td colspan="3" class="left">
					  <input name="hm_from_date" type="text" size="10" id="datepicker" style="width:70px;" value="<%=hm_from_date%>" readonly="true">
                      </td>
                      <th>임대계약<br>종료일</th>
					  <td colspan="3" class="left">
					  <input name="hm_to_date" type="text" size="10" id="datepicker1" style="width:70px;" value="<%=hm_to_date%>" readonly="true">
                      </td>
                    </tr>
                    <tr>
                      <th>월세액<br>(전세보증금)</th>
					  <td colspan="7" class="left">
                      <input name="hm_month_amt" type="text" id="hm_month_amt" style="width:90px;text-align:right" value="<%=formatnumber(hm_month_amt,0)%>" onKeyUp="num_chk(this);">
                      &nbsp;&nbsp;&nbsp;※ 총액이 아닌 한달치 월세액</td>
                    </tr>
                    <tr>
                      <td colspan="8" class="left">※ 거주자간 주택임차차임금 원리금상환인 경우 아래 항목을 입력 하십시요.</td>
                    </tr>
                    <tr>
                      <th>대주</th>
					  <td class="left"><input name="hm_lender" type="text" id="hm_lender" value="<%=hm_lender%>" style="width:90px" ></td>
                      <th>주민등록번호</th>
					  <td class="left"><input name="hm_lender_person" type="text" id="hm_lender_person" value="<%=hm_lender_person%>" style="width:90px" ></td>
                      <th>금전소비대차<br>계약일</th>
					  <td class="left">
					  <input name="hm_lender_from" type="text" size="10" id="datepicker2" style="width:70px;" value="<%=hm_lender_from%>" readonly="true">
                      </td>
                      <th>금전소비대차<br>종료일</th>
					  <td  class="left">
					  <input name="hm_lender_to" type="text" size="10" id="datepicker3" style="width:70px;" value="<%=hm_lender_to%>" readonly="true">
                      </td>
                    </tr>
                    <tr>
                      <th>차입금<br>이자율</th>
					  <td class="left">
                      <input name="hm_lender_rate" type="text" value="<%=formatnumber(hm_lender_rate,3)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);"></td>
                      <th>원금</th>
					  <td class="left">
                      <input name="hm_lender_amt" type="text" id="hm_lender_amt" style="width:90px;text-align:right" value="<%=formatnumber(hm_lender_amt,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>이자</th>
					  <td class="left">
                      <input name="hm_lender_rate_amt" type="text" id="hm_lender_rate_amt" style="width:90px;text-align:right" value="<%=formatnumber(hm_lender_rate_amt,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>계</th>
					  <td class="left">
                      <input name="hm_hap_amt" type="text" id="hm_hap_amt" style="width:90px;text-align:right" value="<%=formatnumber(hm_hap_amt,0)%>" readonly="true"></td>
                    </tr>
                    
                    <tr>
                      <td colspan="8" class="left">※ 오피스텔 제외, 총급여 5000만원 미만자만 해당<br>
                ※ 월세액공제를 받을경우, 임대인의 소득이 신고되므로 임대인과 상의를 필히 하셔야 합니다.</td>
                    </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

