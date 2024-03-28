<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim family_tab(10,5)

u_type = request("u_type")
d_year = request("d_year")
d_emp_no = request("d_emp_no")
d_person_no = request("d_person_no")
d_emp_name = request("d_emp_name")
d_seq = request("d_seq")

d_person_no1 = mid(cstr(d_person_no),1,6)
d_person_no2 = mid(cstr(d_person_no),7,7)

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

sql = "select * from pay_yeartax_family where f_year = '"&d_year&"' and f_emp_no = '"&d_emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
'Set rs_fami = DbConn.Execute(SQL)
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

title_line = " 기부금 세부항목 입력 "
if u_type = "U" then

	Sql="select * from pay_yeartax_donation where d_year = '"&d_year&"' and d_emp_no = '"&d_emp_no&"' and d_person_no = '"&d_person_no&"' and d_seq = '"&d_seq&"'"
	Set rs=DbConn.Execute(Sql)

	d_rel = rs("d_rel")
    d_name = rs("d_name")
	d_trade_no = rs("d_trade_no")
	d_trade_name = rs("d_trade_name")
	d_nts_chk = rs("d_nts_chk")
	d_data_gubun = rs("d_data_gubun")
	d_cnt = rs("d_cnt")
	d_amt = rs("d_amt")

	rs.close()

	title_line = " 기부금 세부항목 변경  "
	
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
												$( "#datepicker" ).datepicker("setDate", "<%=b_from_date%>" );
			});	
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=b_to_date%>" );
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
				if(document.frm.d_data_gubun.value =="") {
					alert('기부코드(유형)을 선택하세요');
					frm.d_data_gubun.focus();
					return false;}
				if(document.frm.d_family.value =="") {
					alert('대상자를 선택하세요');
					frm.d_family.focus();
					return false;}
				if(document.frm.d_amt.value ==0) {
					alert('금액을 입력하세요');
					frm.d_amt.focus();
					return false;}
				if(document.frm.d_trade_name.value =="") {
					alert('기부처명을 입력하세요');
					frm.d_trade_name.focus();
					return false;}
				if(document.frm.d_data_gubun.value == "정치자금기부금") {
					if(document.frm.d_rel.value != "본인") {
							alert('정치자금기부는 본인만 가능합니다');
							frm.d_data_gubun.focus();
							return false;}}
				if(document.frm.d_data_gubun.value == "우리사주조합기부금") {
					if(document.frm.d_rel.value != "본인") {
							alert('우리사주조합기부는 본인만 가능합니다');
							frm.d_data_gubun.focus();
							return false;}}
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			} 
			
			function num_chk(txtObj){
				dd_cnt = parseInt(document.frm.d_cnt.value.replace(/,/g,""));	
				dd_amt = parseInt(document.frm.d_amt.value.replace(/,/g,""));	
		
				dd_cnt = String(dd_cnt);
				num_len = dd_cnt.length;
				sil_len = num_len;
				dd_cnt = String(dd_cnt);
				if (dd_cnt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) dd_cnt = dd_cnt.substr(0,num_len -3) + "," + dd_cnt.substr(num_len -3,3);
				if (sil_len > 6) dd_cnt = dd_cnt.substr(0,num_len -6) + "," + dd_cnt.substr(num_len -6,3) + "," + dd_cnt.substr(num_len -2,3);
				document.frm.d_cnt.value = dd_cnt;
				
				dd_amt = String(dd_amt);
				num_len = dd_amt.length;
				sil_len = num_len;
				dd_amt = String(dd_amt);
				if (dd_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) dd_amt = dd_amt.substr(0,num_len -3) + "," + dd_amt.substr(num_len -3,3);
				if (sil_len > 6) dd_amt = dd_amt.substr(0,num_len -6) + "," + dd_amt.substr(num_len -6,3) + "," + dd_amt.substr(num_len -2,3);
				document.frm.d_amt.value = dd_amt;
			}		
			
			 function setaddr() {
			 var srt = document.frm.d_family.value;
//			 alert(srt);
			 var arr = srt.split(','); 
			 var sub_string = arr[arr.length-3]; 
			 var sub_temp1 = sub_string.substring(0,6); 
			 var sub_temp2 = sub_string.substring(6,13); 
//             alert(sub_temp1);
//			 alert(sub_temp2);
			 document.frm.d_person_no.value = arr[arr.length-3];
			 document.frm.d_person_no1.value = sub_temp1;
			 document.frm.d_person_no2.value = sub_temp2;
			 document.frm.d_name.value = arr[arr.length-4];
			 document.frm.d_rel.value = arr[arr.length-5];
//			 alert(arr[arr.length-2]);
//			 document.frm.d_disab.value = arr[arr.length-2];
//			 document.frm.d_age65.value = arr[arr.length-1];
             }

			
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_donation_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="10%" >
						<col width="15%" >
						<col width="10%" >
						<col width="15%" >
                        <col width="10%" >
						<col width="15%" >
                        <col width="10%" >
						<col width="15%" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="d_emp_no" type="text" id="d_emp_no" size="10" value="<%=d_emp_no%>" readonly="true">
                      <input type="hidden" name="d_year" value="<%=d_year%>" ID="d_year">
                      <input type="hidden" name="d_seq" value="<%=d_seq%>" ID="d_seq"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="d_emp_name" type="text" id="d_emp_name" size="10" value="<%=d_emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th>기부코드<br>(유형)</th>
					  <td colspan="7" class="left">
					  <select name="d_data_gubun" id="d_data_gubun" value="<%=d_data_gubun%>" style="width:150px">
				          <option value="" <% if d_data_gubun = "" then %>selected<% end if %>>선택</option>
				          <option value='정치자금기부금' <%If d_data_gubun = "정치자금기부금" then %>selected<% end if %>>정치자금기부금</option>
				          <option value='법정기부금' <%If d_data_gubun = "법정기부금" then %>selected<% end if %>>법정기부금</option>
				          <option value='우리사주조합기부금' <%If d_data_gubun = "우리사주조합기부금" then %>selected<% end if %>>우리사주조합기부금</option>
                          <option value='종교단체외지정기부금' <%If d_data_gubun = "종교단체외지정기부금" then %>selected<% end if %>>종교단체외지정기부금</option>
                          <option value='종교단체지정기부금' <%If d_data_gubun = "종교단체지정기부금" then %>selected<% end if %>>종교단체지정기부금</option>
                      </select>
                      </td>
                    </tr>
                 	<tr>
                      <th>대상자</th>
                      <td colspan="3" class="left">
					   <select name="d_family" id="d_family" style="width:90px" onChange="setaddr();">
                          <option value="" <% if d_name = "" then %>selected<% end if %>>선택</option>
                  <% 
						for i = 1 to 10
						    if family_tab(i,2) = "" or isnull(family_tab(i,2)) then 
			                           exit for
		                       else
			  	  %>
                		  <option value='<%=family_tab(i,1)%>,<%=family_tab(i,2)%>,<%=family_tab(i,3)%>,<%=family_tab(i,4)%>,<%=family_tab(i,5)%>' <%If d_name = family_tab(i,2) then %>selected<% end if %>><%=family_tab(i,2)%></option>
                  <%
				            end if
						next
				  %>
            		  </select>
                      <th>관계/<br>주민등록번호</th>
					  <td colspan="3" class="left">
                      <input name="d_name" type="hidden" value="<%=d_name%>" readonly="true" style="width:70px">
                      <input name="d_rel" type="text" value="<%=d_rel%>" readonly="true" style="width:60px">
                      <input name="d_person_no1" type="text" value="<%=d_person_no1%>" readonly="true" style="width:50px;text-align:center">
                      -
                      <input name="d_person_no2" type="text" value="<%=d_person_no2%>" readonly="true" style="width:60px;text-align:center">
                      <input name="d_person_no" type="hidden" value="<%=d_person_no%>" readonly="true" style="width:130px">
                      </td>
                      </td>
                    </tr>
                    </tr>
                    <tr>
                      <th>사업자(주민)<br>번호</th>
                      <td class="left">
                      <input name="d_trade_no" type="text" value="<%=d_trade_no%>" style="width:90px" id="d_trade_no"></td>
                      <th>기부처명</th>
                      <td class="left">
                      <input name="d_trade_name" type="text" value="<%=d_trade_name%>" style="width:100px" id="d_trade_name"></td>
                      <th>건수</th>
					  <td class="left">
                      <input name="d_cnt" type="text" id="d_cnt" style="width:90px;text-align:right" value="<%=formatnumber(d_cnt,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>금액</th>
					  <td class="left">
                      <input name="d_amt" type="text" id="d_amt" style="width:90px;text-align:right" value="<%=formatnumber(d_amt,0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <th>국세청<br>자료여부</th>
                      <td colspan="7" class="left">
					  <input type="checkbox" name="d_nts_chk" value="Y" <% if d_nts_chk = "Y" then %>checked<% end if %> id="d_nts_chk">예
					  </td>
                    </tr>
                    <tr>
                      <td colspan="8" class="left">※ 사업자번호 : 해당기부단체의 사업자번호, 사업자번호가 없는 경우 기부단체장의 주민번호 입력<br>
                &nbsp;&nbsp;&nbsp;&nbsp;단, 정당등에 기부한 정치자의 경우 사업자번호가 없으므로 입력하지 않아도 무방.<br>
                ※ 사업자번호는 - 를 빼고 숫자만 입력.</td>
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

