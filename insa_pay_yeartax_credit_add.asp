<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim family_tab(10,5)

u_type = request("u_type")
c_year = request("c_year")
c_emp_no = request("c_emp_no")
c_person_no = request("c_person_no")
c_emp_name = request("c_emp_name")
c_id = request("c_id")
c_seq = request("c_seq")

c_person_no1 = mid(cstr(c_person_no),1,6)
c_person_no2 = mid(cstr(c_person_no),7,7)

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

sql = "select * from pay_yeartax_family where f_year = '"&c_year&"' and f_emp_no = '"&c_emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
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

title_line = c_id + " 세부항목 입력 "
if u_type = "U" then

	Sql="select * from pay_yeartax_credit where c_year = '"&c_year&"' and c_emp_no = '"&c_emp_no&"' and c_person_no = '"&c_person_no&"' and c_id = '"&c_id&"' and c_seq = '"&c_seq&"'"
	Set rs=DbConn.Execute(Sql)

	c_rel = rs("c_rel")
    cc_name = rs("cc_name")
    c_market = rs("c_market")
	c_transit = rs("c_transit")
	c_nts_amt = rs("c_nts_amt")
	c_other_amt = rs("c_other_amt")

	rs.close()

	title_line = c_id + " 세부항목 변경 "
	
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
				if(document.frm.c_family.value =="") {
					alert('대상자를 선택하세요');
					frm.c_family.focus();
					return false;}
				if(document.frm.c_id.value == "현금영수증") {
					if(document.frm.c_other_amt.value != 0) {
							alert('현금영수증은 국세청자료만 가능합니다');
							frm.c_nts_amt.focus();
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
				nts_amt = parseInt(document.frm.c_nts_amt.value.replace(/,/g,""));	
				other_amt = parseInt(document.frm.c_other_amt.value.replace(/,/g,""));	
		
				nts_amt = String(nts_amt);
				num_len = nts_amt.length;
				sil_len = num_len;
				nts_amt = String(nts_amt);
				if (nts_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) nts_amt = nts_amt.substr(0,num_len -3) + "," + nts_amt.substr(num_len -3,3);
				if (sil_len > 6) nts_amt = nts_amt.substr(0,num_len -6) + "," + nts_amt.substr(num_len -6,3) + "," + nts_amt.substr(num_len -2,3);
				document.frm.c_nts_amt.value = nts_amt;
				
				other_amt = String(other_amt);
				num_len = other_amt.length;
				sil_len = num_len;
				other_amt = String(other_amt);
				if (other_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) other_amt = other_amt.substr(0,num_len -3) + "," + other_amt.substr(num_len -3,3);
				if (sil_len > 6) other_amt = other_amt.substr(0,num_len -6) + "," + other_amt.substr(num_len -6,3) + "," + other_amt.substr(num_len -2,3);
				document.frm.c_other_amt.value = other_amt;
			}		
			
			 function setaddr() {
			 var srt = document.frm.c_family.value;
//			 alert(srt);
			 var arr = srt.split(','); 
			 var sub_string = arr[arr.length-3]; 
			 var sub_temp1 = sub_string.substring(0,6); 
			 var sub_temp2 = sub_string.substring(6,13); 
//             alert(sub_temp1);
//			 alert(sub_temp2);
			 document.frm.c_person_no.value = arr[arr.length-3];
			 document.frm.c_person_no1.value = sub_temp1;
			 document.frm.c_person_no2.value = sub_temp2;
			 document.frm.cc_name.value = arr[arr.length-4];
			 document.frm.c_rel.value = arr[arr.length-5];
//			 alert(arr[arr.length-2]);
//			 document.frm.e_disab.value = arr[arr.length-2];
//			 document.frm.e_age65.value = arr[arr.length-1];
             }

			
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_credit_save.asp" method="post" name="frm">
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
					  <input name="c_emp_no" type="text" id="c_emp_no" size="10" value="<%=c_emp_no%>" readonly="true">
                      <input type="hidden" name="c_year" value="<%=c_year%>" ID="m_year">
                      <input type="hidden" name="c_seq" value="<%=c_seq%>" ID="m_seq"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="c_emp_name" type="text" id="c_emp_name" size="10" value="<%=c_emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>대상자</th>
                      <td colspan="3" class="left">
					   <select name="c_family" id="c_family" style="width:90px" onChange="setaddr();">
                          <option value="" <% if cc_name = "" then %>selected<% end if %>>선택</option>
                  <% 
						for i = 1 to 10
						    if family_tab(i,2) = "" or isnull(family_tab(i,2)) then 
			                           exit for
		                       else
			  	  %>
                		  <option value='<%=family_tab(i,1)%>,<%=family_tab(i,2)%>,<%=family_tab(i,3)%>,<%=family_tab(i,4)%>,<%=family_tab(i,5)%>' <%If cc_name = family_tab(i,2) then %>selected<% end if %>><%=family_tab(i,2)%></option>
                  <%
				            end if
						next
				  %>
            		  </select>
                      <th>관계/<br>주민등록번호</th>
					  <td colspan="3" class="left">
                      <input name="cc_name" type="hidden" value="<%=c_name%>" readonly="true" style="width:70px">
                      <input name="c_rel" type="text" value="<%=c_rel%>" readonly="true" style="width:60px">
                      <input name="c_person_no1" type="text" value="<%=c_person_no1%>" readonly="true" style="width:50px;text-align:center">
                      -
                      <input name="c_person_no2" type="text" value="<%=c_person_no2%>" readonly="true" style="width:60px;text-align:center">
                      <input name="c_person_no" type="hidden" value="<%=c_person_no%>" readonly="true" style="width:130px">
                      </td>
                      </td>
                    </tr>
                    </tr>
                    <tr>
                      <th>전통시장</th>
                      <td colspan="3" class="left">
					  <input type="checkbox" name="c_market" value="Y" <% if c_market = "Y" then %>checked<% end if %> id="c_market">예
					  </td>
                      <th>대중교통</th>
                      <td colspan="3" class="left">
					  <input type="checkbox" name="c_transit" value="Y" <% if c_transit = "Y" then %>checked<% end if %> id="c_transit">예
					  </td>
                    </tr>
                    <tr>
                      <th>국세청금액</th>
					  <td colspan="3" class="left">
                      <input name="c_nts_amt" type="text" id="c_nts_amt" style="width:90px;text-align:right" value="<%=formatnumber(c_nts_amt,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>그밖의금액</th>
					  <td colspan="3" class="left">
                      <input name="c_other_amt" type="text" id="c_other_amt" style="width:90px;text-align:right" value="<%=formatnumber(c_other_amt,0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <td colspan="8" class="left">※ 형제,자매의 신용카드등의 사용금액은 공제대상이 아닙니다. 절대 입력하지 마세요<br>
                ※ 신용카드,직불카드,현금영수증 사용분중 전통시장 및 대중교통에 해당하는 금액은 전통시장 또는 대중교통에 체크하고 입력<br>
                &nbsp;&nbsp;&nbsp;&nbsp;신용카드 사용액이 인별합계 5200만원중 전통시장사용분 200만원이라면 200만원은 전통시장에 체크하고, 나머지 5000만원은 따로 입력을 함.<br>
                ※ 국세청에서 발급받은 것은 국세청금엑에 카드사등에서 발급받은 것은 그밖의 금액에 입력.<br>
                ※ 현금영수증은 국세청자료에만 입력 가능.</td>
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
                <input type="hidden" name="c_id" value="<%=c_id%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

