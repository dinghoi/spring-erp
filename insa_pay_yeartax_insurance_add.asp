<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim family_tab(10,3)

u_type = request("u_type")
i_year = request("i_year")
i_emp_no = request("i_emp_no")
i_person_no = request("i_person_no")
i_emp_name = request("i_emp_name")
i_seq = request("i_seq")

i_person_no1 = mid(cstr(i_person_no),1,6)
i_person_no2 = mid(cstr(i_person_no),7,7)

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
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from pay_yeartax_family where f_year = '"&i_year&"' and f_emp_no = '"&i_emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
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

title_line = " 보험료 세부항목 입력 "
if u_type = "U" then

	Sql="select * from pay_yeartax_insurance where i_year = '"&i_year&"' and i_emp_no = '"&i_emp_no&"' and i_person_no = '"&i_person_no&"' and i_seq = '"&i_seq&"'"
	Set rs=DbConn.Execute(Sql)

	i_rel = rs("i_rel")
    i_name = rs("i_name")
    i_nts_amt = rs("i_nts_amt")
    i_other_amt = rs("i_other_amt")
	i_disab_chk = rs("i_disab_chk")

	rs.close()

	title_line = " 보험료 세부항목 변경  "
	
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
				if(document.frm.i_family.value =="") {
					alert('대상자를 선택하세요');
					frm.i_family.focus();
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
				nts_amt = parseInt(document.frm.i_nts_amt.value.replace(/,/g,""));	
				other_amt = parseInt(document.frm.i_other_amt.value.replace(/,/g,""));	
		
				nts_amt = String(nts_amt);
				num_len = nts_amt.length;
				sil_len = num_len;
				nts_amt = String(nts_amt);
				if (nts_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) nts_amt = nts_amt.substr(0,num_len -3) + "," + nts_amt.substr(num_len -3,3);
				if (sil_len > 6) nts_amt = nts_amt.substr(0,num_len -6) + "," + nts_amt.substr(num_len -6,3) + "," + nts_amt.substr(num_len -2,3);
				document.frm.i_nts_amt.value = nts_amt;
				
				other_amt = String(other_amt);
				num_len = other_amt.length;
				sil_len = num_len;
				other_amt = String(other_amt);
				if (other_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) other_amt = other_amt.substr(0,num_len -3) + "," + other_amt.substr(num_len -3,3);
				if (sil_len > 6) other_amt = other_amt.substr(0,num_len -6) + "," + other_amt.substr(num_len -6,3) + "," + other_amt.substr(num_len -2,3);
				document.frm.i_other_amt.value = other_amt;
			}		
			
			 function setaddr() {
			 var srt = document.frm.i_family.value;
//			 alert(srt);
			 var arr = srt.split(','); 
			 var sub_string = arr[arr.length-1]; 
			 var sub_temp1 = sub_string.substring(0,6); 
			 var sub_temp2 = sub_string.substring(6,13); 
//             alert(sub_temp1);
//			 alert(sub_temp2);
			 document.frm.i_person_no.value = arr[arr.length-1];
			 document.frm.i_person_no1.value = sub_temp1;
			 document.frm.i_person_no2.value = sub_temp2;
			 document.frm.i_name.value = arr[arr.length-2];
			 document.frm.i_rel.value = arr[arr.length-3];
             }

			
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_insurance_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="15%" >
						<col width="25%" >
						<col width="20%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="i_emp_no" type="text" id="i_emp_no" size="10" value="<%=i_emp_no%>" readonly="true">
                      <input type="hidden" name="i_year" value="<%=i_year%>" ID="b_year">
                      <input type="hidden" name="i_seq" value="<%=i_seq%>" ID="b_seq"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="i_emp_name" type="text" id="i_emp_name" size="10" value="<%=i_emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>대상자</th>
                      <td class="left">
					   <select name="i_family" id="i_family" style="width:90px" onChange="setaddr();">
                          <option value="" <% if i_name = "" then %>selected<% end if %>>선택</option>
                  <% 
						for i = 1 to 10
						    if family_tab(i,2) = "" or isnull(family_tab(i,2)) then 
			                           exit for
		                       else
			  	  %>
                		  <option value='<%=family_tab(i,1)%>,<%=family_tab(i,2)%>,<%=family_tab(i,3)%>' <%If i_name = family_tab(i,2) then %>selected<% end if %>><%=family_tab(i,2)%></option>
                  <%
				            end if
						next
				  %>
            		  </select>
                      <th>관계/주민등록번호</th>
					  <td class="left">
                      <input name="i_name" type="hidden" value="<%=i_name%>" readonly="true" style="width:70px">
                      <input name="i_rel" type="text" value="<%=i_rel%>" readonly="true" style="width:60px">
                      <input name="i_person_no1" type="text" value="<%=i_person_no1%>" readonly="true" style="width:50px">
                      -
                      <input name="i_person_no2" type="text" value="<%=i_person_no2%>" readonly="true" style="width:60px">
                      <input name="i_person_no" type="hidden" value="<%=i_person_no%>" readonly="true" style="width:130px">
                      </td>
                    </tr>
                    <tr>
                      <th>국세청금액</th>
					  <td class="left">
                      <input name="i_nts_amt" type="text" id="i_nts_amt" style="width:90px;text-align:right" value="<%=formatnumber(i_nts_amt,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>그밖의금액</th>
					  <td class="left">
                      <input name="i_other_amt" type="text" id="i_other_amt" style="width:90px;text-align:right" value="<%=formatnumber(i_other_amt,0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <th>장애인전용<br>보장성보험</th>
                      <td colspan="3" class="left">
					  <input type="checkbox" name="i_disab_chk" value="Y" <% if i_disab_chk = "Y" then %>checked<% end if %> id="i_disab_chk">예
					  </td>
                    </tr>
                    <tr>
                      <td colspan="4" class="left">※ 보험료입력은 피보험자 기준으로 입력<br>
                   &nbsp;&nbsp;&nbsp;&nbsp;(예) 계약자가 본인이고 피보험자가 자녀인경우 보혐료는 자녀를 선택하고 입력해야 함<br>
                ※ 증빙자료를 국세청에서 발급받은 경우 국세청금액에 입력하고 보험사등에서 직접발급받은경우는 그밖의금액에 입력<br>
                ※ 보험료공제는 계약자와 피보험자가 모두 본인의 기본공제 대상만 공제받을 수 있음</td>
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

