<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim family_tab(10,8)

u_type = request("u_type")
m_year = request("m_year")
m_emp_no = request("m_emp_no")
m_person_no = request("m_person_no")
m_emp_name = request("m_emp_name")
m_seq = request("m_seq")

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
	family_tab(i,4) = ""
	family_tab(i,5) = ""
	family_tab(i,6) = ""
	family_tab(i,7) = ""
	family_tab(i,8) = ""
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from pay_yeartax_family where f_year = '"&m_year&"' and f_emp_no = '"&m_emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
'Set rs_fami = DbConn.Execute(SQL)
i = 0
do until rs_fami.eof
   if rs_fami("f_rel") = "����" or rs_fami("f_wife") = "Y" or rs_fami("f_age20") = "Y" or rs_fami("f_age60") = "Y" or rs_fami("f_old") = "Y" then
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
		  family_tab(i,6) = rs_fami("f_national")
		  family_tab(i,7) = rs_fami("f_pensioner")
		  family_tab(i,8) = rs_fami("f_witak")
	end if
	rs_fami.MoveNext()
loop
rs_fami.close()

title_line = " �Ƿ�� �����׸� �Է� "
if u_type = "U" then

	Sql="select * from pay_yeartax_medical where m_year = '"&m_year&"' and m_emp_no = '"&m_emp_no&"' and m_person_no = '"&m_person_no&"' and m_seq = '"&m_seq&"'"
	Set rs=DbConn.Execute(Sql)

	m_rel = rs("m_rel")
    m_name = rs("m_name")
	m_national = rs("m_national")
	m_pensioner = rs("m_pensioner")
	m_witak = rs("m_witak")
    m_disab = rs("m_disab")
    m_age65 = rs("m_age65")
	m_trade_no = rs("m_trade_no")
	m_trade_name = rs("m_trade_name")
	m_eye = rs("m_eye")
	m_data_gubun = rs("m_data_gubun")
	m_cnt = rs("m_cnt")
	m_amt = rs("m_amt")

	rs.close()

	title_line = " �Ƿ�� �����׸� ����  "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
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
				if(document.frm.m_data_gubun.value =="") {
					alert('�Ƿ�������ڵ带 �����ϼ���');
					frm.m_data_gubun.focus();
					return false;}
				if(document.frm.m_family.value =="") {
					alert('����ڸ� �����ϼ���');
					frm.m_family.focus();
					return false;}
				if(document.frm.m_amt.value =="") {
					alert('�ݾ��� �Է��ϼ���');
					frm.m_amt.focus();
					return false;}
			    if(document.frm.m_data_gubun.value != "����û") {
					if(document.frm.m_trade_no.value == "") {
							alert('����ó ����ڵ�Ϲ�ȣ�� �Է��ϼ���');
							frm.m_trade_no.focus();
							return false;}}
				if(document.frm.m_data_gubun.value != "����û") {
					if(document.frm.m_trade_name.value == "") {
							alert('����ó ��ȣ���� �Է��ϼ���');
							frm.m_trade_name.focus();
							return false;}}
				if(document.frm.m_data_gubun.value != "����û") {
					if(document.frm.m_cnt.value == "") {
							alert('�Ǽ��� �Է��ϼ���');
							frm.m_cnt.focus();
							return false;}}
				
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			} 
			
			function num_chk(txtObj){
				mm_cnt = parseInt(document.frm.m_cnt.value.replace(/,/g,""));	
				mm_amt = parseInt(document.frm.m_amt.value.replace(/,/g,""));	
		
				mm_cnt = String(mm_cnt);
				num_len = mm_cnt.length;
				sil_len = num_len;
				mm_cnt = String(mm_cnt);
				if (mm_cnt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) mm_cnt = mm_cnt.substr(0,num_len -3) + "," + mm_cnt.substr(num_len -3,3);
				if (sil_len > 6) mm_cnt = mm_cnt.substr(0,num_len -6) + "," + mm_cnt.substr(num_len -6,3) + "," + mm_cnt.substr(num_len -2,3);
				document.frm.m_cnt.value = mm_cnt;
				
				mm_amt = String(mm_amt);
				num_len = mm_amt.length;
				sil_len = num_len;
				mm_amt = String(mm_amt);
				if (mm_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) mm_amt = mm_amt.substr(0,num_len -3) + "," + mm_amt.substr(num_len -3,3);
				if (sil_len > 6) mm_amt = mm_amt.substr(0,num_len -6) + "," + mm_amt.substr(num_len -6,3) + "," + mm_amt.substr(num_len -2,3);
				document.frm.m_amt.value = mm_amt;
			}		
			
			 function setaddr() {
			 var srt = document.frm.m_family.value;
//			 alert(srt);
			 var arr = srt.split(','); 
			 var sub_string = arr[arr.length-6]; 
			 var sub_temp1 = sub_string.substring(0,6); 
			 var sub_temp2 = sub_string.substring(6,13); 
//             alert(sub_temp1);
//			 alert(sub_temp2);
			 document.frm.m_person_no.value = arr[arr.length-6];
			 document.frm.m_person_no1.value = sub_temp1;
			 document.frm.m_person_no2.value = sub_temp2;
			 document.frm.m_name.value = arr[arr.length-7];
			 document.frm.m_rel.value = arr[arr.length-8];
//			 alert(arr[arr.length-2]);
			 document.frm.m_disab.value = arr[arr.length-5];
			 document.frm.m_age65.value = arr[arr.length-4];
			 document.frm.m_nation.value = arr[arr.length-3];
			 document.frm.m_pensioner.value = arr[arr.length-2];
			 document.frm.m_witak.value = arr[arr.length-1];
             }

			
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_medical_save.asp" method="post" name="frm">
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
                      <th style="background:#FFFFE6">���</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="m_emp_no" type="text" id="m_emp_no" size="10" value="<%=m_emp_no%>" readonly="true">
                      <input type="hidden" name="m_year" value="<%=m_year%>" ID="m_year">
                      <input type="hidden" name="m_seq" value="<%=m_seq%>" ID="m_seq"></td>
                      <th style="background:#FFFFE6">����</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="m_emp_name" type="text" id="m_emp_name" size="10" value="<%=m_emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th>�Ƿ��<br>�����ڵ�</th>
					  <td colspan="7" class="left">
					  <select name="m_data_gubun" id="m_data_gubun" value="<%=m_data_gubun%>" style="width:120px">
				          <option value="" <% if m_data_gubun = "" then %>selected<% end if %>>����</option>
				          <option value='����û' <%If m_data_gubun = "����û" then %>selected<% end if %>>����û</option>
				          <option value='���ΰǰ��������' <%If m_data_gubun = "���ΰǰ��������" then %>selected<% end if %>>���ΰǰ��������</option>
				          <option value='�����/������' <%If m_data_gubun = "�����/������" then %>selected<% end if %>>�����/������</option>
                          <option value='�����޿�' <%If m_data_gubun = "�����޿�" then %>selected<% end if %>>�����޿�</option>
                          <option value='��Ÿ�Ƿ�񿵼���' <%If m_data_gubun = "��Ÿ�Ƿ�񿵼���" then %>selected<% end if %>>��Ÿ�Ƿ�񿵼���</option>
                      </select>
                      </td>
                    </tr>
                 	<tr>
                      <th>�����</th>
                      <td colspan="3" class="left">
					   <select name="m_family" id="m_family" style="width:90px" onChange="setaddr();">
                          <option value="" <% if m_name = "" then %>selected<% end if %>>����</option>
                  <% 
						for i = 1 to 10
						    if family_tab(i,2) = "" or isnull(family_tab(i,2)) then 
			                           exit for
		                       else
			  	  %>
                		  <option value='<%=family_tab(i,1)%>,<%=family_tab(i,2)%>,<%=family_tab(i,3)%>,<%=family_tab(i,4)%>,<%=family_tab(i,5)%>,<%=family_tab(i,6)%>,<%=family_tab(i,7)%>,<%=family_tab(i,8)%>' <%If m_name = family_tab(i,2) then %>selected<% end if %>><%=family_tab(i,2)%></option>
                  <%
				            end if
						next
				  %>
            		  </select>
                      <th>����/<br>�ֹε�Ϲ�ȣ</th>
					  <td colspan="3" class="left">
                      <input name="m_name" type="hidden" value="<%=m_name%>" readonly="true" style="width:70px">
                      <input name="m_rel" type="text" value="<%=m_rel%>" readonly="true" style="width:60px">
                      <input name="m_nation" type="hidden" value="<%=m_nation%>" readonly="true" style="width:60px">
                      <input name="m_pensioner" type="hidden" value="<%=m_pensioner%>" readonly="true" style="width:60px">
                      <input name="m_witak" type="hidden" value="<%=m_witak%>" readonly="true" style="width:60px">
                      <input name="m_person_no1" type="text" value="<%=m_person_no1%>" readonly="true" style="width:50px;text-align:center">
                      -
                      <input name="m_person_no2" type="text" value="<%=m_person_no2%>" readonly="true" style="width:60px;text-align:center">
                      <input name="m_person_no" type="hidden" value="<%=m_person_no%>" readonly="true" style="width:130px">
                      </td>
                      </td>
                    </tr>
                    </tr>
                    <tr>
                      <th>�����</th>
                      <td colspan="3" class="left">
					  <input name="m_disab" type="text" value="<%=m_disab%>" style="width:20px;text-align:center"" id="m_disab" readonly="true"></td>
					  </td>
                      <th>65���̻�</th>
                      <td colspan="3" class="left">
					  <input name="m_age65" type="text" value="<%=m_age65%>" style="width:20px;text-align:center"" id="m_age65" readonly="true"></td>
                      </td>
                    </tr>
                    <tr>
                      <th>����ڵ��<br>��ȣ</th>
                      <td class="left">
                      <input name="m_trade_no" type="text" value="<%=m_trade_no%>" style="width:90px" id="m_trade_no"></td>
                      <th>��ȣ��</th>
                      <td class="left">
                      <input name="m_trade_name" type="text" value="<%=m_trade_name%>" style="width:100px" id="m_trade_name"></td>
                      <th>�Ǽ�</th>
					  <td class="left">
                      <input name="m_cnt" type="text" id="m_cnt" style="width:90px;text-align:right" value="<%=formatnumber(m_cnt,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>�ݾ�</th>
					  <td class="left">
                      <input name="m_amt" type="text" id="m_amt" style="width:90px;text-align:right" value="<%=formatnumber(m_amt,0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <th>�Ȱ��<br>���Կ���</th>
                      <td colspan="7" class="left">
					  <input type="checkbox" name="m_eye" value="Y" <% if m_eye = "Y" then %>checked<% end if %> id="m_eye">��
					  </td>
                    </tr>
                    <tr>
                      <td colspan="8" class="left">�� �Ƿ�� �ݾ� �Է��� �Ƿ�񳻿��� �ִ� ����ڸ� �����ϰ� �Է�<br>
                �� �Ƿ�������ڵ�-����û�� �����ϴ� �Ƿ��� ����ڹ�ȣ,��ȣ,�Ǽ� �Է����� �ʰ� �ݾ׸� �ջ��Ͽ� �Է�<br>
                �� �׿� �����ΰ��� �κ�,���ó���� �Է��� �ؾ� ��</td>
                    </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

