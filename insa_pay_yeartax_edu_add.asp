<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim family_tab(10,5)

u_type = request("u_type")
e_year = request("e_year")
e_emp_no = request("e_emp_no")
e_person_no = request("e_person_no")
e_emp_name = request("e_emp_name")
e_seq = request("e_seq")

e_person_no1 = mid(cstr(e_person_no),1,6)
e_person_no2 = mid(cstr(e_person_no),7,7)

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
	end if
	rs_fami.MoveNext()
loop
rs_fami.close()

title_line = " ������ �����׸� �Է� "
if u_type = "U" then

	Sql="select * from pay_yeartax_edu where e_year = '"&e_year&"' and e_emp_no = '"&e_emp_no&"' and e_person_no = '"&e_person_no&"' and e_seq = '"&e_seq&"'"
	Set rs=DbConn.Execute(Sql)

	e_rel = rs("e_rel")
    e_name = rs("e_name")
    e_disab = rs("e_disab")
	e_uniform = rs("e_uniform")
	e_edu_level = rs("e_edu_level")
	e_nts_amt = rs("e_nts_amt")
	e_other_amt = rs("e_other_amt")

	rs.close()

	title_line = " ������ �����׸� ����  "
	
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
				if(document.frm.e_edu_level.value =="") {
					alert('���������� �����ϼ���');
					frm.e_edu_level.focus();
					return false;}
				if(document.frm.e_family.value =="") {
					alert('����ڸ� �����ϼ���');
					frm.e_family.focus();
					return false;}
				
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			} 
			
			function num_chk(txtObj){
				nts_amt = parseInt(document.frm.e_nts_amt.value.replace(/,/g,""));	
				other_amt = parseInt(document.frm.e_other_amt.value.replace(/,/g,""));	
		
				nts_amt = String(nts_amt);
				num_len = nts_amt.length;
				sil_len = num_len;
				nts_amt = String(nts_amt);
				if (nts_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) nts_amt = nts_amt.substr(0,num_len -3) + "," + nts_amt.substr(num_len -3,3);
				if (sil_len > 6) nts_amt = nts_amt.substr(0,num_len -6) + "," + nts_amt.substr(num_len -6,3) + "," + nts_amt.substr(num_len -2,3);
				document.frm.e_nts_amt.value = nts_amt;
				
				other_amt = String(other_amt);
				num_len = other_amt.length;
				sil_len = num_len;
				other_amt = String(other_amt);
				if (other_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) other_amt = other_amt.substr(0,num_len -3) + "," + other_amt.substr(num_len -3,3);
				if (sil_len > 6) other_amt = other_amt.substr(0,num_len -6) + "," + other_amt.substr(num_len -6,3) + "," + other_amt.substr(num_len -2,3);
				document.frm.e_other_amt.value = other_amt;
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
				<form action="insa_pay_yeartax_edu_save.asp" method="post" name="frm">
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
                      <th style="background:#FFFFE6">���</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="e_emp_no" type="text" id="e_emp_no" size="10" value="<%=e_emp_no%>" readonly="true">
                      <input type="hidden" name="e_year" value="<%=e_year%>" ID="m_year">
                      <input type="hidden" name="e_seq" value="<%=e_seq%>" ID="m_seq"></td>
                      <th style="background:#FFFFE6">����</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="e_emp_name" type="text" id="e_emp_name" size="10" value="<%=e_emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th>��������</th>
					  <td colspan="7" class="left">
					  <select name="e_edu_level" id="e_edu_level" value="<%=e_edu_level%>" style="width:150px">
				          <option value="" <% if e_edu_level = "" then %>selected<% end if %>>����</option>
				          <option value='�ҵ��ں���' <%If e_edu_level = "�ҵ��ں���" then %>selected<% end if %>>�ҵ��ں���</option>
				          <option value='�������Ƶ�' <%If e_edu_level = "�������Ƶ�" then %>selected<% end if %>>�������Ƶ�</option>
				          <option value='��/��/����б�' <%If e_edu_level = "��/��/����б�" then %>selected<% end if %>>��/��/����б�</option>
                          <option value='���л�(���п�������)' <%If e_edu_level = "���л�(���п�������)" then %>selected<% end if %>>���л�(���п�������)</option>
                          <option value='�����' <%If e_edu_level = "�����" then %>selected<% end if %>>�����</option>
                      </select>
                      </td>
                    </tr>
                 	<tr>
                      <th>�����</th>
                      <td colspan="3" class="left">
					   <select name="e_family" id="e_family" style="width:90px" onChange="setaddr();">
                          <option value="" <% if e_name = "" then %>selected<% end if %>>����</option>
                  <% 
						for i = 1 to 10
						    if family_tab(i,2) = "" or isnull(family_tab(i,2)) then 
			                           exit for
		                       else
			  	  %>
                		  <option value='<%=family_tab(i,1)%>,<%=family_tab(i,2)%>,<%=family_tab(i,3)%>,<%=family_tab(i,4)%>,<%=family_tab(i,5)%>' <%If e_name = family_tab(i,2) then %>selected<% end if %>><%=family_tab(i,2)%></option>
                  <%
				            end if
						next
				  %>
            		  </select>
                      <th>����/<br>�ֹε�Ϲ�ȣ</th>
					  <td colspan="3" class="left">
                      <input name="e_name" type="hidden" value="<%=e_name%>" readonly="true" style="width:70px">
                      <input name="e_rel" type="text" value="<%=e_rel%>" readonly="true" style="width:60px">
                      <input name="e_person_no1" type="text" value="<%=e_person_no1%>" readonly="true" style="width:50px;text-align:center">
                      -
                      <input name="e_person_no2" type="text" value="<%=e_person_no2%>" readonly="true" style="width:60px;text-align:center">
                      <input name="e_person_no" type="hidden" value="<%=e_person_no%>" readonly="true" style="width:130px">
                      </td>
                      </td>
                    </tr>
                    </tr>
                    <tr>
                      <th>�����</th>
                      <td colspan="3" class="left">
					  <input name="e_disab" type="text" value="<%=e_disab%>" style="width:20px;text-align:center"" id="e_disab" readonly="true">
					  </td>
                      <th>��������<br>�񿩺�</th>
                      <td colspan="3" class="left">
					  <input type="checkbox" name="e_uniform" value="Y" <% if e_uniform = "Y" then %>checked<% end if %> id="e_uniform">��
					  </td>
                    </tr>
                    <tr>
                      <th>����û�ݾ�</th>
					  <td colspan="3" class="left">
                      <input name="e_nts_amt" type="text" id="e_nts_amt" style="width:90px;text-align:right" value="<%=formatnumber(e_nts_amt,0)%>" onKeyUp="num_chk(this);"></td>
                      <th>�׹��Ǳݾ�</th>
					  <td colspan="3" class="left">
                      <input name="e_other_amt" type="text" id="e_other_amt" style="width:90px;text-align:right" value="<%=formatnumber(e_other_amt,0)%>" onKeyUp="num_chk(this);"></td>
                    </tr>
                    <tr>
                      <td colspan="8" class="left">�� �������� ������ �Է�<br>
                �� �߰���б� �ڳ� �������� �������Աݾ��� �ִ� ��� �������Ժ񿩺ο� üũ�ϰ� �������Ժ� �޾��� �Է�, ������� ����� ���� �Է��ؾ� ��.<br>
                �� ������ ������ ����û���� �߱޹��� ���� ����û�ݾ׿� �Է��ϰ�, �б�.��ġ������ ���ѿ��� ���� �߱޹��� ��� �׹��Ǳݾ׿� �Է�<br>
                �� �ʵ��л��̻��� �缳�п������ �������������� �ƴ�, �н���� ��������� ����� �ƴ�.</td>
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

