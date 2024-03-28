<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

u_type = request("u_type")
insu_id = request("insu_id")
view_condi = request("view_condi")
insu_yyyy = request("insu_yyyy")
insu_class = request("insu_class")

' �ֱ�3���⵵ ���̺�� ����
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "��"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "��"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "��"

' �б� ���̺� ����
curr_mm = mid(now(),6,2)
if curr_mm > 0 and curr_mm < 4 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "1"
end if
if curr_mm > 3 and curr_mm < 7 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "2"
end if
if curr_mm > 6 and curr_mm < 10 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "3"
end if
if curr_mm > 9 and curr_mm < 13 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "4"
end if

quarter_tab(8,2) = cstr(mid(quarter_tab(8,1),1,4)) + "�� " + cstr(mid(quarter_tab(8,1),5,1)) + "/4�б�"

for i = 7 to 1 step -1
	cal_quarter = cint(quarter_tab(i+1,1)) - 1
	if cstr(mid(cal_quarter,5,1)) = "0" then
		quarter_tab(i,1) = cstr(cint(mid(cal_quarter,1,4))-1) + "4"
	  else
		quarter_tab(i,1) = cal_quarter
	end if	 
	quarter_tab(i,2) = cstr(mid(quarter_tab(i,1),1,4)) + "�� " + cstr(mid(quarter_tab(i,1),5,1)) + "/4�б�"
next

' ��� ���̺����
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

insu_id_name = request("view_condi")
from_amt = 0
to_amt = 0
st_amt = 0
hap_rate = 0
emp_rate = 0
com_rate = 0
insu_comment = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "4�뺸����� ���"

if u_type = "U" then

	sql = "select * from pay_insurance where insu_yyyy = '" + insu_yyyy + "' and insu_id = '" + insu_id + "' and insu_class = '" + insu_class + "'"
	set rs = dbconn.execute(sql)

    insu_yyyy = rs("insu_yyyy")
    insu_id = rs("insu_id")
	insu_class = rs("insu_class")
    insu_id_name = rs("insu_id_name")
	from_amt =rs("from_amt")
    to_amt = rs("to_amt")
    st_amt = rs("st_amt")
    hap_rate = rs("hap_rate")
    emp_rate = rs("emp_rate")
    com_rate = rs("com_rate")
    insu_comment = rs("insu_comment")
	rs.close()

	title_line = "4�뺸����� ����"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ�޿� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
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
				if(document.frm.insu_yyyy.value =="" && document.frm.insu_yyyy.value =="") {
					alert('����⵵ �Է��ϼ���');
					frm.insu_yyyy.focus();
					return false;}
				if(document.frm.insu_class.value =="") {
					alert('����� �Է��ϼ���');
					frm.insu_class.focus();
					return false;}
				if(document.frm.emp_rate.value =="") {
					alert('�ٷ��� ������ �Է��ϼ���');
					frm.emp_rate.focus();
					return false;}			
				if(document.frm.com_rate.value =="") {
					alert('����� ������ �Է��ϼ���');
					frm.com_rate.focus();
					return false;}			
				if(document.frm.to_amt.value =="") {
					alert('ǥ�غ��������� �Է��ϼ���');
					frm.to_amt.focus();
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
				f_amt = parseInt(document.frm.from_amt.value.replace(/,/g,""));		
				f_amt = String(f_amt);
				num_len = f_amt.length;
				sil_len = num_len;
				f_amt = String(f_amt);
				if (f_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) f_amt = f_amt.substr(0,num_len -3) + "," + f_amt.substr(num_len -3,3);
				if (sil_len > 6) f_amt = f_amt.substr(0,num_len -6) + "," + f_amt.substr(num_len -6,3) + "," + f_amt.substr(num_len -2,3);
				document.frm.from_amt.value = f_amt; 
				
				t_amt = parseInt(document.frm.to_amt.value.replace(/,/g,""));		
				t_amt = String(t_amt);
				num_len = t_amt.length;
				sil_len = num_len;
				t_amt = String(t_amt);
				if (t_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) t_amt = t_amt.substr(0,num_len -3) + "," + t_amt.substr(num_len -3,3);
				if (sil_len > 6) t_amt = t_amt.substr(0,num_len -6) + "," + t_amt.substr(num_len -6,3) + "," + t_amt.substr(num_len -2,3);
				document.frm.to_amt.value = t_amt; 
				
				s_amt = parseInt(document.frm.st_amt.value.replace(/,/g,""));		
				s_amt = String(s_amt);
				num_len = s_amt.length;
				sil_len = num_len;
				s_amt = String(s_amt);
				if (s_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) s_amt = s_amt.substr(0,num_len -3) + "," + s_amt.substr(num_len -3,3);
				if (sil_len > 6) s_amt = s_amt.substr(0,num_len -6) + "," + s_amt.substr(num_len -6,3) + "," + s_amt.substr(num_len -2,3);
				document.frm.st_amt.value = s_amt; 		
				
                e_rate = parseFloat((document.frm.emp_rate.value),3);	
				c_rate = parseFloat((document.frm.com_rate.value),3);	
				h_rate = e_rate + c_rate;
				document.frm.hap_rate.value = h_rate; 	

			}						
			
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_insurance_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="10%" >
							<col width="23%" >
							<col width="10%" >
							<col width="23%" >
							<col width="10%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">���豸��</th>
								<td class="left">
                                <input name="insu_id_name" type="text" value="<%=insu_id_name%>" style="width:120px" readonly="true"></td>
								<th>����⵵</th>
								<td class="left">
                                <select name="insu_yyyy" id="insu_yyyy" type="text" value="<%=insu_yyyy%>" style="width:90px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If insu_yyyy = year_tab(i,1) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
                                </td>
                                <th>���<br>01����~</th>
								<td class="left">
                                <input name="insu_class" type="text" value="<%=insu_class%>" style="width:30px" onKeyUp="checklength(this,2)"></td>
							</tr>
                           	<tr>
								<th class="first">��������<br>�̻�-�̸�</th>
								<td colspan="3" class="left">
                                <input name="from_amt" type="text" value="<%=formatnumber(from_amt,0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                -
                                <input name="to_amt" type="text" value="<%=formatnumber(to_amt,0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th>ǥ��<br>��������</th>
								<td class="left">
                                <input name="st_amt" type="text" value="<%=formatnumber(st_amt,0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
							</tr>             
							<tr>
								<th colspan="6" class="first" style="background:#F5FFFA">����(�����)</th>
							</tr>
							<tr>
								<th class="first">�ٷ���</th>
								<td class="left">
                                <input name="emp_rate" type="text" value="<%=formatnumber(emp_rate,3)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th class="first">�����</th>
								<td class="left">
                                <input name="com_rate" type="text" value="<%=formatnumber(com_rate,3)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th class="first">�հ�</th>
								<td class="left">
                                <input name="hap_rate" type="text" value="<%=formatnumber(hap_rate,3)%>" style="width:90px;text-align:right" readonly="true">
                                </td>
							</tr>
                        	<tr>
								<th class="first">���</th>
								<td colspan="5" class="left">
                                <input name="insu_comment" type="text" value="<%=insu_comment%>" style="width:570px" onKeyUp="checklength(this,50)">
                                </td>
							</tr>                            
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="insu_id" value="<%=insu_id%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

