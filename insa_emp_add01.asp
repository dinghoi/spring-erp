<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)
curr_hh = int(cstr(datepart("h",now)))
curr_mm = int(cstr(datepart("n",now)))

insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
pay_grade = request.cookies("nkpmg_user")("coo_pay_grade")

' �Է¹޾� ����Ÿ�� ��Ƶ� �ʵ��̸��� ���ǿ� �⺻���� null�� ����Ѱ�

u_type = request("u_type")
emp_no = request("emp_no")
view_condi=Request("view_condi")

code_last = ""
emp_reg_user = ""
emp_mod_user = ""

emp_name = ""
emp_ename = ""
emp_type = ""
emp_sex = ""
emp_person1 = ""
emp_person2 = ""
emp_image = ""
emp_first_date = ""
emp_in_date = ""
emp_gunsok_date = ""
emp_yuncha_date = ""
emp_end_gisan = ""
emp_end_date = ""
emp_company = ""
emp_bonbu = ""
emp_saupbu = ""
emp_team = ""
emp_org_code = ""
emp_org_name = ""
emp_org_baldate = ""
emp_stay_code = ""
emp_stay_name = ""
emp_reside_place = ""
emp_reside_company = ""
emp_grade = ""
emp_grade_date = ""
emp_job = ""
emp_position = ""
emp_jikgun = ""
emp_jikmu = ""
emp_birthday = ""
emp_birthday_id = ""
emp_family_zip = ""
emp_family_sido = ""
emp_family_gugun = ""
emp_family_dong = ""
emp_family_addr = ""
emp_zipcode = ""
emp_sido = ""
emp_gugun = ""
emp_dong = ""
emp_addr = ""
emp_tel_ddd = ""
emp_tel_no1 = ""
emp_tel_no2 = ""
emp_hp_ddd = ""
emp_hp_no1 = ""
emp_hp_no2 = ""
emp_email = ""
emp_military_id = ""
emp_military_date1 = ""
emp_military_date2 = ""
emp_military_grade = ""
emp_military_comm = ""
emp_hobby = ""
emp_faith = ""
emp_last_edu = ""
emp_marry_date = ""
emp_disabled = "�ش���׾���"
emp_disab_grade = ""
emp_sawo_id = "N"
emp_sawo_date = ""
emp_emergency_tel = ""
emp_extension_no = ""
emp_nation_code = ""
cost_center = ""
cost_group = ""
att_file = ""
emp_pay_id = "0"

emp_reg_date = now()
emp_reg_user = in_name
' response.write(emp_reg_date)

first_date = curr_date
request_hh = curr_hh
request_mm = curr_mm

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_owner = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect


title_line = "[ �λ�⺻���� ��� ]"
mg_group = "1"
if u_type = "U" then

	Sql="select * from emp_master where emp_no = '"&emp_no&"'"
	Set rs=DbConn.Execute(Sql)

	emp_name = rs("emp_name")
	emp_ename = rs("emp_ename")
	emp_type = rs("emp_type")
	emp_sex = rs("emp_sex")
	emp_person1 = rs("emp_person1")
	emp_person2 = rs("emp_person2")
	emp_image = rs("emp_image")
	att_file = rs("emp_image")
	emp_first_date = rs("emp_first_date")
	emp_in_date = rs("emp_in_date")
	emp_gunsok_date = rs("emp_gunsok_date")
	emp_yuncha_date = rs("emp_yuncha_date")
	emp_end_gisan = rs("emp_end_gisan")
	emp_end_date = rs("emp_end_date")
	emp_company = rs("emp_company")
	emp_bonbu = rs("emp_bonbu")
	emp_saupbu = rs("emp_saupbu")
	emp_team = rs("emp_team")
	emp_org_code = rs("emp_org_code")
	emp_org_name = rs("emp_org_name")
	emp_org_baldate = rs("emp_org_baldate")
	emp_stay_code = rs("emp_stay_code")
	emp_stay_name = rs("emp_stay_name")
	emp_reside_place = rs("emp_reside_place")
	emp_reside_company = rs("emp_reside_company")
	emp_grade = rs("emp_grade")
	emp_grade_date = rs("emp_grade_date")
	emp_job = rs("emp_job")
	emp_position = rs("emp_position")
	emp_jikgun = rs("emp_jikgun")
	emp_jikmu = rs("emp_jikmu")
	emp_birthday = rs("emp_birthday")
	emp_birthday_id = rs("emp_birthday_id")
	emp_family_zip = rs("emp_family_zip")
	emp_family_sido = rs("emp_family_sido")
	emp_family_gugun = rs("emp_family_gugun")
	emp_family_dong = rs("emp_family_dong")
	emp_family_addr = rs("emp_family_addr")
	emp_zipcode = rs("emp_zipcode")
	emp_sido = rs("emp_sido")
	emp_gugun = rs("emp_gugun")
	emp_dong = rs("emp_dong")
	emp_addr = rs("emp_addr")
	emp_tel_ddd = rs("emp_tel_ddd")
	emp_tel_no1 = rs("emp_tel_no1")
	emp_tel_no2 = rs("emp_tel_no2")
	emp_hp_ddd = rs("emp_hp_ddd")
	emp_hp_no1 = rs("emp_hp_no1")
	emp_hp_no2 = rs("emp_hp_no2")
	emp_email = rs("emp_email")
	emp_military_id = rs("emp_military_id")
	emp_military_date1 = rs("emp_military_date1")
	emp_military_date2 = rs("emp_military_date2")
	emp_military_grade = rs("emp_military_grade")
	emp_military_comm = rs("emp_military_comm")
	emp_hobby = rs("emp_hobby")
	emp_faith = rs("emp_faith")
	emp_last_edu = rs("emp_last_edu")
	emp_marry_date = rs("emp_marry_date")
	emp_disabled = rs("emp_disabled")
	emp_disab_grade = rs("emp_disab_grade")
	emp_sawo_id = rs("emp_sawo_id")

	if rs("emp_sawo_id") = "" or isNull(emp_sawo_id) then
	   emp_sawo_id = "N"
	end if

	emp_sawo_date = rs("emp_sawo_date")
	emp_emergency_tel = rs("emp_emergency_tel")
	emp_nation_code = rs("emp_nation_code")
	emp_pay_id = rs("emp_pay_id")
	emp_extension_no = rs("emp_extension_no")
	cost_center = rs("cost_center")
	cost_group = rs("cost_group")
	'   end_date = mid(cstr(now()),1,10)
	emp_reg_date = rs("emp_reg_date")
	emp_reg_user = rs("emp_reg_user")
	emp_mod_date = rs("emp_mod_date")
	emp_mod_user = rs("emp_mod_user")
	photo_image = "/emp_photo/" + rs("emp_image")
	att_file = rs("emp_image")

	if rs("emp_military_date1") = "1900-01-01" then
  	emp_military_date1 = ""
    emp_military_date2 = ""
  end if
  if rs("emp_birthday") = "1900-01-01" then
    emp_birthday = ""
    end if
	if rs("emp_marry_date") = "1900-01-01" then
    emp_marry_date = ""
    end if
	if rs("emp_grade_date") = "1900-01-01" then
    emp_grade_date = ""
  end if
	if rs("emp_end_date") = "1900-01-01" then
  	emp_end_date = ""
    end if
	if rs("emp_org_baldate") = "1900-01-01" then
  	emp_org_baldate = ""
  end if
	if rs("emp_sawo_date") = "1900-01-01" then
  	emp_sawo_date = ""
  end if

	rs.close()

	sql="select * from memb where user_id='"&emp_no&"'"
	set rs_memb=dbconn.execute(sql)
	if not rs_memb.eof then
	       mg_group = rs_memb("mg_group")
	   else
	       mg_group = "1"
    end if
	rs_memb.close()
	'Sql="select * from emp_org_mst where org_code = '"&owner_org&"'"
	'Set rs_owner=DbConn.Execute(Sql)

    'owner_orgname = rs_owner("org_name")
	'rs_owner.close()

	title_line = "[ �λ�⺻���� ���� ]"
end if

'response.write(org_level)

    sql="select max(emp_no) as max_seq from emp_master where emp_no < '900000'"
	set rs_max=dbconn.execute(sql)

	if	isnull(rs_max("max_seq"))  then
		code_last = "000001"
	  else
		max_seq = "000000" + cstr((int(rs_max("max_seq")) + 1))
		code_last = right(max_seq,6)
	end if
    rs_max.close()

	if u_type = "U" then
	   code_last = emp_no
	end if

emp_no = code_last
'response.write(emp_no)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ�޿� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=emp_first_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
											$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker1" ).datepicker("setDate", "<%=emp_in_date%>" );
			});
			$(function() {    $( "#datepicker2" ).datepicker();
											$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker2" ).datepicker("setDate", "<%=emp_end_gisan%>" );
			});
			$(function() {    $( "#datepicker3" ).datepicker();
											$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker3" ).datepicker("setDate", "<%=emp_gunsok_date%>" );
			});
			$(function() {    $( "#datepicker4" ).datepicker();
											$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker4" ).datepicker("setDate", "<%=emp_yuncha_date%>" );
			});
			$(function() {    $( "#datepicker5" ).datepicker();
											$( "#datepicker5" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker5" ).datepicker("setDate", "<%=emp_birthday%>" );
			});
			$(function() {    $( "#datepicker6" ).datepicker();
											$( "#datepicker6" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker6" ).datepicker("setDate", "<%=emp_sawo_date%>" );
			});
			$(function() {    $( "#datepicker7" ).datepicker();
											$( "#datepicker7" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker7" ).datepicker("setDate", "<%=emp_marry_date%>" );
			});
			$(function() {    $( "#datepicker8" ).datepicker();
											$( "#datepicker8" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker8" ).datepicker("setDate", "<%=emp_military_date1%>" );
			});
			$(function() {    $( "#datepicker9" ).datepicker();
											$( "#datepicker9" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker9" ).datepicker("setDate", "<%=emp_military_date2%>" );
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
				if(document.frm.emp_name.value =="") {
					alert('������ �Է��ϼ���');
					frm.emp_name.focus();
					return false;}
				if(document.frm.emp_ename.value =="") {
					alert('���������� �Է��ϼ���');
					frm.emp_ename.focus();
					return false;}
				if(document.frm.emp_birthday.value =="") {
					alert('��������� �Է��ϼ���');
					frm.emp_birthday.focus();
					return false;}
				if(document.frm.emp_org_code.value =="") {
					alert('�Ҽ��� �����ϼ���');
					frm.emp_org_code.focus();
					return false;}
				if(document.frm.emp_type.value =="") {
					alert('���������� �����ϼ���');
					frm.emp_type.focus();
					return false;}
				if(document.frm.emp_grade.value =="") {
					alert('������ �����ϼ���');
					frm.emp_grade.focus();
					return false;}
				if(document.frm.emp_job.value =="") {
					alert('������ �����ϼ���');
					frm.emp_job.focus();
					return false;}
				if(document.frm.emp_position.value =="") {
					alert('��å�� �����ϼ���');
					frm.emp_position.focus();
					return false;}
				if(document.frm.emp_jikmu.value =="") {
					alert('������ �����ϼ���');
					frm.emp_jikmu.focus();
					return false;}
				if(document.frm.emp_first_date.value =="") {
					alert('�����Ի����� �Է��ϼ���');
					frm.emp_first_date.focus();
					return false;}
				if(document.frm.emp_in_date.value =="") {
					alert('�Ի����� �Է��ϼ���');
					frm.emp_in_date.focus();
					return false;}
				if(document.frm.emp_end_gisan.value =="") {
					alert('����������� �Է��ϼ���');
					frm.emp_end_gisan.focus();
					return false;}
				if(document.frm.emp_gunsok_date.value =="") {
					alert('�ټӱ������ �Է��ϼ���');
					frm.emp_gunsok_date.focus();
					return false;}
				if(document.frm.emp_yuncha_date.value =="") {
					alert('����������� �Է��ϼ���');
					frm.emp_yuncha_date.focus();
					return false;}
				if(document.frm.emp_person1.value =="") {
					alert('�ֹε�Ϲ�ȣ�� �Է��ϼ���');
					frm.emp_person1.focus();
					return false;}
				if(document.frm.emp_person2.value =="") {
					alert('�ֹε�Ϲ�ȣ�� �Է��ϼ���');
					frm.emp_person2.focus();
					return false;}
				if(document.frm.emp_tel_ddd.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					return false;}
				if(document.frm.emp_tel_no1.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					return false;}
				if(document.frm.emp_tel_no2.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					return false;}
				if(document.frm.emp_hp_ddd.value =="") {
					alert('�ڵ�����ȣ�� �Է��ϼ���');
					return false;}
				if(document.frm.emp_hp_no1.value =="") {
					alert('�ڵ�����ȣ�� �Է��ϼ���');
					return false;}
				if(document.frm.emp_hp_no2.value =="") {
					alert('�ڵ�����ȣ�� �Է��ϼ���');
					return false;}
				if(document.frm.mg_group.value =="") {
					alert('�����׷쿩�θ� üũ �ϼ���');
					frm.mg_group.focus();
					return false;}
				if(document.frm.cost_center.value =="") {
					alert('�����׷쿩�θ� üũ �ϼ���');
					frm.cost_center.focus();
					return false;}

				if(document.frm.emp_first_date.value > document.frm.emp_in_date.value) {
						alert('�����Ի����� �Ի��Ϻ��� �ʽ��ϴ�');
						frm.emp_first_date.focus();
						return false;}
				if(document.frm.emp_in_date.value > document.frm.emp_end_gisan.value) {
						alert('����������� �Ի��Ϻ��� �����ϴ�');
						frm.emp_end_gisan.focus();
						return false;}
				if(document.frm.emp_in_date.value > document.frm.emp_yuncha_date.value) {
						alert('����������� �Ի��Ϻ��� �����ϴ�');
						frm.emp_yuncha_date.focus();
						return false;}
				if(document.frm.emp_military_id.value !=="")
					if(document.frm.emp_military_date1.value =="") {
						alert('���� �̷� ���ڸ� �Է��ϼ���');
						frm.emp_military_date1.focus();
						return false;}
				if(document.frm.cost_center.value =="����������")
				   if(document.frm.emp_reside_company.value =="") {
					alert('����óȸ�縦 �����ϼ���');
					frm.emp_reside_company.focus();
					return false;}

				a=confirm('����Ͻðڽ��ϱ�?');
				if (a==true) {
					return true;
				}
				return false;
			}
			function file_browse()	{
           		document.frm.att_file.click();
           		document.frm.text1.value=document.frm.att_file.value;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
    <%
    '<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">
	%>
		<div id="wrap">

			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_add01_save.asp" method="post" name="frm" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
						</colgroup>
						<tbody>
							<tr>
								<td colspan="2" rowspan="4" class="left">
									<img src="<%=photo_image%>" width=110 height=120 alt="">
                </td>
								<th>���&nbsp;&nbsp;��ȣ</th>
                <td class="left"><%=emp_no%><input name="emp_no" type="hidden" value="<%=emp_no%>"></td>
                <th>����(�ѱ�)</th>
                <td class="left"><input name="emp_name" type="text" id="emp_name" size="13" value="<%=emp_name%>"></td>
								<th>����(����)</th>
								<td colspan="2" class="left">
									<input name="emp_ename" type="text" id="emp_ename" style="width:160px" maxlength="20" value="<%=emp_ename%>">
								</td>
								<th>�������</th>
								<td colspan="2" class="left">
									<input name="emp_birthday" type="text" size="10" id="datepicker5" style="width:70px;" value="<%=emp_birthday%>" readonly="true">
									&nbsp;��&nbsp;
									<input type="radio" name="emp_birthday_id" value="��" <% if emp_birthday_id = "��" then %>checked<% end if %>>��
              		<input name="emp_birthday_id" type="radio" value="��" <% if emp_birthday_id = "��" then %>checked<% end if %>>��
                </td>
              </tr>
              <tr>
              	<th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
              	<td colspan="3" class="left">
									<input name="emp_org_code" type="text" id="emp_org_code" style="width:40px" readonly="true" value="<%=emp_org_code%>">
                	&nbsp;��&nbsp;
                	<input name="emp_org_name" type="text" id="emp_org_name" style="width:120px" readonly="true" value="<%=emp_org_name%>">
                	<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=<%="org"%>&mg_org=<%=mg_org%>&view_condi=<%=view_condi%>','orgselect','scrollbars=yes,width=800,height=400')">����</a>
              	</td>
              	<th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
              	<td colspan="5" class="left">
              		<input name="emp_company" type="text" id="emp_company" style="width:100px" readonly="true" value="<%=emp_company%>">
              		<input name="emp_bonbu" type="text" id="emp_bonbu" style="width:120px" readonly="true" value="<%=emp_bonbu%>">
              		<input name="emp_saupbu" type="text" id="emp_saupbu" style="width:120px" readonly="true" value="<%=emp_saupbu%>">
              		<input name="emp_team" type="text" id="emp_team" style="width:120px" readonly="true" value="<%=emp_team%>">
                	<input name="emp_reside_place" type="hidden" id="emp_reside_place" style="width:120px" readonly="true" value="<%=emp_reside_place%>">
                	<input name="emp_org_level" type="hidden" id="emp_org_level" style="width:120px" readonly="true" value="<%=emp_org_level%>">
              	</td>
            	</tr>
            	<tr>
            		<th>��������</th>
            		<td class="left">
            			<select name="emp_type" id="emp_type" value="<%=emp_type%>" style="width:90px">
			            	<option value="" <% if emp_type = "" then %>selected<% end if %>>����</option>
				            <option value='����' <%If emp_type = "����" then %>selected<% end if %>>����</option>
                    <option value='����' <%If emp_type = "����" then %>selected<% end if %>>����</option>
				            <option value='�����' <%If emp_type = "�����" then %>selected<% end if %>>�����</option>
                  </select>
                </td>
                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left">
                	<%
                		Sql="select * from emp_etc_code where emp_etc_type = '02' order by emp_etc_code asc"
                		Rs_etc.Open Sql, Dbconn, 1
							  	%>
									<select name="emp_grade" id="emp_grade" style="width:90px">
                  	<option value="" <% if emp_grade = "" then %>selected<% end if %>>����</option>
                		<%
                			do until rs_etc.eof
                		%>
                		<option value='<%=rs_etc("emp_etc_name")%>' <%If emp_grade = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		<%
                				rs_etc.movenext()
                			loop
                			rs_etc.Close()
							  		%>
            			</select>
                </td>
                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left">
									<%
										Sql="select * from emp_etc_code where emp_etc_type = '03' order by emp_etc_code asc"
										Rs_etc.Open Sql, Dbconn, 1
							  	%>
									<select name="emp_job" id="emp_job" style="width:90px">
                  	<option value="" <% if emp_job = "" then %>selected<% end if %>>����</option>
                		<%
                			do until rs_etc.eof
			  				  	%>
                		<option value='<%=rs_etc("emp_etc_name")%>' <%If emp_job = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		<%
                				rs_etc.movenext()
                			loop
                			rs_etc.Close()
							  		%>
            			</select>
                </td>
                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;å</th>
                <td class="left">
                	<%
										Sql="select * from emp_etc_code where emp_etc_type = '04' order by emp_etc_code asc"
										Rs_etc.Open Sql, Dbconn, 1
							  	%>
									<select name="emp_position" id="emp_position" style="width:90px">
                  	<option value="" <% if emp_position = "" then %>selected<% end if %>>����</option>
                		<%
                			do until rs_etc.eof
			  				  	%>
                		<option value='<%=rs_etc("emp_etc_name")%>' <%If emp_position = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		<%
												rs_etc.movenext()
											loop
											rs_etc.Close()
							  		%>
            			</select>
                </td>
                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left">
                	<%
                		Sql="select * from emp_etc_code where emp_etc_type = '05' order by emp_etc_code asc"
										Rs_etc.Open Sql, Dbconn, 1
							  	%>
									<select name="emp_jikmu" id="emp_jikmu" style="width:90px">
                  	<option value="" <% if emp_jikmu = "" then %>selected<% end if %>>����</option>
                		<%
											do until rs_etc.eof
			  				  	%>
                		<option value='<%=rs_etc("emp_etc_name")%>' <%If emp_jikmu = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		<%
												rs_etc.movenext()
											loop
											rs_etc.Close()
							  		%>
            			</select>
                </td>
              </tr>
              <tr>
              	<th>�����Ի���</th>
                <td class="left">
                	<input name="emp_first_date" type="text" size="10" id="datepicker" style="width:70px;" value="<%=emp_first_date%>" readonly="true">&nbsp;
                </td>
                <th>��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��</th>
                <td class="left">
									<input name="emp_in_date" type="text" size="10" id="datepicker1" style="width:70px;" value="<%=emp_in_date%>" readonly="true">&nbsp;
                </td>
                <th>���������</th>
                <td class="left">
                	<input name="emp_end_gisan" type="text" size="10" id="datepicker2" style="width:70px;" value="<%=emp_end_gisan%>" readonly="true">
                </td>
                <th>�ټӱ����</th>
                <td class="left">
									<input name="emp_gunsok_date" type="text" size="10" id="datepicker3" style="width:70px;" value="<%=emp_gunsok_date%>" readonly="true">
                </td>
                <th>���������</th>
                <td class="left">
									<input name="emp_yuncha_date" type="text" size="10" id="datepicker4" style="width:70px;" value="<%=emp_yuncha_date%>" readonly="true">
                </td>
              </tr>
              <tr>
              	<th colspan="2">�ֹι�ȣ</th>
								<td colspan="2" class="left">
									<input name="emp_person1" type="text" id="emp_person1" size="6" maxlength="6" value="<%=emp_person1%>" >
								  ��
								  <input name="emp_person2" type="text" id="emp_person2" size="7" maxlength="7" value="<%=emp_person2%>" >
                  ����
                  <select name="emp_sex" id="emp_sex" value="<%=emp_sex%>" style="width:50px">
			            	<option value="" <% if emp_sex = "" then %>selected<% end if %>>����</option>
				            <option value='��' <%If emp_sex = "��" then %>selected<% end if %>>��</option>
                    <option value='��' <%If emp_sex = "��" then %>selected<% end if %>>��</option>
                  </select>
                </td>
              	<th>��ȭ��ȣ</th>
								<td colspan="3" class="left">
									<input name="emp_tel_ddd" type="text" id="emp_tel_ddd" size="3" maxlength="3" value="<%=emp_tel_ddd%>" >
								  -
                  <input name="emp_tel_no1" type="text" id="emp_tel_no1" size="4" maxlength="4" value="<%=emp_tel_no1%>" >
                  -
                  <input name="emp_tel_no2" type="text" id="emp_tel_no2" size="4" maxlength="4" value="<%=emp_tel_no2%>" >
                </td>
                <th>�ڵ���</th>
								<td colspan="3" class="left">
									<input name="emp_hp_ddd" type="text" id="emp_hp_ddd" size="3" maxlength="3" value="<%=emp_hp_ddd%>" >
								  -
                  <input name="emp_hp_no1" type="text" id="emp_hp_no1" size="4" maxlength="4" value="<%=emp_hp_no1%>" >
                  -
                  <input name="emp_hp_no2" type="text" id="emp_hp_no2" size="4" maxlength="4" value="<%=emp_hp_no2%>" >
                </td>
              </tr>
              <tr>
              	<th colspan="2" >����(�ּ�)</th>
								<td colspan="7" class="left">
									<input name="emp_family_sido" type="text" id="emp_family_sido" style="width:100px" readonly="true" value="<%=emp_family_sido%>">
              		<input name="emp_family_gugun" type="text" id="emp_family_gugun" style="width:150px" readonly="true" value="<%=emp_family_gugun%>">
              		<input name="emp_family_dong" type="text" id="emp_family_dong" style="width:150px" readonly="true" value="<%=emp_family_dong%>">
              		<input name="emp_family_addr" type="text" id="emp_family_addr" style="width:200px" value="<%=emp_family_addr%>">
              		<input name="emp_family_zip" type="hidden" id="emp_family_zip" value="<%=emp_family_zip%>">
                  <a href="#" class="btnType03" onClick="pop_Window('zipcode_search.asp?gubun=<%="family"%>','family_zip_select','scrollbars=yes,width=600,height=400')">�ּ���ȸ</a>
                </td>
                <th>��󿬶�</th>
								<td colspan="2" class="left">
									<input name="emp_emergency_tel" type="text" id="emp_emergency_tel" size="30" value="<%=emp_emergency_tel%>">
								</td>
              </tr>
              <tr>
								<th colspan="2">�ּ�(��)</th>
								<td colspan="7" class="left">
									<input name="emp_sido" type="text" id="emp_sido" style="width:100px" readonly="true" value="<%=emp_sido%>">
              		<input name="emp_gugun" type="text" id="emp_gugun" style="width:150px" readonly="true" value="<%=emp_gugun%>">
              		<input name="emp_dong" type="text" id="emp_dong" style="width:150px" readonly="true" value="<%=emp_dong%>">
              		<input name="emp_addr" type="text" id="emp_addr" style="width:200px" value="<%=emp_addr%>" >
              		<input name="emp_zipcode" type="hidden" id="emp_zipcode" value="<%=emp_zipcode%>">
              		<a href="#" class="btnType03" onClick="pop_Window('zipcode_search.asp?gubun=<%="juso"%>','family_zip_select','scrollbars=yes,width=600,height=400')">�ּ���ȸ</a>
                </td>
                <th>e-�����ּ�</th>
								<td colspan="2" class="left">
									<input name="emp_email" type="text" id="emp_email" size="12" value="<%=emp_email%>">@k-won.co.kr
                </td>
              </tr>
              <tr>
								<th colspan="2" class="first">�������Կ���</th>
								<td colspan="3" class="left">
									<input type="radio" name="emp_sawo_id" value="Y" <% if emp_sawo_id = "Y" then %>checked<% end if %>>����
              		<input name="emp_sawo_id" type="radio" value="N" <% if emp_sawo_id = "N" then %>checked<% end if %>>����
                </td>
								<th>��ȥ�����</th>
                <td class="left">
                	<input name="emp_marry_date" type="text" size="10" id="datepicker7" style="width:70px;" value="<%=emp_marry_date%>" readonly="true">
                </td>
								<th>���</th>
                <td class="left">
									<input name="emp_hobby" type="text" id="emp_hobby" size="13" value="<%=emp_hobby%>"></td>
                <th>���/���</th>
								<td colspan="2" class="left">
                	<%
                		Sql="select * from emp_etc_code where emp_etc_type = '22' order by emp_etc_code asc"
										Rs_etc.Open Sql, Dbconn, 1
							  	%>
									<select name="emp_disabled" id="emp_disabled" style="width:90px">
                  	<option value="" <% if emp_disabled = "" then %>selected<% end if %>>����</option>
                		<%
											do until rs_etc.eof
			  				  	%>
                		<option value='<%=rs_etc("emp_etc_name")%>' <%If emp_disabled = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		<%
												rs_etc.movenext()
											loop
											rs_etc.Close()
							  		%>
            			</select>
								  -
                  <select name="emp_disab_grade" id="emp_disab_grade" value="<%=emp_disab_grade%>" style="width:50px">
			            	<option value="" <% if emp_disab_grade = "" then %>selected<% end if %>>����</option>
				            <option value='1��' <%If emp_disab_grade = "1��" then %>selected<% end if %>>1��</option>
                    <option value='2��' <%If emp_disab_grade = "2��" then %>selected<% end if %>>2��</option>
                    <option value='3��' <%If emp_disab_grade = "3��" then %>selected<% end if %>>3��</option>
                    <option value='4��' <%If emp_disab_grade = "4��" then %>selected<% end if %>>4��</option>
                    <option value='5��' <%If emp_disab_grade = "5��" then %>selected<% end if %>>5��</option>
                    <option value='6��' <%If emp_disab_grade = "6��" then %>selected<% end if %>>6��</option>
                    <option value='����' <%If emp_disab_grade = "����" then %>selected<% end if %>>����</option>
                    <option value='����' <%If emp_disab_grade = "����" then %>selected<% end if %>>����</option>
                	</select>
                </td>
              </tr>
              <tr>
              	<th colspan="2" >��������</th>
                <td class="left">
                	<%
										Sql="select * from emp_etc_code where emp_etc_type = '06' order by emp_etc_code asc"
										Rs_etc.Open Sql, Dbconn, 1
							  	%>
									<select name="emp_military_id" id="emp_military_id" style="width:90px">
                  	<option value="" <% if emp_military_id = "" then %>selected<% end if %>>����</option>
                		<%
											do until rs_etc.eof
			  				  	%>
                		<option value='<%=rs_etc("emp_etc_name")%>' <%If emp_military_id = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		<%
												rs_etc.movenext()
											loop
											rs_etc.Close()
							  		%>
                	</select>
                </td>
                <th>�������</th>
                <td class="left">
                	<%
										Sql="select * from emp_etc_code where emp_etc_type = '07' order by emp_etc_code asc"
										Rs_etc.Open Sql, Dbconn, 1
							  	%>
									<select name="emp_military_grade" id="emp_military_grade" style="width:90px">
                  	<option value="" <% if emp_military_grade = "" then %>selected<% end if %>>����</option>
                		<%
											do until rs_etc.eof
			  				  	%>
                		<option value='<%=rs_etc("emp_etc_name")%>' <%If emp_military_grade = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		<%
												rs_etc.movenext()
											loop
											rs_etc.Close()
							  		%>
                	</select>
                </td>
                <th>���� �����Ⱓ</th>
                <td colspan="2" class="left">
									<input name="emp_military_date1" type="text" size="10" id="datepicker8" style="width:70px;" value="<%=emp_military_date1%>" readonly="true">
                  ��
                  <input name="emp_military_date2" type="text" size="10" id="datepicker9" style="width:70px;" value="<%=emp_military_date2%>" readonly="true">
                </td>
                <th>��������</th>
								<td class="left">
									<input name="emp_military_comm" type="text" id="emp_military_comm" size="13" value="<%=emp_military_comm%>"></td>
								</td>
                <th>����</th>
                <td class="left">
									<input name="emp_faith" type="text" id="emp_faith" style="width:90px" value="<%=emp_faith%>">
								</td>
							</tr>
							<tr>
              	<th colspan="2" class="first">�Ǳٹ���/�ּ�</th>
                <td colspan="3" class="left">
                	<input name="emp_stay_name" type="text" id="emp_stay_name" size="30"  value="<%=emp_stay_name%>">
                	<a href="#" class="btnType03" onClick="pop_Window('insa_stay_select.asp?gubun=<%="stay"%>&reside_code=<%=emp_stay_code%>','stayselect','scrollbars=yes,width=1000,height=400')">����</a>
                </td>
                <td colspan="5" class="left">
                	<%
                		if emp_stay_code <> "" then
								   		Sql="select * from emp_stay where stay_code = '"&emp_stay_code&"'"
								   		Rs_stay.Open Sql, Dbconn, 1

							    	'do until rs_stay.eof
							    		if not rs_stay.eof then

								      	emp_stay_name = rs_stay("stay_name")
								      	stay_sido = rs_stay("stay_sido")
								      	stay_gugun = rs_stay("stay_gugun")
								      	stay_dong = rs_stay("stay_dong")
								      	stay_addr = rs_stay("stay_addr")
								      	'	rs_stay.movenext()
								      	'loop
								     	end if
								     	rs_stay.Close()
										end if
							  	%>
							  	<input name="emp_stay_code" type="text" id="emp_stay_code" size="4" readonly="true" value="<%=emp_stay_code%>">
                  ~~
                  <input name="stay_sido" type="text" id="stay_sido" style="width:90px" readonly="true" value="<%=stay_sido%>">
                  <input name="stay_gugun" type="text" id="stay_gugun" style="width:90px" readonly="true" value="<%=stay_gugun%>">
                  <input name="stay_dong" type="text" id="stay_dong" style="width:90px" readonly="true" value="<%=stay_dong%>">
                  <input name="stay_addr" type="text" id="stay_addr" style="width:150px" readonly="true" value="<%=stay_addr%>">
								</td>
                <th>���׷�</th>
                <td class="left">
                	<input name="cost_group" type="text" id="cost_group" style="width:90px" readonly="true" value="<%=cost_group%>">
            		</td>
              </tr>
              <tr>
              	<th colspan="2" class="first">������ȣ</th>
                <td colspan="2" class="left">
                	<input name="emp_extension_no" type="text" id="emp_extension_no" size="16 " value="<%=emp_extension_no%>">
                </td>
                <th>�����з�</th>
                <td colspan="2" class="left">
                	<select name="emp_last_edu" id="emp_last_edu" value="<%=emp_last_edu%>" style="width:100px">
			            	<option value="" <% if emp_last_edu = "" then %>selected<% end if %>>����</option>
				            <option value='����б�' <%If emp_last_edu = "����б�" then %>selected<% end if %>>����б�</option>
                    <option value='������' <%If emp_last_edu = "������" then %>selected<% end if %>>������</option>
                    <option value='���б�' <%If emp_last_edu = "���б�" then %>selected<% end if %>>���б�</option>
                    <option value='���п�����' <%If emp_last_edu = "���п�����" then %>selected<% end if %>>���п�����</option>
                    <option value='���п�' <%If emp_last_edu = "���п�" then %>selected<% end if %>>���п�</option>
                	</select>
                </td>
                <th>��뱸��</th>
                <td class="left">
                	<%
                		Sql="select * from emp_etc_code where emp_etc_type = '70' order by emp_etc_code asc"
										Rs_etc.Open Sql, Dbconn, 1
							  	%>
									<select name="cost_center" id="cost_center" style="width:90px">
                  	<option value="" <% if cost_center = "" then %>selected<% end if %>>����</option>
                		<%
                			do until rs_etc.eof
			  				  	%>
                		<option value='<%=rs_etc("emp_etc_name")%>' <%If cost_center = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		<%
												rs_etc.movenext()
											loop
											rs_etc.Close()
							  		%>
                	</select>
                </td>
                <th>�����׷쿩��</th>
                <td colspan="2" class="left">
									<input type="radio" name="mg_group" value="1" <% if mg_group = "1" then %>checked<% end if %>>�Ϲݱ׷�
              		<input name="mg_group" type="radio" value="2" <% if mg_group = "2" then %>checked<% end if %>>�����׷�
                </td>
              </tr>
              <tr>
              	<th colspan="2" class="first">�Է���</th>
                <td colspan="2" class="left"><%=emp_reg_date%>&nbsp;(<%=emp_reg_user%>)</td>
                <th>������</th>
                <td colspan="2" class="left"><%=emp_mod_date%>&nbsp;(<%=emp_mod_user%>)</td>
                <th>����ó ȸ��</th>
								<td colspan="2" class="left"><input name="emp_reside_company" type="text" id="emp_reside_company" style="width:90px"  value="<%=emp_reside_company%>">
									<a href="#" class="btnType03" onClick="pop_Window('insa_trade_search.asp?gubun=<%="5"%>','tradesearch','scrollbars=yes,width=600,height=400')">ã��</a>
            		</td>
                <th>�޿����</th>
                <td class="left">
                	<select name="emp_pay_id" id="emp_pay_id" value="<%=emp_pay_id%>" style="width:90px">
			            	<option value="" <% if emp_pay_id = "" then %>selected<% end if %>>����</option>
				            <option value='0' <%If emp_pay_id = "0" then %>selected<% end if %>>����</option>
                    <option value='1' <%If emp_pay_id = "1" then %>selected<% end if %>>����</option>
                    <option value='2' <%If emp_pay_id = "2" then %>selected<% end if %>>����</option>
                    <option value='3' <%If emp_pay_id = "3" then %>selected<% end if %>>¡��</option>
                    <option value='5' <%If emp_pay_id = "5" then %>selected<% end if %>>����</option>
                  </select>
                </td>
              </tr>
						</tbody>
					</table>
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th scope="row">�������</th>
								<td class="left">
									<input type="file" name= "att_file"  size="70" accept="image/gif"> * ÷�������� 1���� �����ϸ� �ִ�뷮�� 2MB
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
        <input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">
        <input type="hidden" name="emp_end_date" value="<%=emp_end_date%>" ID="Hidden1">
        <input type="hidden" name="emp_org_baldate" value="<%=emp_org_baldate%>" ID="Hidden1">
        <input type="hidden" name="emp_grade_date" value="<%=emp_grade_date%>" ID="Hidden1">
        <input type="hidden" name="v_att_file" value="<%=att_file%>" ID="Hidden1">
			</form>
		</div>
	</div>
	</body>
</html>

