<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)
curr_hh = int(cstr(datepart("h",now)))
curr_mm = int(cstr(datepart("n",now)))

' �Է¹޾� ����Ÿ�� ��Ƶ� �ʵ��̸��� ���ǿ� �⺻���� null�� ����Ѱ�

u_type = request("u_type")
u_type = "U"
'emp_no = request("emp_no")
in_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

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
emp_reside_place = ""
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
emp_disabled = ""
emp_disab_grade = ""
emp_sawo_id = "N"
emp_sawo_date = ""
emp_emergency_tel = ""
emp_nation_code = ""
att_file = ""

emp_mod_date = now()
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
Dbconn.open dbconnect

Sql="select * from emp_master where emp_no = '"&emp_no&"'"
Set rs=DbConn.Execute(Sql)

'response.write(Sql)

if not rs.EOF or not rs.BOF then

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
	emp_extension_no = rs("emp_extension_no")
'   end_date = mid(cstr(now()),1,10)
    reg_user = rs("emp_reg_user")
    mod_user = rs("emp_mod_user")
	if rs("emp_military_date1") = "1900-01-01" then
           emp_military_date1 = ""
           emp_military_date2 = ""
    end if
    if rs("emp_marry_date") = "1900-01-01" then
           emp_marry_date = ""
    end if
	if rs("emp_birthday") = "1900-01-01" then
           emp_birthday = ""
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


title_line = "[ �λ�⺻���� ���� ]"

photo_image = "/emp_photo/" + rs("emp_image")
att_file = photo_image
'response.write(att_file)
  else
    response.write"<script language=javascript>"
	response.write"alert('��ϵ� ����� �ƴմϴ�. �ٽ��ѹ� Ȯ���� �ֽʽÿ�');"
	response.write"location.replace('insa_person_mg.asp');"
	response.write"</script>"
	Response.End
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
			function getPageCode(){
				return "0 1";
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
				if(document.frm.emp_ename.value =="") {
					alert('���������� �Է��ϼ���');
					frm.emp_ename.focus();
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
				if(document.frm.emp_family_sido.value =="") {
					alert('�����ּҸ� ��ȸ �ϼ���');
					return false;}
				if(document.frm.emp_family_addr.value =="") {
					alert('������ �Է��ϼ���');
					frm.emp_family_addr.focus();
					return false;}	
				if(document.frm.emp_sido.value =="") {
					alert('���ּҸ� ��ȸ �ϼ���');
					return false;}
				if(document.frm.emp_addr.value =="") {
					alert('���ּ� ������ �Է��ϼ���');
					frm.emp_addr.focus();
					return false;}		
				if(document.frm.emp_email.value =="") {
					alert('��-�����ּҸ� �Է��ϼ���');
					frm.emp_email.focus();
					return false;}			
				if(document.frm.emp_emergency_tel.value =="") {
					alert('��󿬶� ��ȭ��ȣ�� �Է��ϼ���');
					frm.emp_emergency_tel.focus();
					return false;}			
//				if(document.frm.emp_extension_no.value =="") {
//					alert('������ȣ�� �Է��ϼ���');
//					frm.emp_extension_no.focus();
//					return false;}	
				if(document.frm.emp_last_edu.value =="") {
					alert('�����з��� �Է��ϼ���');
					frm.emp_last_edu.focus();
					return false;}	
				if(document.frm.v_att_file.value =="") 
  				    if(document.frm.att_file.value =="") {
					   alert('������ ��� �ϼ���');
					   frm.att_file.focus();
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
	<body>
    <%
    '<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">
	%>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_open_emp_save.asp" method="post" name="frm" enctype="multipart/form-data">
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
                                <td class="left"><%=emp_no%>&nbsp;<input name="emp_no" type="hidden" value="<%=emp_no%>"></td>
                                <th>����(�ѱ�)</th>
                                <td class="left"><%=emp_name%>&nbsp;<input name="emp_name" type="hidden" value="<%=emp_name%>"></td>
								<th>����(����)</th>
								<td colspan="2" class="left">
                                <input name="emp_ename" type="text" id="emp_ename" style="width:160px" maxlength="20" value="<%=emp_ename%>"></td>
                                <th>�������</th>
                                <td colspan="2" class="left">
								<input name="emp_birthday" type="text" size="10" id="datepicker5" style="width:70px;" value="<%=emp_birthday%>" readonly="true">
                                &nbsp;��&nbsp;
                                <input type="radio" name="emp_birthday_id" value="��" <% if emp_birthday_id = "��" then %>checked<% end if %>>�� 
              					<input name="emp_birthday_id" type="radio" value="��" <% if emp_birthday_id = "��" then %>checked<% end if %>>��
                                </td>
                            </tr>   
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td colspan="3" class="left"><%=emp_org_code%>&nbsp;��&nbsp;<%=emp_org_name%>
                                <input name="emp_org_code" type="hidden" id="emp_org_code" style="width:40px" readonly="true" value="<%=emp_org_code%>">
                                <input name="emp_org_name" type="hidden" id="emp_org_name" style="width:120px" readonly="true" value="<%=emp_org_name%>">
                                </td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                                <td colspan="5" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;(<%=emp_reside_company%>)
                                <input name="emp_company" type="hidden" id="emp_company" style="width:100px" readonly="true" value="<%=emp_company%>">
              					<input name="emp_bonbu" type="hidden" id="emp_bonbu" style="width:120px" readonly="true" value="<%=emp_bonbu%>">
              					<input name="emp_saupbu" type="hidden" id="emp_saupbu" style="width:120px" readonly="true" value="<%=emp_saupbu%>">
              					<input name="emp_team" type="hidden" id="emp_team" style="width:120px" readonly="true" value="<%=emp_team%>">
                                <input name="emp_reside_place" type="hidden" id="emp_reside_place" style="width:120px" readonly="true" value="<%=emp_reside_place%>">
                                <input name="emp_reside_company" type="hidden" id="emp_reside_company" style="width:120px" readonly="true" value="<%=emp_reside_company%>">
                                </td>
                            </tr>
                                <th>��������</th>
                                <td class="left"><%=emp_type%>&nbsp;</td>
                               	<th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_grade%>&nbsp;<input name="emp_grade" type="hidden" value="<%=emp_grade%>"></td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_job%>&nbsp;<input name="emp_job" type="hidden" value="<%=emp_job%>"></td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;å</th>
                                <td class="left"><%=emp_position%>&nbsp;<input name="emp_position" type="hidden" value="<%=emp_position%>"></td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_jikmu%>&nbsp;<input name="emp_jikmu" type="hidden" value="<%=emp_jikmu%>"></td>
                           </tr>
                           <tr>     
                                <th>�����Ի���</th>
                                <td class="left"><%=emp_first_date%>&nbsp;
                                <input name="emp_first_date" type="hidden" size="10"  style="width:70px;" readonly="true" value="<%=emp_first_date%>">&nbsp;
                                </td>
                                <th>��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��</th>
                                <td class="left"><%=emp_in_date%>&nbsp;
								<input name="emp_in_date" type="hidden" size="10" style="width:70px;" readonly="true" value="<%=emp_in_date%>">&nbsp;
                                </td>
                                <th>���������</th>
                                <td class="left"><%=emp_end_gisan%>&nbsp;
                                <input name="emp_end_gisan" type="hidden" size="10" style="width:70px;" readonly="true" value="<%=emp_end_gisan%>">
                                </td>
                                <th>�ټӱ����</th>
                                <td class="left"><%=emp_gunsok_date%>&nbsp;
								<input name="emp_gunsok_date" type="hidden" size="10" style="width:70px;" readonly="true" value="<%=emp_gunsok_date%>">
                                </td>
                                <th>���������</th>
                                <td class="left"><%=emp_yuncha_date%>&nbsp;
								<input name="emp_yuncha_date" type="hidden" size="10" style="width:70px;" readonly="true" value="<%=emp_yuncha_date%>">
                                </td>
                            </tr>
                            <tr>
                                <th colspan="2">�ֹι�ȣ</th>
								<td colspan="2" class="left"><%=emp_person1%>-<%=emp_person2%>&nbsp;(<%=emp_sex%>)&nbsp;
                                <input name="emp_person1" type="hidden" id="emp_person1" size="6" maxlength="6" value="<%=emp_person1%>" >
                                <input name="emp_person2" type="hidden" id="emp_person2" size="7" maxlength="7" value="<%=emp_person2%>" >
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
              					<input name="emp_family_addr" type="text" id="emp_family_addr" style="width:200px" value="<%=emp_family_addr%>" notnull errname="����" onKeyUp="checklength(this,50)">
              					<input name="emp_family_zip" type="hidden" id="emp_family_zip" value="<%=emp_family_zip%>">
                                <a href="#" class="btnType03" onClick="pop_Window('zipcode_search.asp?gubun=<%="family"%>','family_zip_select','scrollbars=yes,width=600,height=400')">�ּ���ȸ</a>
                                </td>
                                <th>��󿬶�</th>
								<td colspan="2" class="left">
								<input name="emp_emergency_tel" type="text" id="emp_emergency_tel" size="30" value="<%=emp_emergency_tel%>" onKeyUp="checklength(this,13)"></td>
                            </tr>
                            <tr>
								<th colspan="2">�ּ�(��)</th>
								<td colspan="7" class="left">
								<input name="emp_sido" type="text" id="emp_sido" style="width:100px" readonly="true" value="<%=emp_sido%>">
              					<input name="emp_gugun" type="text" id="emp_gugun" style="width:150px" readonly="true" value="<%=emp_gugun%>">
              					<input name="emp_dong" type="text" id="emp_dong" style="width:150px" readonly="true" value="<%=emp_dong%>">
              					<input name="emp_addr" type="text" id="emp_addr" style="width:200px" value="<%=emp_addr%>" notnull errname="����" onKeyUp="checklength(this,50)">
              					<input name="emp_zipcode" type="hidden" id="emp_zipcode" value="<%=emp_zipcode%>">
              					<a href="#" class="btnType03" onClick="pop_Window('zipcode_search.asp?gubun=<%="juso"%>','family_zip_select','scrollbars=yes,width=600,height=400')">�ּ���ȸ</a>
                                </td>
                                <th>e-�����ּ�</th>
								<td colspan="2" class="left">
								<input name="emp_email" type="text" id="emp_email" size="12" value="<%=emp_email%>">
                                @k-won.co.kr
                                </td>
                            </tr>
                         	<tr>
                                <th colspan="2" class="first">�������Կ���</th>
                            <%
							    if rs("emp_sawo_id") = "Y" then
								      sawo_id = "����"
								   else
								      sawo_id = "����"
							    end if
							%>
                                <td class="left"><%=sawo_id%>&nbsp;</td>
                                </td>
								<th>����������</th>
                                <td class="left"><%=emp_sawo_date%>&nbsp;<input name="emp_sawo_date" type="hidden" value="<%=emp_sawo_date%>"></td>
								<th>��ȥ�����</th>
                                <td class="left">
                                <input name="emp_marry_date" type="text" size="10" id="datepicker7" style="width:70px;" value="<%=emp_marry_date%>" readonly="true"></td>
								<th>���</th>
                                <td class="left">
								<input name="emp_hobby" type="text" id="emp_hobby" size="13" value="<%=emp_hobby%>"></td>
                                <th>���/���</th>
								<td colspan="2" class="left"><%=emp_disabled%> - <%=emp_disab_grade%>&nbsp;
                                <input name="emp_disabled" type="hidden" value="<%=emp_disabled%>">
                                <input name="emp_disab_grade" type="hidden" value="<%=emp_disab_grade%>">
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
								<input name="emp_military_comm" type="text" id="emp_military_comm" size="13" value="<%=emp_military_comm%>"></td></td>
                                <th>����</th>
                                <td class="left">
								<input name="emp_faith" type="text" id="emp_faith" size="13" value="<%=emp_faith%>"></td></td>
							</tr>
                            <tr>
                        		<th colspan="2" class="first">������ȣ</th>
                                <td colspan="2" class="left"><input name="emp_extension_no" type="text" id="emp_extension_no" size="16 " value="<%=emp_extension_no%>">
                                </td>
                                <th>�����з�</th>
                                <td colspan="3" class="left">
                                <select name="emp_last_edu" id="emp_last_edu" value="<%=emp_last_edu%>" style="width:100px">
			            	        <option value="" <% if emp_last_edu = "" then %>selected<% end if %>>����</option>
				                    <option value='����б�' <%If emp_last_edu = "����б�" then %>selected<% end if %>>����б�</option>
                                    <option value='������' <%If emp_last_edu = "������" then %>selected<% end if %>>������</option>
                                    <option value='���б�' <%If emp_last_edu = "���б�" then %>selected<% end if %>>���б�</option>
                                    <option value='���п�����' <%If emp_last_edu = "���п�����" then %>selected<% end if %>>���п�����</option>
                                    <option value='���п�' <%If emp_last_edu = "���п�" then %>selected<% end if %>>���п�</option>
                                </select>
                                </td>
                                <th>�����׷쿩��</th>
                                <td colspan="3" class="left">
								<input type="radio" name="mg_group" value="1" <% if mg_group = "1" then %>checked<% end if %>>�Ϲݱ׷� 
              					<input name="mg_group" type="radio" value="2" <% if mg_group = "2" then %>checked<% end if %>>�����׷�
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
								<input type="file" name= "att_file" size="70" accept="image/gif"> * ÷�������� 1���� �����ϸ� �ִ�뷮�� 2MB
                                </td>
							</tr>              
						</tbody>
                    </table>                    
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <div class="btnCenter">
                         <span class="btnType01"><input type="button" value="�Է�" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                         <span class="btnType01"><input type="button" value="����" onclick="javascript:goBefore();"></span>
                    </div>
                    </td>
				    <td width="52%">
					<div class="btnCenter">
                    <a href="#" class="btnType04">�� �������� �� �з»��� �� ��»��� �� �ڰݻ��� �� �������� �� ���дɷ��� ����Ͻñ� �ٶ��ϴ�</a>
					</div>                  
                    </td>
			      </tr>
				  </table>                
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="emp_end_date" value="<%=emp_end_date%>" ID="Hidden1">
                <input type="hidden" name="emp_org_baldate" value="<%=emp_org_baldate%>" ID="Hidden1">
                <input type="hidden" name="emp_grade_date" value="<%=emp_grade_date%>" ID="Hidden1">
                <input type="hidden" name="v_att_file" value="<%=att_file%>" ID="Hidden1">
				</form>
		</div>				
	</div>        				
	</body>
</html>

