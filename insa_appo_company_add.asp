<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)

u_type = request("u_type")
emp_no = request("emp_no")
emp_name = request("emp_name")
be_pg = request("be_pg")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Set rs_into = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

app_seq = ""
app_id = "�迭����"
app_date = ""
app_id_type = ""
app_to_company = ""
app_to_org = ""
app_to_grade = ""
app_to_job = ""
app_to_grade = ""
app_to_enddate = ""
app_be_company = ""
app_be_org = ""
app_be_grade = ""
app_be_job = ""
app_be_grade = ""
app_be_enddate = ""
app_first_date = ""
app_end_date = ""
app_comment = ""

if u_type = "U" then

	Sql="select * from emp_master where emp_no = '"&emp_no&"'"
	Set rs_emp=DbConn.Execute(Sql)

	emp_name = rs_emp("emp_name")
    emp_ename = rs_emp("emp_ename")
    emp_type = rs_emp("emp_type")
    emp_sex = rs_emp("emp_sex")
    emp_person1 = rs_emp("emp_person1")
    emp_person2 = rs_emp("emp_person2")
    emp_image = rs_emp("emp_image")
    emp_first_date = rs_emp("emp_first_date")
    emp_in_date = rs_emp("emp_in_date")
    emp_gunsok_date = rs_emp("emp_gunsok_date")
    emp_yuncha_date = rs_emp("emp_yuncha_date")
    emp_end_gisan = rs_emp("emp_end_gisan")
    emp_end_date = rs_emp("emp_end_date")
    emp_company = rs_emp("emp_company")
    emp_bonbu = rs_emp("emp_bonbu")
    emp_saupbu = rs_emp("emp_saupbu")
    emp_team = rs_emp("emp_team")
    emp_org_code = rs_emp("emp_org_code")
    emp_org_name = rs_emp("emp_org_name")
    emp_org_baldate = rs_emp("emp_org_baldate")
    emp_stay_code = rs_emp("emp_stay_code")
	emp_stay_name = rs_emp("emp_stay_name")
    emp_reside_place = rs_emp("emp_reside_place")
	emp_reside_company = rs_emp("emp_reside_company")
    emp_grade = rs_emp("emp_grade")
    emp_grade_date = rs_emp("emp_grade_date")
    emp_job = rs_emp("emp_job")
    emp_position = rs_emp("emp_position")
    emp_jikgun = rs_emp("emp_jikgun")
    emp_jikmu = rs_emp("emp_jikmu")
    emp_birthday = rs_emp("emp_birthday")
    emp_birthday_id = rs_emp("emp_birthday_id")
    emp_family_zip = rs_emp("emp_family_zip")
    emp_family_sido = rs_emp("emp_family_sido")
    emp_family_gugun = rs_emp("emp_family_gugun")
    emp_family_dong = rs_emp("emp_family_dong")
    emp_family_addr = rs_emp("emp_family_addr")
    emp_zipcode = rs_emp("emp_zipcode")
    emp_sido = rs_emp("emp_sido")
    emp_gugun = rs_emp("emp_gugun")
    emp_dong = rs_emp("emp_dong")
    emp_addr = rs_emp("emp_addr")
    emp_tel_ddd = rs_emp("emp_tel_ddd")
    emp_tel_no1 = rs_emp("emp_tel_no1")
    emp_tel_no2 = rs_emp("emp_tel_no2")
    emp_hp_ddd = rs_emp("emp_hp_ddd")
    emp_hp_no1 = rs_emp("emp_hp_no1")
    emp_hp_no2 = rs_emp("emp_hp_no2")
    emp_email = rs_emp("emp_email")
    emp_military_id = rs_emp("emp_military_id")
    emp_military_date1 = rs_emp("emp_military_date1")
    emp_military_date2 = rs_emp("emp_military_date2")
    emp_military_grade = rs_emp("emp_military_grade")
    emp_military_comm = rs_emp("emp_military_comm")
    emp_hobby = rs_emp("emp_hobby")
    emp_faith = rs_emp("emp_faith")
    emp_last_edu = rs_emp("emp_last_edu")
    emp_marry_date = rs_emp("emp_marry_date")
	emp_disabled_yn = rs_emp("emp_disabled_yn")
    emp_disabled = rs_emp("emp_disabled")
    emp_disab_grade = rs_emp("emp_disab_grade")
    emp_sawo_id = rs_emp("emp_sawo_id")
    emp_sawo_date = rs_emp("emp_sawo_date")
    emp_emergency_tel = rs_emp("emp_emergency_tel")
    emp_nation_code = rs_emp("emp_nation_code")
	cost_center = rs_emp("cost_center")
	cost_group = rs_emp("cost_group")

	photo_image = "/emp_photo/" + rs_emp("emp_image")
	emp_email = emp_email + "@k-won.co.kr"
	
	if emp_person2 <> "" then
	   sex_id = mid(cstr(emp_person2),1,1)
	   if sex_id = "1" then
	         emp_sex = "��"
		  else
		     emp_sex = "��"
	   end if
	end if

	if rs_emp("emp_org_baldate") = "1900-01-01" then
	   emp_org_baldate = ""
	   else 
	   emp_org_baldate = rs_emp("emp_org_baldate")
	end if
	if rs_emp("emp_grade_date") = "1900-01-01" then
	   emp_grade_date = ""
	   else 
	   emp_grade_date = rs_emp("emp_grade_date")
	end if
	if rs_emp("emp_sawo_date") = "1900-01-01" then
	   emp_sawo_date = ""
	   else 
	   emp_sawo_date = rs_emp("emp_sawo_date")
	end if
	
	sql="select * from memb where user_id='"&emp_no&"'"
	set rs_memb=dbconn.execute(sql)
	if not rs_memb.eof then
	       mg_group = rs_memb("mg_group")
	   else
	       mg_group = "1"
    end if
	rs_memb.close()
end if

    'sql="select max(emp_no) as max_seq from emp_master"
	sql="select max(emp_no) as max_seq from emp_master where emp_no < '900000'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_last = "000001"
	  else
		max_seq = "000000" + cstr((int(rs_max("max_seq")) + 1))
		code_last = right(max_seq,6)
	end if
    rs_max.close()
	
    new_emp_no = code_last


title_line = " �迭���� �λ� �߷�ó�� "

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
				return "2 1";
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=app_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
											$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker1" ).datepicker("setDate", "<%=emp_end_gisan%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
											$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker2" ).datepicker("setDate", "<%=emp_gunsok_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
											$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker3" ).datepicker("setDate", "<%=emp_yuncha_date%>" );
			});	  
			$(function() {    $( "#datepicker4" ).datepicker();
											$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker4" ).datepicker("setDate", "<%=app_distart_date%>" );
			});	  
			$(function() {    $( "#datepicker5" ).datepicker();
											$( "#datepicker5" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker5" ).datepicker("setDate", "<%=app_difinish_date%>" );
			});	  
			function frmcheck () {
				if (chkfrm() && formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.app_date.value =="") {
					alert('�迭�������� �Է��ϼ���');
					frm.app_date.focus();
					return false;}
				if(document.frm.app_be_orgcode.value =="") {
					alert('�߷ɼҼ��� �Է��ϼ���');
					frm.app_be_orgcode.focus();
					return false;}			
				
				{
				a=confirm('�迭�߷��� �Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_appoint_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_appo_company_addsave.asp" method="post" name="frm">
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
                                <td class="left"><%=emp_no%>&nbsp;</td>
                                <th>����(�ѱ�)</th>
                                <td class="left"><%=emp_name%>&nbsp;</td>
								<th>����(����)</th>
								<td colspan="2" class="left"><%=emp_ename%>&nbsp;</td>
                                <th>�������</th>
                                <td colspan="2" class="left"><%=emp_birthday%>&nbsp;(<%=emp_birthday_id%>)&nbsp;</td>
                           </tr>
                           <tr>   
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td colspan="3" class="left"><%=emp_org_code%>&nbsp;��&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                                <td colspan="5" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
                           </tr>
                           <tr>
                                <th>��������</th>
                                <td class="left"><%=emp_type%>&nbsp;</td>
                               	<th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_grade%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_job%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;å</th>
                                <td class="left"><%=emp_position%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=emp_jikmu%>&nbsp;</td>
                           </tr>
                           <tr>     
                                <th>�����Ի���</th>
                                <td class="left"><%=emp_first_date%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��</th>
                                <td class="left"><%=emp_in_date%>&nbsp;</td>
                                <th>���������</th>
                                <td class="left"><%=emp_end_gisan%>&nbsp;</td>
                                <th>�ټӱ����</th>
                                <td class="left"><%=emp_gunsok_date%>&nbsp;</td>
                                <th>���������</th>
                                <td class="left"><%=emp_yuncha_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="2">�ֹι�ȣ</th>
								<td colspan="2" class="left"><%=emp_person1%>&nbsp;��&nbsp;<%=emp_person2%>&nbsp;(<%=emp_sex%>)&nbsp;</td>
                                <th>��ȭ��ȣ</th>
								<td colspan="3" class="left"><%=emp_tel_ddd%>��<%=emp_tel_no1%>��<%=emp_tel_no2%>&nbsp;</td>
                                <th>�ڵ���</th>
								<td class="left"><%=emp_hp_ddd%>��<%=emp_hp_no1%>��<%=emp_hp_no2%>&nbsp;</td>
                                <th>������</th>
								<td class="left"><%=emp_end_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="2">���Ҽӹ߷���</th>
								<td class="left"><%=emp_org_baldate%>&nbsp;</td>
                                <th>�����޽�����</th>
								<td class="left"><%=emp_grade_date%>&nbsp;</td>
                                <th>e_�����ּ�</th>
								<td colspan="2" class="left"><%=emp_email%>&nbsp;</td>
                                <th>��������</th>
								<td colspan="3" class="left">
                                <input type="radio" name="emp_sawo_id" value="Y" <% if emp_sawo_id = "Y" then %>checked<% end if %>>���� 
              					<input name="emp_sawo_id" type="radio" value="N" <% if emp_sawo_id = "N" then %>checked<% end if %>>����
                                 &nbsp;&nbsp;��&nbsp;&nbsp;<%=emp_sawo_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="12" class="left" style="background:#FFC">�� �迭���� �λ�߷� ��</th>&nbsp;
                            </tr>
                            <tr>                            
                                <th colspan="2" class="first">�迭��������</th>
                                <td colspan="2" class="left">
                                <input name="app_date" type="text" size="10" readonly="true" id="datepicker" style="width:70px;">&nbsp;</td>
                                <th>�迭���� ���</th>
                                <td class="left"><%=new_emp_no%><input name="new_emp_no" type="hidden" value="<%=new_emp_no%>"></td>
                                <th>���������</th>
                                <td class="left">
                                <input name="emp_end_gisan" type="text" size="10" id="datepicker1" style="width:70px;" value="<%=emp_end_gisan%>">
                                </td>
                                <th>�ټӱ����</th>
                                <td class="left">
								<input name="emp_gunsok_date" type="text" size="10" id="datepicker2" style="width:70px;" value="<%=emp_gunsok_date%>">
                                </td>
                                <th>���������</th>
                                <td class="left">
								<input name="emp_yuncha_date" type="text" size="10" id="datepicker3" style="width:70px;" value="<%=emp_yuncha_date%>">
                            </tr>    
              <% '�߷ɱ��к� �޴� ���� %>
							<tr style="" id="mv_menu1">
								<th colspan="2" class="first" >���Ҽ�</th>
								<td colspan="3" class="left"><%=emp_org_code%>&nbsp;��&nbsp;<%=emp_org_name%>&nbsp;</td>
                                <th class="first" >������</th>
                                <td colspan="6" class="left"><%=emp_company%>&nbsp;&nbsp;<%=emp_bonbu%>&nbsp;&nbsp;<%=emp_saupbu%>&nbsp;&nbsp;<%=emp_team%>&nbsp;&nbsp;<%=emp_reside_place%>&nbsp;</td>
							</tr>
                            <tr style="" id="mv_menu2">
								<th colspan="2" class="first" style="background:#FFC">�߷ɼҼ�</th>
								<td colspan="3" class="left">
								<input name="app_be_orgcode" type="text" id="app_be_orgcode" style="width:40px" readonly="true" value="<%=app_be_orgcode%>">
                                &nbsp;��&nbsp;
                                <input name="app_be_org" type="text" id="app_be_org" style="width:120px" readonly="true" value="<%=app_be_org%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="apporg"%>&mg_org=<%=mg_org%>','orgselect','scrollbars=yes,width=800,height=400')">����</a>
                                </td>
                                <th style="background:#FFC">�߷�����</th>
								<td colspan="6" class="left">
                                <input name="app_company" type="text" id="app_company" style="width:100px" readonly="true" value="<%=app_company%>">
              					<input name="app_bonbu" type="text" id="app_bonbu" style="width:120px" readonly="true" value="<%=app_bonbu%>">
              					<input name="app_saupbu" type="text" id="app_saupbu" style="width:120px" readonly="true" value="<%=app_saupbu%>">
              					<input name="app_team" type="text" id="app_team" style="width:120px" readonly="true" value="<%=app_team%>">
                                <input name="app_reside_place" type="hidden" id="app_reside_place" style="width:120px" readonly="true" value="<%=app_reside_place%>">
                                <input name="app_reside_company" type="hidden" id="app_reside_company" style="width:120px" readonly="true" value="<%=app_reside_company%>">
                                <input name="app_org_level" type="hidden" id="app_org_level" style="width:120px" readonly="true" value="<%=app_org_level%>">
                                </td>
                            </tr>
                            <%
								stay_name = emp_stay_name
								if emp_stay_code <> "" then
								   Sql="select * from emp_stay where stay_code = '"&emp_stay_code&"'"
								   Rs_stay.Open Sql, Dbconn, 1
							  
							    	'do until rs_stay.eof 
							    	if not rs_stay.eof then
								
								       stay_name = rs_stay("stay_name")
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
                            
                            <tr style="" id="mv_menu3">  
                                <th colspan="2" class="first" style="background:#FFC">�Ǳٹ���/�ּ�</th>
                                <td colspan="3" class="left">
                                <input name="emp_stay_code" type="text" id="emp_stay_code" style="width:40px" readonly="true" value="<%=emp_stay_code%>">
                                &nbsp;��&nbsp;
                                <input name="stay_name" type="text" id="stay_name" style="width:150px" readonly="true" value="<%=stay_name%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_stay_select.asp?gubun=<%="stay"%>&reside_code=<%=emp_stay_code%>','stayselect','scrollbars=yes,width=1000,height=400')">����</a>
                                </td>
                                <th>�ּ���</th>
                                <td colspan="6" class="left">
                                <input name="stay_sido" type="text" id="stay_sido" style="width:100px" readonly="true" value="<%=stay_sido%>">
                                <input name="stay_gugun" type="text" id="stay_gugun" style="width:150px" readonly="true" value="<%=stay_gugun%>">
                                <input name="stay_dong" type="text" id="stay_dong" style="width:150px" readonly="true" value="<%=stay_dong%>">
                                <input name="stay_addr" type="text" id="stay_addr" style="width:200px" readonly="true" value="<%=stay_addr%>">
								</td>
                            </tr>
                            <tr style="" id="mv_menu4">
                                <th colspan="2" class="first" style="background:#FFC">����</th>
                                <td colspan="3" class="left">
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
								<th style="background:#FFC">�߷ɻ���</th>
								<td colspan="6" class="left">
								<input name="app_mv_comment" type="text" id="app_mv_comment" style="width:500px" onKeyUp="checklength(this,50)" value="<%=app_comment%>">
                                </td>
                            </tr>
                            <tr style="" id="mv_menu5">
                                <th colspan="2" class="first" style="background:#FFC">���׷�</th>
                                <td colspan="3" class="left">
                                <input name="app_cost_group" type="text" id="app_cost_group" style="width:90px" readonly="true" value="<%=cost_group%>">
            					</td>
								<th style="background:#FFC">��뱸��</th>
								<td colspan="2" class="left">
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
                                <th style="background:#FFC">�����׷쿩��</th>
								<td colspan="3" class="left">
								<input type="radio" name="mg_group" value="1" <% if mg_group = "1" then %>checked<% end if %>>�Ϲݱ׷� 
              					<input name="mg_group" type="radio" value="2" <% if mg_group = "2" then %>checked<% end if %>>�����׷�
                                </td>                                
                            </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:goBefore();"></span>
                </div>
                <input name="be_pg" type="hidden" id="be_pg" value="<%=be_pg%>">
                <input type="hidden" name="app_id" value="<%=app_id%>" ID="Hidden1">
                <input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="emp_name" value="<%=emp_name%>" ID="Hidden1">
                <input type="hidden" name="app_grade" value="<%=emp_grade%>" ID="Hidden1">
                <input type="hidden" name="app_position" value="<%=emp_position%>" ID="Hidden1">
                <input type="hidden" name="app_job" value="<%=emp_job%>" ID="Hidden1">
                <input type="hidden" name="app_to_company" value="<%=emp_company%>" ID="Hidden1">
                <input type="hidden" name="app_to_bonbu" value="<%=emp_bonsu%>" ID="Hidden1">
                <input type="hidden" name="app_to_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="app_to_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="app_org" value="<%=emp_org_code%>" ID="Hidden1">
                <input type="hidden" name="app_org_name" value="<%=emp_org_name%>" ID="Hidden1">
                <input type="hidden" name="cost_center" value="<%=cost_center%>" ID="Hidden1">
                <input type="hidden" name="cost_group" value="<%=cost_group%>" ID="Hidden1">
        	</form>
		</div>				
	</div>        				
	</body>
</html>

