<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
'car_no = ""
car_no = request("car_no")

car_name = ""
oil_kind = ""
car_owner = ""
buy_gubun = "����"
owner_emp_no = ""
emp_name = ""
emp_grade = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "���� ���"
if u_type = "U" then

    Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
    Set rs_car = DbConn.Execute(SQL)
    if not rs_car.eof then
        car_name       = rs_car("car_name")
        car_year       = rs_car("car_year")
        car_reg_date   = rs_car("car_reg_date")
        car_use_dept   = rs_car("car_use_dept")
        car_company    = rs_car("car_company")
        car_use        = rs_car("car_use")
        owner_emp_name = rs_car("owner_emp_name")
        owner_emp_no   = rs_car("owner_emp_no")
        oil_kind       = rs_car("oil_kind")
        car_owner      = rs_car("car_owner")
        buy_gubun      = rs_car("buy_gubun")
    else
        car_name       = ""
        car_year       = ""
        car_reg_date   = ""
        car_use_dept   = ""
        car_company    = ""
        car_use        = ""
        owner_emp_name = ""
        owner_emp_no   = ""
        oil_kind       = ""
        car_owner      = ""
        buy_gubun      = ""
    end if
    rs_car.close()
'
'	work_date = request("work_date")
'	mg_ce_id = request("mg_ce_id")
'
'	sql = "select * from overtime where work_date = '" + work_date + "' and mg_ce_id = '" + mg_ce_id + "'"
'	set rs = dbconn.execute(sql)
'
'	sql="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
'	set rs_memb=dbconn.execute(sql)
'
'	if	rs_memb.eof or rs_memb.bof then
'		mg_ce = "ERROR"
'	  else
'		mg_ce = rs_memb("user_name")
'	end if
'	rs_memb.close()						
'
'	if isnull(rs("acpt_no")) then
'		acpt_no = 0
'	  else
'		acpt_no = rs("acpt_no")
'	end if
'	mg_ce_id = rs("mg_ce_id")
'	company = rs("company")
'	dept = rs("dept")
'	work_item = rs("work_item")
'	from_hh = mid(rs("from_time"),1,2)
'	from_mm = mid(rs("from_time"),3,2)
'	to_hh = mid(rs("to_time"),1,2)
'	to_mm = mid(rs("to_time"),3,2)
'	work_gubun = rs("work_gubun") + "/" + cstr(rs("overtime_amt"))
'	overtime_amt = int(rs("overtime_amt"))
'	work_memo = rs("work_memo")
'	reg_id = rs("reg_id")
'	reg_date = rs("reg_date")
'	mod_id = rs("mod_id")
'	mod_date = rs("mod_date")
'	rs.close()
'
	title_line = "���� ����"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {  $( "#datepicker" ).datepicker();
							$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker" ).datepicker("setDate", "<%=car_reg_date%>" );
			});	  
			$(function() {  $( "#datepicker1" ).datepicker();
							$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker1" ).datepicker("setDate", "<%=last_check_date%>" );
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
				if(document.frm.car_no.value =="" ) {
					alert('������ȣ�� �Է��ϼ���');
					frm.car_no.focus();
					return false;}
				if(document.frm.car_name.value =="") {
					alert('������ �Է��ϼ���');
					frm.car_name.focus();
					return false;}
				if(document.frm.oil_kind.value =="") {
					alert('������ �����ϼ���');
					frm.oil_kind.focus();
					return false;}			
				if(document.frm.car_owner.value =="") {
					alert('�����ڸ� �����ϼ���');
					frm.car_owner.focus();
					return false;}			
				if(document.frm.car_reg_date.value =="") {
					alert('����������� �Է��ϼ���');
					frm.car_reg_date.focus();
					return false;}			
				if(document.frm.owner_emp_no.value =="" ) {
					alert('�����˻��� �ϼ���');
					frm.emp_name.focus();
					return false;}
			
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
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
				<h3 class="tit"><%=title_line%></h3>
				<form action="car_info_add_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">������ȣ</th>
								<td class="left">
                                <input name="car_no" type="text" value="<%=car_no%>" style="width:150px" onKeyUp="checklength(this,20)"></td>
								<th>����</th>
								<td class="left">
                                <input name="car_name" type="text" value="<%=car_name%>" style="width:150px" onKeyUp="checklength(this,30)"></td>
							</tr>
							<tr>
								<th class="first">����</th>
								<td class="left">
                                <select name="oil_kind" id="oil_kind" style="width:150px">
								  <option value="">����</option>
								  <option value="�޹���" <%If oil_kind = "�ֹ���" then %>selected<% end if %>>�ֹ���</option>
								  <option value="����" <%If oil_kind = "����" then %>selected<% end if %>>����</option>
								  <option value="����" <%If oil_kind = "����" then %>selected<% end if %>>����</option>
							    </select>
                                </td>
								<th>����</th>
                                <td class="left"><select name="car_owner" id="car_owner" style="width:150px">
								  <option value="">����</option>
								  <option value="ȸ��" <%If car_owner = "ȸ��" then %>selected<% end if %>>ȸ��</option>
								  <option value="����" <%If car_owner = "����" then %>selected<% end if %>>����</option>
							    </select></td>
							</tr>
							<tr>
								<th class="first">���ű���</th>
								<td class="left">
                                <input type="radio" name="buy_gubun" value="����" <% if buy_gubun = "����" then %>checked<% end if %> style="width:40px" id="Radio1">����
                                <input type="radio" name="buy_gubun" value="����" <% if buy_gubun = "����" then %>checked<% end if %> style="width:40px" id="Radio2">����
                                <input type="radio" name="buy_gubun" value="��Ʈ" <% if buy_gubun = "��Ʈ" then %>checked<% end if %> style="width:40px" id="Radio2">��Ʈ
                                </td>
								<th>���������</th>
								<td class="left"><input name="car_reg_date" type="text" value="<%=car_reg_date%>" style="width:70px" id="datepicker"></td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td colspan="3" class="left">
                                <input name="emp_name" type="text" id="emp_name" style="width:80px" value="<%=emp_name%>" readonly="true">
                                <input name="emp_grade" type="text" id="emp_grade" style="width:80px" value="<%=emp_grade%>" readonly="true">
                                <input name="owner_emp_no" type="text" id="owner_emp_no" style="width:80px" value="<%=owner_emp_no%>" readonly="true">
							    <a href="#" class="btnType03" onClick="pop_Window('emp_search_pop.asp?gubun=<%=1%>','emp_search_pop','scrollbars=yes,width=600,height=400')">�����˻�</a></td>
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
			</form>
		</div>				
	</body>
</html>

