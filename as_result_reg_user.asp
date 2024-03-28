<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)

acpt_no = request("acpt_no")
be_pg = request("be_pg")

page = request("page")
from_date = request("from_date")
to_date = request("to_date")
date_sw = request("date_sw")
process_sw = request("process_sw")
field_check = request("field_check")
field_view = request("field_view")
view_sort = request("view_sort")
page_cnt = request("page_cnt")
condi_com = request("company")
view_c = request("view_c")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_into = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect


'if rs("as_process") = "�԰�" then
	sql = "select max(in_seq) as max_seq from as_into where acpt_no = "&int(acpt_no)
	set rs_into = dbconn.execute(sql)
	max_seq = rs_into("max_seq")

	if isnull(max_seq) then	
		in_process = "����"
		in_place = "����"
	  else
		in_seq = max_seq
		sql = "select in_process,in_place from as_into where acpt_no = "&int(acpt_no)&" and in_seq = "&int(in_seq)
		Set rs_into = DbConn.Execute(SQL)
		if rs_into.eof then
			in_process = "����"
			in_place = "����"
		  else
			in_process = rs_into("in_process")
			in_place = rs_into("in_place")
		end if
	end if
'end if

Sql = "select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs = DbConn.Execute(SQL)

acpt_date = mid(cstr(rs("acpt_date")),1,10)
acpt_hh = int(datepart("h",rs("acpt_date")))
acpt_mm = int(datepart("n",rs("acpt_date")))
acpt_ss = datepart("s",rs("acpt_date"))

if acpt_hh < 10 then
	acpt_hh = "0" + cstr(acpt_hh)
end if

if acpt_mm < 10 then
	acpt_mm = "0" + cstr(acpt_mm)
end if

if arrival_time = "0000" or arrival_time = null or arrival_time = "" then
	arrival_date = curr_date
	arrival_time = "0000"
end if

if isnull(rs("dev_inst_cnt")) or rs("dev_inst_cnt") = "" then
	dev_inst_cnt = "1"
  else
  	dev_inst_cnt = rs("dev_inst_cnt")
end if

as_type = rs("as_type")
if rs("sms") = "Y" then
	sms_view = "�߼�"
  else
  	sms_view = "�߼۾���"
end if
new_sms = "N"

title_line = "A/S ��� ���"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
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
											$( "#datepicker" ).datepicker("setDate", "<%=rs("request_date")%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
											$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker1" ).datepicker("setDate", "<%=rs("visit_date")%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
											$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker2" ).datepicker("setDate", "<%=rs("in_date")%>" );
			});	  
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.as_process.value == "�԰�" && document.frm.as_process_old.value == "�԰�") {
					alert('�԰� ���¿����� ������ �Ұ� �մϴ� !!!');
					frm.as_process.focus();
					return false;}
				if(document.frm.c_grade.value >"4") {
					alert('���� �Ǵ� ��� ������ �����ϴ� !!!');
					frm.addr.focus();
					return false;}
				if(document.frm.acpt_user.value == "") {
					alert('����ڸ� �Է��ϼ��� !!!');
					frm.acpt_user.focus();
					return false;}
				if(document.frm.addr.value =="") {
					alert('������ �ּҸ� �Է��ϼ���');
					frm.addr.focus();
					return false;}
				if(document.frm.as_memo.value =="") {
					alert('��ֳ����� �Է��ϼ���');
					frm.as_memo.focus();
					return false;}			
				if(document.frm.request_date.value =="") {
					alert('��û���� �Է��ϼ���');
					frm.request_date.focus();
					return false;}
				if(document.frm.request_date.value < document.frm.acpt_date.value) {
					alert('��û���� �����Ϻ��� �����ϴ�');
					frm.request_date.focus();
					return false;}
				if(document.frm.request_hh.value >"23"||document.frm.request_hh.value <"00") {
					alert('��û�ð��� �߸��Ǿ����ϴ�');
					frm.request_hh.focus();
					return false;}
				if(document.frm.request_mm.value >"59"||document.frm.request_mm.value <"00") {
					alert('��û���� �߸��Ǿ����ϴ�');
					frm.request_mm.focus();
					return false;}
				if(document.frm.request_date.value == document.frm.acpt_date.value) {
					if(document.frm.request_hh.value < document.frm.acpt_hh.value) {
						alert('��û�ð��� �����ð� ���� �����ϴ�');
						frm.request_hh.focus();
						return false;}}
				if(document.frm.request_date.value == document.frm.acpt_date.value) {
					if(document.frm.request_hh.value == document.frm.acpt_hh.value) {
						if(document.frm.request_mm.value <= document.frm.acpt_mm.value) {
							alert('��û���� ������ ���� �����ϴ�');
							frm.request_mm.focus();
							return false;}}}
			
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.visit_date.value =="") {
						alert('�Ϸ����� �Է��ϼ���');
						frm.visit_date.focus();
						return false;}
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.visit_date.value < document.frm.acpt_date.value) {
						alert('�Ϸ����� �����Ϻ��� �����ϴ�');
						frm.visit_date.focus();
						return false;}
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.visit_date.value > document.frm.curr_date.value) {
						alert('�Ϸ����� �����Ϻ��� �����ϴ�');
						frm.visit_date.focus();
						return false;}
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.visit_hh.value >"23"||document.frm.visit_hh.value <"00") {
						alert('�Ϸ�ð��� �߸��Ǿ����ϴ�');
						frm.visit_hh.focus();
						return false;}
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.visit_mm.value >"59"||document.frm.visit_mm.value <"00") {
						alert('�Ϸ���� �߸��Ǿ����ϴ�');
						frm.visit_mm.focus();
						return false;}
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.visit_date.value == document.frm.acpt_date.value) {
						if(document.frm.visit_hh.value < document.frm.acpt_hh.value) {
							alert('�Ϸ�ð��� �����ð� ���� �����ϴ�');
							frm.visit_hh.focus();
							return false;}}
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.visit_date.value == document.frm.acpt_date.value) {
						if(document.frm.visit_hh.value == document.frm.acpt_hh.value) {
							if(document.frm.visit_mm.value <= document.frm.acpt_mm.value) {
								alert('�Ϸ���� ������ ���� �����ϴ�');
								frm.visit_mm.focus();
								return false;}}}
			
				if(document.frm.as_process_old.value =="�԰�" || document.frm.as_process_old.value =="��ü�԰�") 
					if(document.frm.as_process.value =="����" || document.frm.as_process.value =="����") {
						document.frm.as_process.value = document.frm.as_process_old.value
						alert('�԰� ������ ����� ���� �Ұ�');
						frm.as_process.focus();
						return false;}
				if(document.frm.as_process_old.value =="�԰�" || document.frm.as_process_old.value =="��ü�԰�") 
					if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="���") {
						if(document.frm.in_process.value !="�����Ϸ�") {
							if(document.frm.in_process.value !="�԰����") {
								alert('�����Ϸ� �Ǵ� �԰���� ���� �ʾ� �Ϸ� �Ǵ� ��� ��� �Ҽ� �����ϴ�');
								frm.as_process.focus();
								return false;}}}
				if(document.frm.as_process.value =="�԰�" || document.frm.as_process.value =="��ü�԰�" || document.frm.as_process.value =="��ü") 
					if(document.frm.as_type.value !="�湮ó��") {
						alert('�԰�,��ü �� ��ü�԰�� �ݵ�� �湮ó���̾�� ��');
						frm.as_type.focus();
					return false;}
				if(document.frm.as_process.value =="�԰�" || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.into_reason.value =="") {
						alert('�԰� �� ���� ������ �Է��ϼ���');
						frm.into_reason.focus();
					return false;}
				if(document.frm.as_process.value =="�԰�" || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.in_date.value < document.frm.acpt_date.value) {
						alert('�԰����ڰ� �������ں��� �۽��ϴ�');
						frm.in_date.focus();
					return false;}
				if(document.frm.as_process.value =="�԰�" || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.in_date.value =="") {
						alert('�԰����ڸ� �Է��ϼ���');
						frm.in_date.focus();
					return false;}
				if(document.frm.as_process.value =="�԰�") 
					if(document.frm.in_date.value > document.frm.curr_date.value) {
						alert('�԰����� �����Ϻ��� �����ϴ�');
						frm.in_date.focus();
						return false;}
				if(document.frm.as_process.value =="�԰�" || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.in_place.value =="����") {
						alert('�԰�ó�� �Է��ϼ���');
						frm.in_place.focus();
					return false;}
				if(document.frm.as_process.value =="�԰�" || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.in_replace.value =="") {
						alert('��ü���θ� �����Ͽ��� �մϴ�');
						frm.in_replace.focus();
					return false;}
			
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
					if(document.frm.as_history.value =="") {
						alert('ó�� ������ �Է��ϼ���');
						frm.as_history.focus();
					return false;}
				if(document.frm.as_process.value =="�Ϸ�") 
					if(document.frm.as_type.value =="�űԼ�ġ" || document.frm.as_type.value =="�űԼ�ġ����" || document.frm.as_type.value =="������ġ" || document.frm.as_type.value =="������ġ����" || document.frm.as_type.value =="������" || document.frm.as_type.value =="����������" || document.frm.as_type.value =="���ȸ��" || document.frm.as_type.value =="��������") {
						if(document.frm.dev_inst_cnt.value < 0 || document.frm.dev_inst_cnt.value > 999 || document.frm.dev_inst_cnt.value == "") {
							alert('��ġ����� 999���� ũ�ų� �߸��Ǿ����ϴ�');
							frm.dev_inst_cnt.focus();
					return false;}}
				if(document.frm.as_process.value =="�Ϸ�") 
					if(document.frm.as_type.value =="�űԼ�ġ" || document.frm.as_type.value =="�űԼ�ġ����" || document.frm.as_type.value =="������ġ" || document.frm.as_type.value =="������ġ����" || document.frm.as_type.value =="������" || document.frm.as_type.value =="����������" || document.frm.as_type.value =="���ȸ��" || document.frm.as_type.value =="��������") {
						if(document.frm.ran_cnt.value < 0 || document.frm.ran_cnt.value > 999 || document.frm.ran_cnt.value == "") {
							alert('�������� 999���� ũ�ų� �߸��Ǿ����ϴ�');
							frm.ran_cnt.focus();
					return false;}}
				if(document.frm.as_process.value =="�Ϸ�") 
					if(document.frm.as_type.value =="�űԼ�ġ" || document.frm.as_type.value =="�űԼ�ġ����" || document.frm.as_type.value =="������ġ" || document.frm.as_type.value =="������ġ����" || document.frm.as_type.value =="������" || document.frm.as_type.value =="����������") {
						if(document.frm.work_man_cnt.value < 1 || document.frm.work_man_cnt.value > 30 || document.frm.work_man_cnt.value == "") {
							alert('�۾� �ο��� 30���� ũ�ų� �߸��Ǿ����ϴ�');
							frm.work_man_cnt.focus();
					return false;}}
				if(document.frm.as_process.value =="�Ϸ�") 
					if(document.frm.as_type.value =="�űԼ�ġ" || document.frm.as_type.value =="�űԼ�ġ����" || document.frm.as_type.value =="������ġ" || document.frm.as_type.value =="������ġ����" || document.frm.as_type.value =="������" || document.frm.as_type.value =="����������") {
						if(document.frm.alba_cnt.value < 0 || document.frm.alba_cnt.value > 30 || document.frm.alba_cnt.value == "") {
							alert('�˹� �ο��� 30���� ũ�ų� �߸��Ǿ����ϴ�');
							frm.alba_cnt.focus();
					return false;}}
			
				j=0;
				   for(i=0;i<document.frm.err01.length;i++){  
					if (document.frm.err01[i].checked==true){   
					 j++;
					}
				   }
				k=0;
				   for(i=0;i<document.frm.err02.length;i++){  
					if (document.frm.err02[i].checked==true){   
					 k++;
					}
				   }
			
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
				 if(document.frm.as_type.value =="����ó��" || document.frm.as_type.value =="�湮ó��" || document.frm.as_type.value =="��Ÿ") 
					if(document.frm.as_device.value =="����ũž" || document.frm.as_device.value =="��Ʈ��" || document.frm.as_device.value =="DTO" || document.frm.as_device.value =="DTS") 
						if(j == 0 && k == 0) {
							alert('���ó���� CHECK �ϼ���');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err03.length;i++){  
					if (document.frm.err03[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
				 if(document.frm.as_type.value =="����ó��" || document.frm.as_type.value =="�湮ó��" || document.frm.as_type.value =="��Ÿ") 
					if(document.frm.as_device.value =="�����") 
						if(j == 0) {
							alert('����� ���ó���� CHECK �ϼ���');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err04.length;i++){  
					if (document.frm.err04[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
				 if(document.frm.as_type.value =="����ó��" || document.frm.as_type.value =="�湮ó��" || document.frm.as_type.value =="��Ÿ") 
					if(document.frm.as_device.value =="������" || document.frm.as_device.value =="���ɳ�" || document.frm.as_device.value =="�÷���") 
						if(j == 0) {
							alert('���ó���� CHECK �ϼ���');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err05.length;i++){  
					if (document.frm.err05[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
				 if(document.frm.as_type.value =="����ó��" || document.frm.as_type.value =="�湮ó��" || document.frm.as_type.value =="��Ÿ") 
					if(document.frm.as_device.value =="������" || document.frm.as_device.value =="AP" || document.frm.as_device.value =="���" || document.frm.as_device.value =="�����" || document.frm.as_device.value =="TA" || document.frm.as_device.value =="��Ʈ�����" || document.frm.as_device.value =="ȸ��") 
						if(j == 0) {
							alert('������ ���ó���� CHECK �ϼ���');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err06.length;i++){  
					if (document.frm.err06[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
				 if(document.frm.as_type.value =="����ó��" || document.frm.as_type.value =="�湮ó��" || document.frm.as_type.value =="��Ÿ") 
					if(document.frm.as_device.value =="����" || document.frm.as_device.value =="��ũ�����̼�") 
						if(j == 0) {
							alert('���ó���� CHECK �ϼ���');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err07.length;i++){  
					if (document.frm.err07[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
				 if(document.frm.as_type.value =="����ó��" || document.frm.as_type.value =="�湮ó��" || document.frm.as_type.value =="��Ÿ") 
					if(document.frm.as_device.value =="�ƴ���") 
						if(j == 0) {
							alert('���ó���� CHECK �ϼ���');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err09.length;i++){  
					if (document.frm.err09[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="�Ϸ�" || document.frm.as_process.value =="��ü" || document.frm.as_process.value =="���"  || document.frm.as_process.value =="��ü�԰�") 
				 if(document.frm.as_type.value =="����ó��" || document.frm.as_type.value =="�湮ó��" || document.frm.as_type.value =="��Ÿ") 
					if(document.frm.as_device.value =="��Ÿ") 
						if(j == 0) {
							alert('��Ÿ ������ CHECK �ϼ���');
							frm.as_history.focus();
						return false;}
					
				if(document.frm.as_process.value =="�Ϸ�"){
					if (document.frm.as_type.value == '�űԼ�ġ' || document.frm.as_type.value == '�űԼ�ġ����' || document.frm.as_type.value == '������ġ' || document.frm.as_type.value == '������ġ����' || document.frm.as_type.value == '������' || document.frm.as_type.value == '����������' || document.frm.as_type.value == '���ȸ��' || document.frm.as_type.value == '��������') {
						if(document.frm.att_file1.value =="" && document.frm.att_file2.value =="" && document.frm.att_file3.value =="" && document.frm.att_file4.value =="" && document.frm.att_file5.value =="") {
							alert('���� ÷�ΰ� ���� �ʾҽ��ϴ�');
							frm.att_file1.focus();
							return false;}}}
					
				if(document.frm.as_process.value =="�Ϸ�"){
					if (document.frm.as_type.value == '�űԼ�ġ' || document.frm.as_type.value == '�űԼ�ġ����' || document.frm.as_type.value == '������ġ' || document.frm.as_type.value == '������ġ����' || document.frm.as_type.value == '������' || document.frm.as_type.value == '����������' || document.frm.as_type.value == '���ȸ��' || document.frm.as_type.value == '��������') {
					{
					b=confirm('�۾��ο��� ' + document.frm.work_man_cnt.value +'�� �½��ϱ�?')
					if (b==false) {
						return false;
						}
					}
				}}
				{
				a=confirm('�����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function ce_mod_view() 
			{
				if (document.frm.ce_mod_ck.checked == true) {
					document.getElementById('s_ce').style.display = ''; 
					document.getElementById('s_ce_id').style.display = ''; 
					document.getElementById('ce_mod').style.display = ''; }
				if (document.frm.ce_mod_ck.checked == false) {
					document.getElementById('ce_mod').style.display = 'none'; 
					document.getElementById('s_ce').style.display = 'none'; 
					document.getElementById('s_ce_id').style.display = 'none'; }
			}
			function inview() {
			var c = document.frm.as_process.options[document.frm.as_process.selectedIndex].value;
				if (c == '�԰�' || c == '��ü�԰�') 
				{
					document.getElementById('in_menu').style.display = '';
				}
			}
			function menu1() {

				var c = document.frm.as_process.options[document.frm.as_process.selectedIndex].value;
				var d = document.frm.as_device.options[document.frm.as_device.selectedIndex].value;
				var e = document.frm.as_type.options[document.frm.as_type.selectedIndex].value;
				var f = document.frm.company.value;
					 {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = 'none';
						document.getElementById('end_keyin2').style.display = 'none';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					}
					if (c == '�԰�') 
					{
						document.getElementById('in_menu').style.display = '';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = 'none';
						document.getElementById('end_keyin2').style.display = 'none';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					}
					if (c == '�Ϸ�' || c == '��ü' || c == '���') 
					  if (e == '����ó��' || e == '�湮ó��' || e == '��Ÿ') {
						if (d == '����ũž' || d == '��Ʈ��' || d == 'DTO' || d == 'DTS') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = '';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '�Ϸ�') 
						if (e == '�űԼ�ġ' || e == '�űԼ�ġ����' || e == '������ġ' || e == '������ġ����' || e == '������' || e == '����������' || e == '���ȸ��' || e == '��������') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = '';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = '';		
					}
					if (c == '���') 
						if (e == '�űԼ�ġ' || e == '�űԼ�ġ����' || e == '������ġ' || e == '������ġ����' || e == '������' || e == '����������' || e == '���ȸ��' || e == '��������') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					}
					if (c == '�Ϸ�' || c == '��ü' || c == '���') 
					  if (e == '����ó��' || e == '�湮ó��' || e == '��Ÿ') {
						if (d == '�����') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = '';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '�Ϸ�' || c == '��ü' || c == '���') 
					  if (e == '����ó��' || e == '�湮ó��' || e == '��Ÿ') {
						if (d == '������' || d == '���ɳ�' || d == '�÷���') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = '';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '�Ϸ�' || c == '��ü' || c == '���') 
					  if (e == '����ó��' || e == '�湮ó��' || e == '��Ÿ') {
						if (d == '������' || d == 'AP' || d == '���' || d == '�����' || d == 'TA' || d == '��Ʈ�����' || d == 'ȸ��') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = '';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '�Ϸ�' || c == '��ü' || c == '���') 
					  if (e == '����ó��' || e == '�湮ó��' || e == '��Ÿ') {
						if (d == '����' || d == '��ũ�����̼�') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = '';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '�Ϸ�' || c == '��ü' || c == '���') 
					  if (e == '����ó��' || e == '�湮ó��' || e == '��Ÿ') {
						if (d == '�ƴ���') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = '';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '�Ϸ�' || c == '��ü' || c == '���')
					  if (e == '����ó��' || e == '�湮ó��' || e == '��Ÿ') {
						if (d == '��Ÿ') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = '';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}				
				}
		</script>

	</head>
	<body onLoad="inview()">
		<div id="wrap">			
			<!--#include virtual = "/include/user_header.asp" -->
			<!--#include virtual = "/include/as_sub_menu_user.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_result_reg_user_ok.asp" method="post" enctype="multipart/form-data" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">������ȣ</th>
								<td class="left"><%=rs("acpt_no")%>
                                <input name="acpt_no" type="hidden" id="acpt_no" value="<%=rs("acpt_no")%>">
                				<input name="c_grade" type="hidden" id="c_grade" value="<%=c_grade%>">
                                </td>
								<th>������</th>
								<td class="left"><%=rs("acpt_date")%>
								<input name="acpt_date" type="hidden" id="acpt_date" value="<%=acpt_date%>">
                				<input name="acpt_hh" type="hidden" id="acpt_hh" value="<%=acpt_hh%>">
                				<input name="acpt_mm" type="hidden" id="acpt_mm" value="<%=acpt_mm%>">
                                </td>
								<th>������</th>
								<td class="left"><%=rs("acpt_man")%>
                				<input name="curr_date" type="hidden" id="curr_date" value="<%=curr_date%>">
            					</td>
								<th>ȸ��</th>
								<td class="left"><%=rs("company")%>
                				<input name="company" type="hidden" id="company" value="<%=rs("company")%>">
                                </td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td class="left"><%=rs("dept")%><input name="dept" type="hidden" id="dept" value="<%=rs("dept")%>"></td>
								<th>��ȭ��ȣ1</th>
								<td class="left"><%=rs("tel_ddd")%>-<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
								<th>�����</th>
								<td class="left">
                                <input name="acpt_user" type="text" size="10" onKeyUp="checklength(this,20)" value="<%=rs("acpt_user")%>">
								  &nbsp;<strong>����</strong>
                                <input name="user_grade" type="text" size="8" onKeyUp="checklength(this,20)" value="<%=rs("user_grade")%>"></td>
								<th>��ȭ��ȣ2</th>
								<td class="left">
								<select name="hp_ddd" id="hp_ddd">
									<option>����</option>
									<option value="02" <%If rs("hp_ddd") = "02" then %>selected<% end if %>>02</option>
									<option value="010" <%If rs("hp_ddd") = "010" then %>selected<% end if %>>010</option>
				  					<option value="011" <%If rs("hp_ddd") = "011" then %>selected<% end if %>>011</option>
				  					<option value="016" <%If rs("hp_ddd") = "016" then %>selected<% end if %>>016</option>
				  					<option value="017" <%If rs("hp_ddd") = "017" then %>selected<% end if %>>017</option>
				  					<option value="018" <%If rs("hp_ddd") = "018" then %>selected<% end if %>>018</option>
				  					<option value="019" <%If rs("hp_ddd") = "019" then %>selected<% end if %>>019</option>
								</select>-              	
								<input name="hp_no1" type="text" id="hp_no1" size="4" maxlength="4" value="<%=rs("hp_no1")%>">-
                            	<input name="hp_no2" type="text" id="hp_no2" size="4" maxlength="4" value="<%=rs("hp_no2")%>">
                              </td>
							</tr>
							<tr>
								<th class="first">�ּ�</th>
								<td colspan="5" class="left"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>
                                <input name="sido" type="hidden" id="sido" value="<%=rs("sido")%>">
                                <input name="gugun" type="hidden" id="gugun" value="<%=rs("gugun")%>">
                                <input name="dong" type="hidden" id="dong2" value="<%=rs("dong")%>">
              					<input name="addr" type="text" id="addr" style="width:250px" onKeyUp="checklength(this,50)" value="<%=rs("addr")%>">
              					<input name="view_ok" type="hidden" id="view_ok" value="">
                                </td>
								<th>��������</th>
								<td class="left"><%=sms_view%></td>
							</tr>
							<tr>
								<th class="first">����CE</th>
								<td class="left"><%=rs("mg_ce")%>&nbsp;(&nbsp;<%=rs("mg_ce_id")%>&nbsp;)
                                <input name="mg_ce_id" type="hidden" id="mg_ce_id2" value="<%=rs("mg_ce_id")%>">
                				<input name="mg_ce" type="hidden" id="mg_ce" value="<%=rs("mg_ce")%>">                           					
                                </td>
								<th>����CE</th>
								<td class="left" colspan="3"><strong>������ ���ϸ� �����ϼ���</strong>
								<input name="ce_mod_ck" type="checkbox" id="ce_mod_ck" value="1"  onClick="ce_mod_view()">
                				<input name="s_ce" id="s_ce" type="text" value="<%=user_name%>" size="10" readonly="true" style="display:none">
                				<input name="s_ce_id" id="s_ce_id" type="text" value="<%=user_id%>" size="10" readonly="true" style="display:none">
                				<input name="s_reside_place" type="hidden" id="s_reside_place">
                				<input name="s_team" type="hidden" id="s_team">
                                <a href="#" class="btnType03" onClick="pop_Window('ce_select.asp?gubun=<%="����"%>&mg_group=<%=mg_group%>','ceselect','scrollbars=yes,width=500,height=400')" id="ce_mod" style="display:none">CE����</a>
                                </td>
								<th>������߼�</th>
								<td class="left">
                                <input type="radio" name="new_sms" value="Y" <% if new_sms = "Y" then %>checked<% end if %>>�߼� 
              					<input name="new_sms" type="radio" value="N" <% if new_sms = "N" then %>checked<% end if %>>�߼۾���
                                </td>
							</tr>
							<tr>
								<th class="first">��ֳ���</th>
								<td class="left" colspan="7">
                                <textarea name="as_memo" rows="5" id="textarea"><%=rs("as_memo")%></textarea>
                                </td>
							</tr>
							<tr>
								<th class="first">��û��/�ð�</th>
								<td class="left">
                                <input name="request_date" type="text" size="10" readonly="true" id="datepicker" style="width:70px;">&nbsp;
                                <input name="request_hh" type="text" id="request_hh" value="<%=mid(rs("request_time"),1,2)%>" size="2" maxlength="2">
                                <strong>��</strong>
                                <input name="request_mm" type="text" id="request_mm" value="<%=mid(rs("request_time"),3,2)%>" size="2" maxlength="2"><strong>��</strong>
							  	</td>
								<th>�Ϸ���/�ð�</th>
								<td class="left">
                                <input name="visit_date" type="text" size="10" readonly="true" id="datepicker1" style="width:70px;">&nbsp;
                                <input name="visit_hh" type="text" id="visit_hh" value="<%=mid(rs("visit_time"),1,2)%>" size="2" maxlength="2">
                                <strong>��</strong>
                                <input name="visit_mm" type="text" id="visit_mm" value="<%=mid(rs("visit_time"),3,2)%>" size="2" maxlength="2"><strong>��</strong>
                                </td>
							  <th>ó������</th>
								<td class="left">
								<% if (as_type = "�űԼ�ġ" or as_type = "�űԼ�ġ����" or as_type = "������ġ" or as_type = "������ġ����" or as_type = "������" or as_type = "����������" or as_type = "���ȸ��" or as_type = "��������") then %>
                                <select name="as_type" id="as_type" style="width:150px" onChange="menu1()">
                                  <option value="<%=as_type%>" <%If as_type = as_type then %>selected<% end if %>><%=as_type%></option>
                                </select>
                                <%   else %>
                                <select name="as_type" id="as_type" style="width:150px" onChange="menu1()">
								  <option value="�湮ó��" <%If as_type = "�湮ó��" then %>selected<% end if %>>�湮ó��</option>
								  <option value="����ó��" <%If as_type = "����ó��" then %>selected<% end if %>>����ó��</option>
								  <option value="�űԼ�ġ" <%If as_type = "�űԼ�ġ" then %>selected<% end if %>>�űԼ�ġ</option>
								  <option value="�űԼ�ġ����" <%If as_type = "�űԼ�ġ����" then %>selected<% end if %>>�űԼ�ġ����</option>
								  <option value="������ġ" <%If as_type = "������ġ" then %>selected<% end if %>>������ġ</option>
								  <option value="������ġ����" <%If as_type = "������ġ����" then %>selected<% end if %>>������ġ����</option>
								  <option value="������" <%If as_type = "������" then %>selected<% end if %>>������</option>
								  <option value="����������" <%If as_type = "����������" then %>selected<% end if %>>����������</option>
								  <option value="���ȸ��" <%If as_type = "���ȸ��" then %>selected<% end if %>>���ȸ��</option>
								  <option value="��������" <%If as_type = "��������" then %>selected<% end if %>>��������</option>
								  <option value="��Ÿ" <%If as_type = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
							    </select>
			 					<% end if %>
             					<input name="as_type_old" type="hidden" id="as_type_old" value="<%=as_type%>">
                                </td>
								<th>ó����Ȳ</th>
								<td class="left">
                                <select name="as_process" style="width:150px" onChange="menu1()">
                                  <option value="����"  <%If rs("as_process") = "����" then %>selected<% end if %>>����</option>
                                  <option value="�Ϸ�"  <%If rs("as_process") = "�Ϸ�" then %>selected<% end if %>>�Ϸ�</option>
                                  <option value="�԰�"  <%If rs("as_process") = "�԰�" then %>selected<% end if %>>�԰�</option>
                                  <option value="����"  <%If rs("as_process") = "����" then %>selected<% end if %>>����</option>
                                  <option value="���"  <%If rs("as_process") = "���" then %>selected<% end if %>>���</option>
                                </select>                
                                <input name="as_process_old" type="hidden" id="as_process_old" value="<%=rs("as_process")%>">
                                </td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td class="left">
                            <%
								Sql="select * from etc_code where etc_type = '31' order by etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
							%>
								<select name="as_device" id="select" style="width:150px" onChange="menu1()">
                			<% 
								do until rs_etc.eof 
			  				%>
                					<option value='<%=rs_etc("etc_name")%>' <%If rs("as_device") = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			<%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							%>
            					</select>
            					</td>
								<th>������</th>
								<td class="left">
                            <%
								Sql="select * from etc_code where etc_type = '21' order by etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
							%>
              					<select name="maker" id="maker" style="width:150px">
                			<% 
								do until rs_etc.eof 
			  				%>
                					<option value='<%=rs_etc("etc_name")%>' <%If rs("maker") = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			<%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							%>
            					</select>
            					</td>
								<th>�𵨸�</th>
								<td class="left">
                                <input name="model_no" type="text" id="model_no" style="width:150px" onKeyUp="checklength(this,20)" value="<%=rs("model_no")%>">
                            </td>
								<th>�ø����ȣ</th>
								<td class="left"><input name="serial_no" type="text" id="serial_no" style="width:150px"  onKeyUp="checklength(this,20)" value="<%=rs("serial_no")%>"></td>
							</tr>
							<tr>
								<th class="first">�ڻ��ȣ</th>
								<td class="left"><input name="asets_no" type="text" id="asets_no" style="width:150px" onKeyUp="checklength(this,20)" value="<%=rs("asets_no")%>"></td>
								<th>����/�԰����</th>
								<td class="left" colspan="5"><textarea name="into_reason"><%=rs("into_reason")%></textarea></td>
							</tr>
							<tr style="display:none" id="in_menu">
								<th class="first" style="background:#FCF">�԰�����</th>
								<td class="left"><input name="in_date" type="text" id="datepicker2" size="10" readonly="true"></td>
								<td style="background:#FCF">�԰�ó</td>
								<td class="left">
							  <% if rs("as_process") = "�԰�" then	%>
                              	<%=in_place%>
                              <%	  else	%>
                              	<select name="in_place" class="style12" id="select2">
                                	<option value="����">����</option>
                                	<option value="��ü�԰�">��ü�԰�</option>
                                	<option value="�����԰�">�����԰�</option>
                                	<option value="Repair Shop">Repair Shop</option>
                              	</select>
                              <% end if %>
                                </td>
								<td style="background:#FCF">��ü</td>
								<td class="left">
                				<select name="in_replace" id="in_replace">
                					<option></option>
                					<option value="����" <%If rs("in_replace") = "����" then %>selected<% end if %>>����</option>
                					<option value="��ü" <%If rs("in_replace") = "��ü" then %>selected<% end if %>>��ü</option>
              					</select>
            					</td>
								<td style="background:#FCF">�԰�����</td>
								<td class="left"><%=in_process%>
                				<input name="in_process" type="hidden" id="in_process" value="<%= in_process%>">
                                </td>
							</tr>
							<tr style="display:none" id="end_keyin1">
								<th class="first" style="background:#FFC">����ǰ</th>
								<td class="left" colspan="7"><input name="as_parts" type="text" id="as_parts" onKeyUp="checklength(this,50)" value="<%=rs("as_parts")%>" size="50"></td>
							</tr>
							<tr style="display:none" id="end_keyin2">
								<th class="first" style="background:#FFC">ó������</th>
								<td class="left" colspan="7"><textarea name="as_history" rows="2" id="textarea"></textarea></td>
							</tr>
							<tr id="inst_menu" style="display:none">
								<th class="first" colspan="2" style="background:#E1FFE1">��ġ,����,����,ȸ��,�������� ����</th>
								<td class="left" colspan="6" bgcolor="#E1FFE1">
								��ġ���
                                <input name="dev_inst_cnt" type="text" id="dev_inst_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);"  maxlength="3" value="<%=dev_inst_cnt%>">
                                ��&nbsp; 
                                ������
                                <input name="ran_cnt" type="text" id="ran_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);" value="0" maxlength="3">��&nbsp;
                                �۾��η�
                                <input name="work_man_cnt" type="text" id="work_man_cnt" style="width:30px;text-align:right" value="1" maxlength="2" readonly="true">��&nbsp;
                                �˹��ο�
                                <input name="alba_cnt" type="text" id="alba_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);" value="0" maxlength="2">��&nbsp;
								<a href="#" id="work_ce" class="btnType03" onClick="pop_Window('work_ce_add.asp?acpt_no=<%=rs("acpt_no")%>','work_ce_add_pop','scrollbars=yes,width=700,height=500')">2���̻��۾��ηµ��</a>
                                <br><strong>�۾��η��� 1���� ���� ��ġ,������ �� �˹��η��� �Է� �ϰ�, ���� 2���̻��� ���� 2���̻��ư�� ���� ���λ����� �Է��Ѵ�.</strong>
                                </td>
            				</tr>
						</tbody>
					</table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu1" style="display:none">
<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '01' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '01' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>����ũž<br>��Ʈ��<br>S/W ���</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err01" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						  <%
                                sql = "select count(*) from etc_code where etc_type = '02' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '02' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#E1FFE1">
                                <p>����ũž<br>��Ʈ��<br>H/W ���</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#E1FFE1">
                                <input type="checkbox" name="err02" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
                    	</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu2" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '03' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '03' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>����� ���</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err03" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu3" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '04' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '04' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>������<br>���ɳ�<br>�÷��� ���</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err04" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu4" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '05' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '05' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>������<br>��Ʈ�� ���</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err05" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu5" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '06' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '06' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>��ũ�����̼�<br>���� ���</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err06" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu6" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '07' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '07' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>�ƴ��� ���</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err07" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu7" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '09' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '09' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>��Ÿ</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err09" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
					</table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="att_menu" style="display:none">
						<colgroup>
							<col width="8%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">����÷��1</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file1" type="file" id="att_file1" size="100"></td>
            				</tr>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">����÷��2</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file2" type="file" id="att_file2" size="100"></td>
            				</tr>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">����÷��3</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file3" type="file" id="att_file3" size="100"></td>
            				</tr>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">����÷��4</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file4" type="file" id="att_file4" size="100"></td>
            				</tr>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">����÷��5</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file5" type="file" id="att_file5" size="100"></td>
            				</tr>
						</tbody>
                    </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:goBefore();"></span>
                </div>
                <br>
				<input name="reside_place" type="hidden" id="reside_place" value="<%=rs("reside_place")%>">
                <input name="team" type="hidden" id="team" value="<%=rs("team")%>">
                <input name="sms_old" type="hidden" id="sms_old" value="<%=rs("sms")%>" size="1">
                <input name="be_pg" type="hidden" id="be_pg" value="<%=be_pg%>">
                <input name="write_date" type="hidden" id="write_date" value="<%=rs("write_date")%>">
                <input name="write_cnt" type="hidden" id="write_cnt" value="<%=rs("write_cnt")%>">
                <input name="page" type="hidden" id="page" value="<%=page%>">
                <input name="from_date" type="hidden" id="from_date" value="<%=from_date%>">
                <input name="to_date" type="hidden" id="to_date" value="<%=to_date%>">
                <input name="date_sw" type="hidden" id="date_sw" value="<%=date_sw%>">
                <input name="process_sw" type="hidden" id="process_sw" value="<%=process_sw%>">
                <input name="field_check" type="hidden" id="field_check" value="<%=field_check%>">
                <input name="field_view" type="hidden" id="field_view" value="<%=field_view%>">
                <input name="view_sort" type="hidden" id="view_sort" value="<%=view_sort%>">
                <input name="condi_com" type="hidden" id="condi_com" value="<%=condi_com%>">
                <input name="view_c" type="hidden" id="view_c" value="<%=view_c%>">
                <input name="tel_ddd" type="hidden" id="tel_ddd" value="<%=rs("tel_ddd")%>">
                <input name="tel_no1" type="hidden" id="tel_no1" value="<%=rs("tel_no1")%>">
                <input name="tel_no2" type="hidden" id="tel_no2" value="<%=rs("tel_no2")%>">
                <input name="mg_group" type="hidden" id="mg_group" value="<%=rs("mg_group")%>">
        	</form>
		</div>				
	</div>        				
	</body>
</html>

