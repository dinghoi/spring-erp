<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "���� ���� ��Ư�� ���"

work_man = 1
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=work_date%>" );
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
				if(document.frm.acpt_no.value =="" || document.frm.acpt_no.value =="0") {
					alert('A/S������ ������ �ȵ� !!!');
					frm.as_view.focus();
					return false;}			
				if(document.frm.from_hh.value >"23"||document.frm.from_hh.value <"00") {
					alert('���� �ð��� �߸��Ǿ����ϴ�');
					frm.from_hh.focus();
					return false;}
				if(document.frm.from_mm.value >"59"||document.frm.from_mm.value <"00") {
					alert('���� ���� �߸��Ǿ����ϴ�');
					frm.from_mm.focus();
					return false;}
				if(document.frm.to_hh.value >"23"||document.frm.to_hh.value <"00") {
					alert('���� �ð��� �߸��Ǿ����ϴ�');
					frm.to_hh.focus();
					return false;}
				if(document.frm.to_mm.value >"59"||document.frm.to_mm.value <"00") {
					alert('���� ���� �߸��Ǿ����ϴ�');
					frm.to_mm.focus();
					return false;}			
				if(document.frm.to_hh.value < document.frm.from_hh.value) {
					alert('����ð��� ���۽ð� ���� �����ϴ�');
					frm.to_hh.focus();
					return false;}
			
				if(document.frm.from_hh.value == document.frm.to_hh.value) {
					if(document.frm.to_mm.value <= document.frm.from_mm.value) {
						alert('����ð��� ���۽ð� ���� �����ϴ�');
						frm.to_mm.focus();
						return false;}}
				
				if(document.frm.work_item.value =="") {
					alert('�۾��׸��� �����ϼ���');
					frm.work_item.focus();
					return false;}

				if(document.frm.work_gubun.value =="") {
					alert('�۾������� �����ϼ���');
					frm.work_gubun.focus();
					return false;}
			
				if(document.frm.work_cnt.value <"1") {
					alert('���� ������ �߸��Ǿ����ϴ�');
					frm.work_cnt.focus();
					return false;}
			
				if(document.frm.work_man.value <"1") {
					alert('�ҿ� �η��� �߸� �Ǿ����ϴ�');
					frm.work_man.focus();
					return false;}
			
				if(document.frm.work_man.value >"0") {
					if(document.frm.mg_ce1.value == "") {
						alert('1��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view1.focus();
						return false;}}
			
				if(document.frm.work_man.value >"1") {
					if(document.frm.mg_ce2.value == "") {
						alert('2��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view2.focus();
						return false;}}
			
				if(document.frm.work_man.value >"2") {
					if(document.frm.mg_ce3.value == "") {
						alert('3��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view3.focus();
						return false;}}
			
				if(document.frm.work_man.value >"3") {
					if(document.frm.mg_ce4.value == "") {
						alert('4��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view4.focus();
						return false;}}
			
				if(document.frm.work_man.value >"4") {
					if(document.frm.mg_ce5.value == "") {
						alert('5��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view5.focus();
						return false;}}
			
				if(document.frm.work_man.value >"5") {
					if(document.frm.mg_ce6.value == "") {
						alert('6��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view6.focus();
						return false;}}
			
				if(document.frm.work_man.value >"6") {
					if(document.frm.mg_ce7.value == "") {
						alert('7��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view7.focus();
						return false;}}
			
				if(document.frm.work_man.value >"7") {
					if(document.frm.mg_ce8.value == "") {
						alert('8��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view8.focus();
						return false;}}
			
				if(document.frm.work_man.value >"8") {
					if(document.frm.mg_ce9.value == "") {
						alert('9��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view9.focus();
						return false;}}
			
				if(document.frm.work_man.value >"9") {
					if(document.frm.mg_ce10.value == "") {
						alert('10��° �۾��ڰ� ������ ���� �ʾҽ��ϴ�');
						frm.ce_view10.focus();
						return false;}}
			
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function menu1() {
			var c = document.frm.work_man.value;
			var d = document.frm.work_cnt.value;
				if (c == '0' || c == '') 
				{
					document.getElementById('ce_01').style.display = 'none';
					document.getElementById('ce_02').style.display = 'none';
					document.getElementById('ce_03').style.display = 'none';
					document.getElementById('ce_04').style.display = 'none';
					document.getElementById('ce_05').style.display = 'none';
					document.getElementById('ce_06').style.display = 'none';
					document.getElementById('ce_07').style.display = 'none';
					document.getElementById('ce_08').style.display = 'none';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '1') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = 'none';
					document.getElementById('ce_03').style.display = 'none';
					document.getElementById('ce_04').style.display = 'none';
					document.getElementById('ce_05').style.display = 'none';
					document.getElementById('ce_06').style.display = 'none';
					document.getElementById('ce_07').style.display = 'none';
					document.getElementById('ce_08').style.display = 'none';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '2') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = 'none';
					document.getElementById('ce_04').style.display = 'none';
					document.getElementById('ce_05').style.display = 'none';
					document.getElementById('ce_06').style.display = 'none';
					document.getElementById('ce_07').style.display = 'none';
					document.getElementById('ce_08').style.display = 'none';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '3') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = '';
					document.getElementById('ce_04').style.display = 'none';
					document.getElementById('ce_05').style.display = 'none';
					document.getElementById('ce_06').style.display = 'none';
					document.getElementById('ce_07').style.display = 'none';
					document.getElementById('ce_08').style.display = 'none';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '4') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = '';
					document.getElementById('ce_04').style.display = '';
					document.getElementById('ce_05').style.display = 'none';
					document.getElementById('ce_06').style.display = 'none';
					document.getElementById('ce_07').style.display = 'none';
					document.getElementById('ce_08').style.display = 'none';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '5') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = '';
					document.getElementById('ce_04').style.display = '';
					document.getElementById('ce_05').style.display = '';
					document.getElementById('ce_06').style.display = 'none';
					document.getElementById('ce_07').style.display = 'none';
					document.getElementById('ce_08').style.display = 'none';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '6') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = '';
					document.getElementById('ce_04').style.display = '';
					document.getElementById('ce_05').style.display = '';
					document.getElementById('ce_06').style.display = '';
					document.getElementById('ce_07').style.display = 'none';
					document.getElementById('ce_08').style.display = 'none';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '7') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = '';
					document.getElementById('ce_04').style.display = '';
					document.getElementById('ce_05').style.display = '';
					document.getElementById('ce_06').style.display = '';
					document.getElementById('ce_07').style.display = '';
					document.getElementById('ce_08').style.display = 'none';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '8') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = '';
					document.getElementById('ce_04').style.display = '';
					document.getElementById('ce_05').style.display = '';
					document.getElementById('ce_06').style.display = '';
					document.getElementById('ce_07').style.display = '';
					document.getElementById('ce_08').style.display = '';		
					document.getElementById('ce_09').style.display = 'none';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '9') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = '';
					document.getElementById('ce_04').style.display = '';
					document.getElementById('ce_05').style.display = '';
					document.getElementById('ce_06').style.display = '';
					document.getElementById('ce_07').style.display = '';
					document.getElementById('ce_08').style.display = '';		
					document.getElementById('ce_09').style.display = '';
					document.getElementById('ce_10').style.display = 'none';		
				}
				if (c == '10') 
				{
					document.getElementById('ce_01').style.display = '';
					document.getElementById('ce_02').style.display = '';
					document.getElementById('ce_03').style.display = '';
					document.getElementById('ce_04').style.display = '';
					document.getElementById('ce_05').style.display = '';
					document.getElementById('ce_06').style.display = '';
					document.getElementById('ce_07').style.display = '';
					document.getElementById('ce_08').style.display = '';		
					document.getElementById('ce_09').style.display = '';
					document.getElementById('ce_10').style.display = '';		
				}
			}
			function overtime() {
			var o = document.frm.work_gubun.options[document.frm.work_gubun.selectedIndex].value;
				if (o == '����') 
				{
					document.frm.overtime_amt.value = '0';
				}
				if (o == '�߱�') 
				{
					document.frm.overtime_amt.value = '15,000';
				}
				if (o == '����') 
				{
					document.frm.overtime_amt.value = '30,000';
				}
				if (o == '����') 
				{
					document.frm.overtime_amt.value = '50,000';
				}
				if (o == '����') 
				{
					document.frm.overtime_amt.value = '70,000';
				}
				if (o == '��Ÿ') 
				{
					document.frm.overtime_amt.value = '100,000';
				}
				
			}
        </script>
	</head>
	<body onLoad="menu1()">
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="overtime_as_add_save.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>����NO</th>
							  <td class="left">
							  <input name="acpt_no" type="text" id="acpt_no" style="width:80px" readonly="true">
							  <a href="#" class="btnType03" onClick="pop_Window('as_search.asp?work_item=<%=work_item%>','as_search','scrollbars=yes,width=700,height=400')">������ȸ</a>
                              </td>
							  <th>ȸ��</th>
							  <td class="left"><input name="company" type="text" id="company" style="width:150px" readonly="true"></td>
							  <th>������</th>
							  <td class="left"><input name="dept" type="text" id="dept" style="width:150px" readonly="true"></td>
					      	</tr>
							<tr>
							  <th>�۾���</th>
							  <td class="left"><input name="work_date" type="text" style="width:80px" readonly="true">
                              <input name="week" type="text" style="width:50px" readonly="true"></td>
							  <th>�۾��׸�</th>
							  <td class="left">
                                <select name="work_item" id="work_item" style="width:150px">
                                    <option value="">����</option>
                                    <option value="��ġ/����" <%If work_item = "��ġ/����" then %>selected<% end if %>>��ġ/����</option>
                                    <option value="��ġ" <%If work_item = "��ġ" then %>selected<% end if %>>��ġ</option>
                                    <option value="����" <%If work_item = "����" then %>selected<% end if %>>����</option>
                                    <option value="������ġ/����" <%If work_item = "������ġ/����" then %>selected<% end if %>>������ġ/����</option>
                                    <option value="������ġ" <%If work_item = "������ġ" then %>selected<% end if %>>������ġ</option>
                                    <option value="��������" <%If work_item = "��������" then %>selected<% end if %>>��������</option>
                                    <option value="���" <%If work_item = "���" then %>selected<% end if %>>���</option>
                                    <option value="��������" <%If work_item = "��������" then %>selected<% end if %>>��������</option>
                                    <option value="���ȸ��" <%If work_item = "���ȸ��" then %>selected<% end if %>>���ȸ��</option>
                                    <option value="��Ÿ" <%If work_item = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                                </select>
                              </td>
							  <th>�۾��ð�</th>
							  <td class="left">
                                <input name="from_hh" type="text" id="from_hh6" size="2" maxlength="2">��
                                <input name="from_mm" type="text" id="from_mm4" size="2" maxlength="2">�� ~
                                <input name="to_hh" type="text" id="to_hh4" size="2" maxlength="2">��
                                <input name="to_mm" type="text" id="to_mm4" size="2" maxlength="2">��
                              </td>
					      	</tr>
							<tr>
							  <th>�۾�����</th>
							  <td class="left">
                                <select name="work_gubun" id="select5" onChange="overtime()" style="width:150px">
                                    <option value="">����</option>
                                    <option value="�߱�" <%If work_gubun = "�߱�" then %>selected<% end if %>>�߱�</option>
                                    <option value="����" <%If work_gubun = "����" then %>selected<% end if %>>����</option>
                                    <option value="����" <%If work_gubun = "����" then %>selected<% end if %>>����</option>
                                    <option value="����" <%If work_gubun = "����" then %>selected<% end if %>>����</option>
                                    <option value="��Ÿ" <%If work_gubun = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                                </select>
                              </td>
							  <th>��û�ݾ�</th>
								<td class="left"><input name="overtime_amt" type="text" id="overtime_amt" value="0" style="width:150px;text-align:right" readonly="true"></td>
							  <th>�۾�����</th>
							  <td class="left">
							  ���� <input name="acpt_cnt" type="text" id="acpt_cnt" size="3" onlynum  errname="��������" maxlength="3" readonly="true" style="text-align:right">
							  &nbsp;/&nbsp;���� <input name="work_cnt" type="text" id="work_cnt" onlynum  errname="��������" size="3" maxlength="3" style="text-align:right">                              </td>
					      	</tr>
							<tr>
							  <th>û���ݾ�</th>
							  <td class="left"><input name="ask_amt" type="text" id="ask_amt" style="text-align:right;width:150px"  onKeyUp="plusComma(this);" value="0"></td>
							  <th>�ҿ��η�</th>
								<td class="left"><input name="work_man" type="text" id="work_man" onChange="menu1()" value="1" size="2" onlynum  errname="�ҿ��η�" maxlength="2" style="text-align:right"></td>
							  <th>�����</th>
							  <td class="left">
								<input name="mg_ce_id" type="text" id="mg_ce_id" style="width:80px" readonly="true">
            					<input name="mg_ce" type="text" id="mg_ce" style="width:80px" readonly="true">
                              </td>
					      	</tr>
						</tbody>
					</table>
          <h3 class="stit">* �۾��� ����</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="10%" >
							<col width="10%" >
							<col width="6%" >
							<col width="10%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">NO</th>
								<th scope="col">�η°˻�</th>
								<th scope="col">�̸�</th>
								<th scope="col">����</th>
								<th scope="col">���̵�</th>
								<th scope="col">�����</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">�Ҽ�</th>
							</tr>
						</thead>
						<tbody>
			  				<tr id="ce_01"  style="display:none">
								<td class="first">1</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=1%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce1" type="text" id="mg_ce1" style="width:80px" readonly="true"></td>
								<td><input name="grade1" type="text" id="grade1" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id1" type="text" id="mg_ce_id1" style="width:80px" readonly="true"></td>
								<td><input name="bonbu1" type="text" id="bonbu1" style="width:140px" readonly="true"></td>
								<td><input name="saupbu1" type="text" id="saupbu1" style="width:140px" readonly="true"></td>
								<td><input name="team1" type="text" id="team1" style="width:140px" readonly="true"></td>
								<td><input name="belong1" type="text" id="belong1" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_02"  style="display:none">
								<td class="first">2</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=2%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce2" type="text" id="mg_ce2" style="width:80px" readonly="true"></td>
								<td><input name="grade2" type="text" id="grade2" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id2" type="text" id="mg_ce_id2" style="width:80px" readonly="true"></td>
								<td><input name="bonbu2" type="text" id="bonbu2" style="width:140px" readonly="true"></td>
								<td><input name="saupbu2" type="text" id="saupbu2" style="width:140px" readonly="true"></td>
								<td><input name="team2" type="text" id="team2" style="width:140px" readonly="true"></td>
								<td><input name="belong2" type="text" id="belong2" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_03"  style="display:none">
								<td class="first">3</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=3%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce3" type="text" id="mg_ce3" style="width:80px" readonly="true"></td>
								<td><input name="grade3" type="text" id="grade3" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id3" type="text" id="mg_ce_id3" style="width:80px" readonly="true"></td>
								<td><input name="bonbu3" type="text" id="bonbu3" style="width:140px" readonly="true"></td>
								<td><input name="saupbu3" type="text" id="saupbu3" style="width:140px" readonly="true"></td>
								<td><input name="team3" type="text" id="team3" style="width:140px" readonly="true"></td>
								<td><input name="belong3" type="text" id="belong3" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_04"  style="display:none">
								<td class="first">4</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=4%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce4" type="text" id="mg_ce4" style="width:80px" readonly="true"></td>
								<td><input name="grade4" type="text" id="grade4" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id4" type="text" id="mg_ce_id4" style="width:80px" readonly="true"></td>
								<td><input name="bonbu4" type="text" id="bonbu4" style="width:140px" readonly="true"></td>
								<td><input name="saupbu4" type="text" id="saupbu4" style="width:140px" readonly="true"></td>
								<td><input name="team4" type="text" id="team4" style="width:140px" readonly="true"></td>
								<td><input name="belong4" type="text" id="belong4" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_05"  style="display:none">
								<td class="first">5</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=5%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce5" type="text" id="mg_ce5" style="width:80px" readonly="true"></td>
								<td><input name="grade5" type="text" id="grade5" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id5" type="text" id="mg_ce_id5" style="width:80px" readonly="true"></td>
								<td><input name="bonbu5" type="text" id="bonbu5" style="width:140px" readonly="true"></td>
								<td><input name="saupbu5" type="text" id="saupbu5" style="width:140px" readonly="true"></td>
								<td><input name="team5" type="text" id="team5" style="width:140px" readonly="true"></td>
								<td><input name="belong5" type="text" id="belong5" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_06"  style="display:none">
								<td class="first">6</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=6%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce6" type="text" id="mg_ce6" style="width:80px" readonly="true"></td>
								<td><input name="grade6" type="text" id="grade6" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id6" type="text" id="mg_ce_id6" style="width:80px" readonly="true"></td>
								<td><input name="bonbu6" type="text" id="bonbu6" style="width:140px" readonly="true"></td>
								<td><input name="saupbu6" type="text" id="saupbu6" style="width:140px" readonly="true"></td>
								<td><input name="team6" type="text" id="team6" style="width:140px" readonly="true"></td>
								<td><input name="belong6" type="text" id="belong6" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_07"  style="display:none">
								<td class="first">7</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=7%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce7" type="text" id="mg_ce7" style="width:80px" readonly="true"></td>
								<td><input name="grade7" type="text" id="grade7" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id7" type="text" id="mg_ce_id7" style="width:80px" readonly="true"></td>
								<td><input name="bonbu7" type="text" id="bonbu7" style="width:140px" readonly="true"></td>
								<td><input name="saupbu7" type="text" id="saupbu7" style="width:140px" readonly="true"></td>
								<td><input name="team7" type="text" id="team7" style="width:140px" readonly="true"></td>
								<td><input name="belong7" type="text" id="belong7" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_08"  style="display:none">
								<td class="first">8</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=8%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce8" type="text" id="mg_ce8" style="width:80px" readonly="true"></td>
								<td><input name="grade8" type="text" id="grade8" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id8" type="text" id="mg_ce_id8" style="width:80px" readonly="true"></td>
								<td><input name="bonbu8" type="text" id="bonbu8" style="width:140px" readonly="true"></td>
								<td><input name="saupbu8" type="text" id="saupbu8" style="width:140px" readonly="true"></td>
								<td><input name="team8" type="text" id="team8" style="width:140px" readonly="true"></td>
								<td><input name="belong8" type="text" id="belong8" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_09"  style="display:none">
								<td class="first">9</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=9%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce9" type="text" id="mg_ce9" style="width:80px" readonly="true"></td>
								<td><input name="grade9" type="text" id="grade9" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id9" type="text" id="mg_ce_id9" style="width:80px" readonly="true"></td>
								<td><input name="bonbu9" type="text" id="bonbu9" style="width:140px" readonly="true"></td>
								<td><input name="saupbu9" type="text" id="saupbu9" style="width:140px" readonly="true"></td>
								<td><input name="team9" type="text" id="team9" style="width:140px" readonly="true"></td>
								<td><input name="belong9" type="text" id="belong9" style="width:140px" readonly="true"></td>
							</tr>
			  				<tr id="ce_10"  style="display:none">
								<td class="first">10</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=10%>','ce_search','scrollbars=yes,width=650,height=400')">��ȸ</a></td>
								<td><input name="mg_ce10" type="text" id="mg_ce10" style="width:80px" readonly="true"></td>
								<td><input name="grade10" type="text" id="grade10" style="width:40px" readonly="true"></td>
								<td><input name="mg_ce_id10" type="text" id="mg_ce_id10" style="width:80px" readonly="true"></td>
								<td><input name="bonbu10" type="text" id="bonbu10" style="width:140px" readonly="true"></td>
								<td><input name="saupbu10" type="text" id="saupbu10" style="width:140px" readonly="true"></td>
								<td><input name="team10" type="text" id="team10" style="width:140px" readonly="true"></td>
								<td><input name="belong10" type="text" id="belong10" style="width:140px" readonly="true"></td>
							</tr>
						</tbody>
					</table>                    
				</form>
					<br>
     				<div class="noprint">
                   		<div align=center>
                            <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();"></span>
                            <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goBefore();"></span>
                    	</div>
    				</div>
				</div>
			</div>
	</body>
</html>

