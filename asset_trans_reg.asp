<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
asset_no = request("asset_no")
dept_name = request("dept_name")
company_name = request("company_name")
install_date = mid(now(),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql="select * from asset where asset_no='" + asset_no + "'"
set rs=dbconn.execute(sql)

sql="select * from asset_code where company='" + rs("company") + "' and gubun='" + rs("gubun") + "' and code_seq='" + rs("code_seq") + "'"
set rs_code=dbconn.execute(sql)

curr_date = mid(cstr(now()),1,10)
curr_hh = int(cstr(datepart("h",now)))
curr_mm = int(cstr(datepart("n",now)))

title_line = "�ڻ� ���� ���"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
												$( "#datepicker" ).datepicker("setDate", "<%=install_date%>" );
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

				if(document.frm.dept_code.value == "") {
					alert('��ġ������ Ȯ�� �ٶ��ϴ�');
					frm.dept_search.focus();
					return false;}
				if(document.frm.user_name.value == "") {
					alert('����ڸ� �Է��ϼ���!!');
					frm.user_name.focus();
					return false;}
				if(document.frm.install_date.value == "") {
					alert('��ġ���ڸ� �Է��ϼ���!!');
					frm.install_date.focus();
					return false;}
				if(document.frm.install_date.value < document.frm.curr_date.value) {
					alert('��ġ���ڰ� �������ں��� ������ �����ϴ�');
					frm.install_date.focus();
					return false;}
				if(document.frm.request_hh.value >"23"||document.frm.request_hh.value <"00") {
					alert('��û�ð��� �߸��Ǿ����ϴ�');
					frm.request_hh.focus();
					return false;}
				if(document.frm.request_mm.value >"59"||document.frm.request_mm.value <"00") {
					alert('��û���� �߸��Ǿ����ϴ�');
					frm.request_mm.focus();
					return false;}
				if(document.frm.install_date.value == document.frm.curr_date.value) {
					if(document.frm.request_hh.value < document.frm.curr_hh.value) {
						alert('��û�ð��� �����ð� ���� �����ϴ�');
						frm.request_hh.focus();
						return false;}}
				if(document.frm.install_date.value == document.frm.curr_date.value) {
					if(document.frm.request_hh.value == document.frm.curr_hh.value) {
						if(document.frm.request_mm.value <= document.frm.curr_mm.value) {
							alert('��û���� ������ ���� �����ϴ�');
							frm.request_mm.focus();
							return false;}}}	
				if(document.frm.trans_memo.value == "") {
					alert('���������� �Է��ϼ���!!');
					frm.trans_memo.focus();
					return false;}
				{
				a=confirm('���� ��ġ ����� �Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="asset_trans_reg_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">�ڻ��ȣ</th>
								<td class="left">
								<%=mid(asset_no,1,2)%>-<%=mid(asset_no,3,6)%>-<%=right(asset_no,4)%>
                                <input name="asset_no" type="hidden" id="asset_no" value="<%=asset_no%>">
                                <input name="serial_no" type="hidden" id="serial_no" value="<%=rs("serial_no")%>">
                                <input name="gubun" type="hidden" id="gubun" value="<%=rs("gubun")%>">
                                <input name="maker" type="hidden" id="maker" value="<%=rs_code("maker")%>"> 
                                </td>
							</tr>
							<tr>
								<th class="first">�ڻ��</th>
								<td class="left"><%=rs("asset_name")%><input name="company_name" type="hidden" id="company_name" value="<%=company_name%>"></td>
							</tr>
							<tr>
								<th class="first">��������</th>
								<td class="left"><%=dept_name%>&nbsp;<%=user_name%>
                                <input name="old_code" type="hidden" id="old_code" value="<%=rs("dept_code")%>">
                                <input name="old_name" type="hidden" id="old_name" value="<%=dept_name%>">
            					</td>
							</tr>
							<tr>
								<th class="first">��ġ����</th>
								<td class="left">
								<input name="dept_name" type="text" id="dept_name" style="width:150px" readonly="true">
								<a href="#" class="btnType03" onClick="pop_Window('dept_search.asp?company=<%=rs("company")%>','deptcode','scrollbars=yes,width=550,height=400')">������ȸ</a><input name="dept_code" type="hidden" id="dept_code" value="">
                                </td>
							</tr>
							<tr>
								<th class="first">�����</th>
								<td class="left">
                                <input name="user_name" type="text" id="user_name" style="width:150px" onKeyUp="checklength(this,20)">
								<input name="old_user" type="hidden" id="old_user" value="<%=rs("user_name")%>">
                                </td>
							</tr>
							<tr>
								<th class="first">��������</th>
								<td class="left">
                                <input name="install_date" type="text" value="<%=install_date%>" style="width:70px" id="datepicker">              
                                <input name="old_date" type="hidden" id="old_date" value="<%=rs("install_date")%>">
                                <input name="curr_date" type="hidden" id="curr_date" value="<%=curr_date%>">
                                ��û�ð�
                                <input name="request_hh" type="text" id="request_hh" size="3" maxlength="2"> 
                                �� 
                                <input name="request_mm" type="text" id="request_mm" size="3" maxlength="2">
                                ��
                                <input name="curr_hh" type="hidden" id="curr_hh" value="<%=curr_hh%>">
                                <input name="curr_mm" type="hidden" id="curr_mm" value="<%=curr_mm%>">
                                </td>
							</tr>
							<tr>
								<th class="first">��������</th>
								<td class="left"><input name="trans_memo" type="text" id="trans_memo" style="width:250px" onKeyUp="checklength(this,50)"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

