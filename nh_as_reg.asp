<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)
curr_hh = int(cstr(datepart("h",now)))
curr_mm = int(cstr(datepart("n",now)))

asset_company = request.cookies("nkpmg_user")("coo_asset_company")
if asset_company = "00" then
	asset_company = "01"
end if
company = user_name

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_memb = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = company + " A/S �������"
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
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=request_date%>" );
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
			function as_check(yn) {
			if (yn.value == 'Y') 
				{
					document.getElementById('change_addr').style.display = 'none';
					document.form1.as_sw.value = yn.value;
				}
				else 
				{
					document.getElementById('change_addr').style.display = '';
					document.form1.as_sw.value = yn.value;
				}
			
			}
			function chkfrm() {

				if(document.frm.acpt_user.value =="") {
					alert('����ڸ� �Է��ϼ���');
					frm.acpt_user.focus();
					return false;}
				if(document.frm.tel_no1.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_no1.focus();
					return false;}
				if(document.frm.tel_no2.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_no2.focus();
					return false;}
				if(document.frm.org_first.value =="") {
					alert('������ȸ�� �Ͽ� ������ �����ϼ���');
					frm.dept_search.focus();
					return false;}
				if(document.frm.old_sido.value =="") {
					alert('�ش� ������ �ּҸ� Ȯ���ϼ���');
					frm.dept_search.focus();
					return false;}
				if(document.frm.as_sw.value =="N") {
					if(document.frm.sido.value =="") {
						alert('������ȸ�� �Ͽ� �ش� �ּҸ� �����ϼ���');
						frm.area_view.focus();
						return false;}}
				if(document.frm.as_sw.value =="N") {
					if(document.frm.addr.value =="") {
						alert('������ �Է��ϼ���');
						frm.addr.focus();
						return false;}}
				if(document.frm.as_memo.value =="") {
					alert('��ֳ����� �Է��ϼ���');
					frm.as_memo.focus();
					return false;}
			
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/asset_header.asp" -->
			<!--#include virtual = "/include/asset_as_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="nh_as_reg_ok.asp" method="post" name="frm">
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
								<th>������</th>
								<td class="left"><%=now()%>
                                <input name="curr_date" type="hidden" id="now_date2" value="<%=curr_date%>">
              					<input name="curr_hh" type="hidden" id="curr_hh" value="<%=curr_hh%>">
              					<input name="curr_mm" type="hidden" id="curr_mm" value="<%=curr_mm%>">
              					<input name="curr_date_time" type="hidden" id="curr_date_time" value="<%=now()%>">
                                </td>
								<th class="first">������</th>
								<td class="left"><input name="acpt_man" type="text" value="<%="���ͳ�"%>" style="width:150px" readonly="true"></td>
								<th>ȸ��</th>
								<td class="left"><input name="company" type="text" value="<%=company%>" style="width:150px" readonly="true"></td>
								<th>�����</th>
								<td class="left"><input name="acpt_user" type="text" onKeyUp="checklength(this,20)"  style="width:60px">&nbsp;����&nbsp;
								<input name="user_grade" type="text" onKeyUp="checklength(this,20)"  style="width:50px"></td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td class="left" colspan="3">
									<input name="org_first" type="text" id="dept2" size="15" maxlength="30" readonly="true">
            						<input name="org_second" type="text" id="org_second" size="15" maxlength="30" readonly="true">
            						<input name="dept_name" type="text" id="dept_name" size="15" maxlength="30" readonly="true">
									<a href="#" class="btnType03" onClick="pop_Window('dept_search_nh.asp?company=<%=asset_company%>','deptcode','scrollbars=yes,width=600,height=500')">������ȸ</a>
            						<input name="dept_code" type="hidden" id="dept_code" value="">
            						<input name="internet_no" type="hidden" id="internet_no" value="">                                
                                </td>
								<th>��ȭ��ȣ</th>
								<td class="left">
                                <input name="tel_ddd" type="text" id="tel_ddd2" size="3" maxlength="3" readonly="true">
								  -
                                <input name="tel_no1" type="text" id="tel_no" size="4" maxlength="4" readonly="true">
                                  -
                                <input name="tel_no2" type="text" id="tel_no2" size="4" maxlength="4" readonly="true">
                                </td>
								<th>�ڵ���</th>
								<td class="left">
								<select name="hp_ddd" id="hp_ddd">
									<option>����</option>
									<option value="010">010</option>
				  					<option value="011">011</option>
				  					<option value="016">016</option>
				  					<option value="017">017</option>
				  					<option value="018">018</option>
				  					<option value="019">019</option>
								</select>-              	
								<input name="hp_no1" type="text" id="tel_no12" size="4" maxlength="4">-
                            	<input name="hp_no2" type="text" id="tel_no22" size="4" maxlength="4">
                              </td>
							</tr>
							<tr>
								<th class="first">�ּ�</th>
								<td class="left" colspan="5">
			  					<input name="old_sido" type="text" id="sido2" size="6" maxlength="6" readonly="true">
              					<input name="old_gugun" type="text" id="gugun2" size="10" maxlength="10" readonly="true">
              					<input name="old_dong" type="text" id="old_dong" size="18" readonly="true">
              					<input name="old_addr" type="text" class="style12" id="addr3" size="50" maxlength="50" readonly="true">
              					<input name="old_mg_ce" type="text" id="old_mg_ce" value="">
              					<input name="old_mg_ce_id" type="text" id="old_mg_ce_id" value="">
              					<input name="old_team" type="text" id="old_team" value="">
              					<input name="old_reside_place" type="text" id="old_reside_place" value="">
                                </td>
								<th class="left">A/S����</th>
                                <td>
                                <input name="as_yn" type="radio" value="Y" checked onClick="as_check(this)">A/S����
                                <input name="as_yn" type="radio" value="N" onClick="as_check(this)">������ġ 
            					<input name="as_sw" type="hidden" id="as_sw" value="Y">
                                </td>
							</tr>
							<tr id="change_addr" style="display:none">
								<th class="first">�����ּ�</th>
								<td class="left" colspan="7">
			  					<input name="sido" type="text" id="sido6" size="6" maxlength="6" readonly="true">
              					<input name="gugun" type="text" id="gugun7" size="10" maxlength="10" readonly="true">
              					<input name="dong" type="text" id="dong3" size="18" readonly="true">
              					<input name="addr" type="text" class="style12" id="addr6" size="50" maxlength="50">
              					<a href="#" class="btnType03" onclick="javascript:pop_area()" >������ȸ</a>
              					<input name="mg_ce" type="hidden" id="mg_ce" value="">
              					<input name="mg_ce_id" type="hidden" id="mg_ce_id" value="">
              					<input name="team" type="hidden" id="team" value="">
              					<input name="reside_place" type="hidden" id="reside_place" value="">
                                </td>
							</tr>
							<tr>
								<th class="first">��ֳ���</th>
								<td class="left" colspan="7">
                                <textarea name="as_memo" cols="115" rows="5" id="textarea"></textarea>
                                </td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td class="left">
                            <%
								Sql="select * from etc_code where etc_type = '31' order by etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
							%>
								<select name="as_device" id="select" style="width:150px">
                			<% 
								do until rs_etc.eof 
			  				%>
                					<option value=<%=rs_etc("etc_name")%>><%=rs_etc("etc_name")%></option>
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
                					<option value=<%=rs_etc("etc_name")%>><%=rs_etc("etc_name")%></option>
                			<%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							%>
            					</select>
            					</td>
								<th>�𵨸�</th>
								<td class="left"><input name="model_no" type="text" id="model_no" style="width:150px"  onKeyUp="checklength(this,20)"></td>
								<th>�ø���NO</th>
								<td class="left"><input name="serial_no" type="text" id="serial_no" style="width:150px"  onKeyUp="checklength(this,20)"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:goBefore();"></span>
                </div>
        <p>&nbsp;1. ������ȸ ��ư�� ���� A/S �Ǵ� ������ġ�� ���ϴ� ������ �����մϴ�.</p>
        <p>&nbsp;2. ����� ���� �Ǵ� ������ ��ȭ ��ȣ�� A/S ������ �����Ͻð�</p>
        <p>&nbsp;3. ���� ��ġ�� ��� ������ȸ�� ���� �����ϰ��� �ϴ� �ּҸ� �����Ͻ���</p>
        <p>&nbsp;4. ��ֳ��뿡 ���������� ��� ���� �Ǵ� ������ġ�� ��û ������ �Է��Ͻð�</p>
        <p>&nbsp;5. A/S ��� �����ô� �ش� ���� �����縦 �����Ͻð� ��ǰ�� �ø���NO�� �ƽø� �Է��Ͻø� �˴ϴ�.</p>
        <p>&nbsp;6. ���� ���� ��ġ�ô� Ư���� �Ѵ��� ��� ���� ������, ������, �ø���NO�� �Է��Ͻø� �˴ϴ�.</p>
        <p>&nbsp;</p></td>
				</form>
		</div>				
	</div>        				
	</body>
</html>

