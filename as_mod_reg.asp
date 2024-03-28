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
condi_com = request("company")


Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_ddd = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

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

title_line = "A/S ���� ����"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� ����</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}
			function goBefore () {
				window.close () ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.acpt_user.value =="") {
					alert('����ڸ� �Է��ϼ���');
					frm.acpt_user.focus();
					return false;}
				if(document.frm.tel_ddd.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_ddd.focus();
					return false;}
				if(document.frm.tel_no1.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_no1.focus();
					return false;}
				if(document.frm.tel_no2.value =="") {
					alert('��ȭ��ȣ�� �Է��ϼ���');
					frm.tel_no2.focus();
					return false;}
				if(document.frm.dept.value =="") {
					alert('�������� �Է��ϼ���');
					frm.dept.focus();
					return false;}
				if(document.frm.sido.value =="") {
					alert('�ּҷ��� ����ϼ���');
					frm.area_view.focus();
					return false;}
				if(document.frm.gugun.value =="") {
					alert('�ּҷ��� ����ϼ���');
					frm.area_view.focus();
					return false;}
				if(document.frm.dong.value =="") {
					alert('�ּҷ��� ����ϼ���');
					frm.area_view.focus();
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
				/**/	
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
				/**/
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
							
				{
				a=confirm('����Ͻðڽ��ϱ�?');
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
					document.getElementById('ce_mod').style.display = ''; }
				if (document.frm.ce_mod_ck.checked == false) {
					document.getElementById('ce_mod').style.display = 'none'; 
					document.getElementById('s_ce').style.display = 'none'; }
			}
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=rs("request_date")%>" );
			});	  
        </script>

	</head>
	<body>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="as_mod_reg_ok.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="20%" >
							<col width="11%" >
							<col width="*" >
							<col width="11%" >
							<col width="19%" >
						</colgroup>
						<tbody>
							<tr>
							  <th>������ȣ</th>
							  <td class="left"><%=rs("acpt_no")%></td>
							  <th>��������</th>
							  <td class="left"><%=rs("acpt_date")%></td>
							  <th>������</th>
							  <td class="left"><%=rs("acpt_man")%></td>
					    </tr>
							<tr>
								<th>�����/����</th>
							  <td class="left">
							  	<input name="acpt_user" type="text" id="acpt_user" value="<%=rs("acpt_user")%>" size="10">
                	<input name="user_grade" type="text" id="user_grade" value="<%=rs("user_grade")%>" size="6">
                </td>
							  <th>��ȭ��ȣ</th>
							  <td class="left">
								<% 
									Sql="select * from etc_code where etc_type = '71' and used_sw = 'Y' order by etc_code asc"
                  Rs_ddd.Open Sql, Dbconn, 1
                %>
                	<select name="tel_ddd" id="select3">
                  <% 
                  	do until rs_ddd.eof 
                  %>
                  	<option value='<%=rs_ddd("etc_name")%>' <%If rs_ddd("etc_name") = rs("tel_ddd") then %>selected<% end if %>><%=rs_ddd("etc_name")%></option>
                  <%
                  		rs_ddd.movenext()
                  		loop
                  		rs_ddd.close()						
                  %>
                  </select>
                  -
                  <input name="tel_no1" type="text" id="tel_no1" value="<%=rs("tel_no1")%>" size="4" maxlength="4">
                  -
                  <input name="tel_no2" type="text" id="tel_no2" value="<%=rs("tel_no2")%>" size="4" maxlength="4">
                </td>
							  <th>�ڵ���</th>
							  <td class="left">
                	<input name="hp_ddd" type="text" id="hp_ddd" value="<%=rs("hp_ddd")%>" size="3" maxlength="3"> 
                	-
                	<input name="hp_no1" type="text" id="hp_no1" value="<%=rs("hp_no1")%>" size="4" maxlength="4">
                	-
                	<input name="hp_no2" type="text" id="hp_no2" value="<%=rs("hp_no2")%>" size="4" maxlength="4">
                </td>
              </tr>
							<tr>
							  <th>ȸ���</th>
							  <td class="left">
								<%
									sql="select * from trade where use_sw = 'Y' and mg_group = '" + mg_group + "' order by trade_name asc"
									Rs_etc.Open Sql, Dbconn, 1
                %>
                	<select name="company" id="company">
                  <% 
                  	do until rs_etc.eof 
                  %>
                  	<option value='<%=rs_etc("trade_name")%>' <%If rs_etc("trade_name") = rs("company")  then %>selected<% end if %>><%=rs_etc("trade_name")%></option>
                  <%
                  		rs_etc.movenext()  
                      loop 
                      rs_etc.Close()
                  %>
                  </select>
                </td>
							  <th>������</th>
							  <td class="left" colspan="3"><input name="dept" type="text" id="dept" onKeyUp="checklength(this,50)" value="<%=rs("dept")%>" size="30"></td>
					    </tr>
							<tr>
							  <th>�ּ�</th>
							  <td class="left" colspan="5">
							  	<input name="sido" type="text" id="sido3" value="<%=rs("sido")%>" size="6" readonly="true">
                  <input name="gugun" type="text" id="gugun4" value="<%=rs("gugun")%>" size="20" readonly="true">
                  <input name="dong" type="text" value="<%=rs("dong")%>" size="18" readonly="true">
                  <input name="addr" type="text" id="addr" value="<%=rs("addr")%>" size="40" onKeyUp="checklength(this,50)">
                  <input name="view_ok" type="hidden" id="view_ok" value="">
              		<a href="#" class="btnType03" onclick="javascript:pop_area()" >������ȸ</a>
                </td>
					    </tr>
							<tr>
							  <th>����CE</th>
							  <td class="left">
                	<input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=rs("mg_ce_id")%>">
                  <input name="mg_ce" type="text" id="mg_ce" value="<%=rs("mg_ce")%>" size="10" readonly="true">                  
                  <input name="reside_place" type="hidden" id="reside_place" value="">
                  <input name="team" type="hidden" id="team">
                </td>
								<th>CE����</th>
							  <td class="left">
                  <input name="ce_mod_ck" type="checkbox" id="ce_mod_ck" value="1"  onClick="ce_mod_view()">
                  <input name="s_ce_id" type="hidden" value="<%=user_id%>">
                  <input name="s_ce" type="text" value="<%=user_name%>" size="8" readonly="true" style="display:none">
                  <input name="s_reside_place" type="hidden" id="s_reside_place2" value="">
                  <input name="s_team" type="hidden" id="s_reside2">
             			<a href="#" class="btnType03" onClick="pop_Window('ce_select.asp?gubun=<%="����"%>&mg_group=<%=mg_group%>','ceselect','scrollbars=yes,width=500,height=400')">CE����</a>
                </td>
							  <th>���ڹ߼�</th>
							  <td class="left">
                  <input type="radio" name="sms_yn" value="Y">�߼�
                	<input name="sms_yn" type="radio" value="N" checked>�߼۾���
                </td>
					    </tr>
							<tr>
							  <th>��ֳ���</th>
							  <td class="left" colspan="3">
							  	<textarea name="as_memo" cols="115" rows="5" class="style12"><%=rs("as_memo")%></textarea>
                </td>
                <th>��������</th>
								<td class="left">
                				<input type="radio" name="cowork_yn" value="N" <% if rs("cowork_yn") = "N" then %>checked<% end if %>>�Ϲ� 
              					<input type="radio" name="cowork_yn" value="Y" <% if rs("cowork_yn") = "Y" then %>checked<% end if %>>���� 
                </td>
					    </tr>
							<tr>
							  <th>��û����</th>
							  <td class="left" colspan="3">
              					<input name="request_date" type="text" id="datepicker" style="width:70px;" readonly="true">&nbsp;
                                <input name="request_hh" type="text" id="request_hh" value="<%=mid(rs("request_time"),1,2)%>" size="2" maxlength="2">��
                                <input name="request_mm" type="text" id="request_mm" value="<%=mid(rs("request_time"),3,2)%>" size="2" maxlength="2">��
	                          </td>
							  <th>ó������</th>
							  <td class="left">
                                <select name="as_type" id="select2">
                                  <option value="����ó��" <%If Rs("as_type") = "����ó��" then %>selected<% end if %>>����ó��</option>
                                  <option value="�湮ó��" <%If Rs("as_type") = "�湮ó��" then %>selected<% end if %>>�湮ó��</option>
                                  <option value="�űԼ�ġ" <%If Rs("as_type") = "�űԼ�ġ" then %>selected<% end if %>>�űԼ�ġ</option>
                                  <option value="�űԼ�ġ����" <%If Rs("as_type") = "�űԼ�ġ����" then %>selected<% end if %>>�űԼ�ġ����</option>
                                  <option value="������ġ" <%If Rs("as_type") = "������ġ" then %>selected<% end if %>>������ġ</option>
                                  <option value="������ġ����" <%If Rs("as_type") = "������ġ����" then %>selected<% end if %>>������ġ����</option>
                                  <option value="������" <%If Rs("as_type") = "������" then %>selected<% end if %>>������</option>
                                  <option value="����������" <%If Rs("as_type") = "����������" then %>selected<% end if %>>����������</option>
                                  <option value="���ȸ��" <%If Rs("as_type") = "���ȸ��" then %>selected<% end if %>>���ȸ��</option>
                                  <option value="��������" <%If Rs("as_type") = "��������" then %>selected<% end if %>>��������</option>
                                  <option value="��Ÿ" <%If Rs("as_type") = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                                </select>
                              </td>
					      	</tr>
						</tbody>
					</table>
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                        <span class="btnType01"><input type="button" value="���" onclick="javascript:goBefore();"></span>
                    </div>
                    <input name="acpt_no" type="hidden" id="acpt_no" value="<%=rs("acpt_no")%>">
                    <input name="acpt_date" type="hidden" id="acpt_date" value="<%=acpt_date%>">
                    <input name="acpt_hh" type="hidden" id="acpt_hh" value="<%=acpt_hh%>">
                    <input name="acpt_mm" type="hidden" id="acpt_mm2" value="<%=acpt_mm%>">
                    <input name="as_type_old" type="hidden" id="as_type_old2" value="<%=rs("as_type")%>">
                    <input name="sms_old" type="hidden" id="sms_old" value="<%=rs("sms")%>">
                    <input name="be_pg" type="hidden" id="be_pg" value="<%=be_pg%>">
                    <input name="page" type="hidden" id="page" value="<%=page%>">
                    <input name="from_date" type="hidden" id="from_date" value="<%=from_date%>">
                    <input name="to_date" type="hidden" id="to_date" value="<%=to_date%>">
                    <input name="date_sw" type="hidden" id="date_sw" value="<%=date_sw%>">
                    <input name="process_sw" type="hidden" id="process_sw" value="<%=process_sw%>">
                    <input name="field_check" type="hidden" id="field_check" value="<%=field_check%>">
                    <input name="field_view" type="hidden" id="field_view" value="<%=field_view%>">
                    <input name="condi_com" type="hidden" id="condi_com" value="<%=condi_com%>">
				</form>
				</div>
			</div>
	</body>
</html>

