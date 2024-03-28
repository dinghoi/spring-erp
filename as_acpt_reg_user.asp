<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)
curr_hh = int(cstr(datepart("h",now)))
curr_mm = int(cstr(datepart("n",now)))
request_date = curr_date
request_hh = curr_hh
request_mm = curr_mm

if curr_hh < 10 then
	curr_hh = "0" + cstr(curr_hh)
end if

if curr_mm < 10 then
	curr_mm = "0" + cstr(curr_mm)
end if

if request_mm < "30" then
	request_mm = "30"
end if

if request_mm > "30" then
	request_mm = "00"
	request_hh = cstr(request_hh + 1)
end if

request_hh = cstr(request_hh + 4)

if request_hh = "18" then
	request_mm = "00"
end if

if request_hh > "18" then
	request_hh = request_hh - 18
	request_date = mid(cstr(now()+1),1,10)
	select case request_hh
		case 1
			request_hh = "10"
		case 2
			request_hh = "11"
		case 3
			request_hh = "12"
		case else
			request_hh = "13"
	end select	
end if

c_w = datepart("w",curr_date)

if c_w = 7 or c_w = 1 then
	request_hh = "13"
	request_mm = "00"
end if

w_cnt = 1
if help_yn = "Y" then
	help_view = "����"
  else
  	help_view = ""
end if

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_memb = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
'Set Rs_hol = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql_type="select * from type_code where etc_type='91' and etc_seq ='"+mg_group+"'"
set rs_type=dbconn.execute(sql_type)
if rs_type.eof then
	mg_group = ""
	mg_group_name = "ERROR"
  else  	
	mg_group = rs_type("etc_seq")
	mg_group_name = rs_type("type_name")
end if
rs_type.Close()		

for k = 1 to 15

	w = datepart("w",request_date)

	if w = 7 then
		request_date = dateadd("d",2,request_date)
	end if
	
	if w = 1 then
		request_date = dateadd("d",1,request_date)
	end if
	Set Rs_hol = Server.CreateObject("ADODB.Recordset")
	Sql="select * from holiday where holiday = '"&request_date&"'"
	Rs_hol.Open Sql, Dbconn, 1
	if 	rs_hol.eof then
		request_date = request_date
		exit for
	else
		request_date = dateadd("d",1,request_date)
	end if

	k = k + 1
'	rs_hol.Close()
next
rs_hol.Close()

title_line = "A/S ���� ���"
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
			function history_view () {
				if(document.frm.company.value =="") {
					alert('�ּ�DB�� �˻��ϼ���');
					return false;}
				if(document.frm.dept.value =="") {
					alert('�ּ�DB�� �˻��ϼ���');
					return false;}
				if(document.frm.acpt_user.value =="") {
					alert('����ڸ� �Է��ϼ���');
					frm.acpt_user.focus();
					return false;}
				var company = document.frm.company.value;
				var dept = document.frm.dept.value;
				var acpt_user = document.frm.acpt_user.value;
				var url = "as_history.asp?company="+company+"&dept="+dept+"&acpt_user="+acpt_user;				
				pop_Window(url,'ceselect','scrollbars=yes,width=1200,height=400');
			}			
			function chkfrm() {
				if(document.frm.company.value =="") {
					alert('�ּ�DB�� �˻��ϼ���');
					return false;}
				if(document.frm.acpt_user.value =="") {
					alert('����ڸ� �Է��ϼ���');
					frm.acpt_user.focus();
					return false;}
				if(document.frm.sido.value =="") {
					alert('������ȸ�� �ϼ���');
					return false;}
				if(document.frm.gugun.value =="") {
					alert('������ȸ�� �ϼ���');
					return false;}
				if(document.frm.dong.value =="") {
					alert('������ȸ�� �ϼ���');
					return false;}
				if(document.frm.addr.value =="") {
					alert('������ �Է��ϼ���');
					frm.addr.focus();
					return false;}
				if(document.frm.mg_ce_id.value =="") {
					if(document.frm.s_ce_id.value =="") {
						alert('��� CE�� �����Ǿ� ���� ����');
						frm.ce_mod.focus();
						return false;}}
				if(document.frm.as_memo.value =="") {
					alert('��ֳ����� �Է��ϼ���');
					frm.as_memo.focus();
					return false;}
				if(document.frm.request_date.value =="") {
					alert('��û���� �Է��ϼ���');
					frm.request_date.focus();
					return false;}
				if(document.frm.request_date.value < document.frm.curr_date.value) {
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
				if(document.frm.request_date.value == document.frm.curr_date.value) {
					if(document.frm.request_hh.value < document.frm.curr_hh.value) {
						alert('��û�ð��� �����ð� ���� �����ϴ�');
						frm.request_hh.focus();
						return false;}}
				if(document.frm.request_date.value == document.frm.curr_date.value) {
					if(document.frm.request_hh.value == document.frm.curr_hh.value) {
						if(document.frm.request_mm.value <= document.frm.curr_mm.value) {
							alert('��û���� ������ ���� �����ϴ�');
							frm.request_mm.focus();
							return false;}}}	

				a=confirm('����Ͻðڽ��ϱ�?');
				if (a==true) {
					return true;
				}
				return false;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/user_header.asp" -->
			<!--#include virtual = "/include/as_sub_menu_user.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_acpt_reg_user_ok.asp" method="post" name="frm">
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
								<th class="first">�ּҷ���ȸ</th>
								<td class="left">
                                <a href="#" class="btnType03" onClick="pop_Window('juso_search_user.asp?company=<%=reside_company%>','juso_search_pop','scrollbars=yes,width=800,height=400')">�ּҷ�DB</a></td>
								<th>������</th>
								<td class="left"><%=now()%>
                                <input name="curr_date" type="hidden" id="now_date2" value="<%=curr_date%>">
              					<input name="curr_hh" type="hidden" id="curr_hh" value="<%=curr_hh%>">
              					<input name="curr_mm" type="hidden" id="curr_mm" value="<%=curr_mm%>">
              					<input name="curr_date_time" type="hidden" id="curr_date_time" value="<%=now()%>">
                                </td>
								<th>������</th>
								<td class="left"><%=user_name%>
                                <input name="acpt_man" type="hidden" value="<%=user_name%>">
            					<input name="help_yn" type="hidden" id="help_yn" value="<%=help_yn%>">
            					<%=help_view%>
                                </td>
								<th>ȸ��</th>
								<td class="left"><input name="company" type="text" id="company"  style="width:150px" readonly="true"></td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td class="left"><input name="dept" type="text" id="dept"  style="width:150px" readonly="true"></td>
								<th>��ȭ��ȣ</th>
								<td class="left"><input name="tel_ddd" type="text" id="tel_ddd2" size="3" maxlength="3" readonly="true">
								  -
                                    <input name="tel_no1" type="text" id="tel_no" size="4" maxlength="4" readonly="true">
                                    -
                                <input name="tel_no2" type="text" id="tel_no2" size="4" maxlength="4" readonly="true"></td>
								<th>�����</th>
								<td class="left"><input name="acpt_user" type="text" size="10" onKeyUp="checklength(this,20)" notnull errname="�����">
								  &nbsp;<strong>����</strong>
                                <input name="user_grade" type="text" size="8" onKeyUp="checklength(this,20)"></td>
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
                                <input name="sido" type="text" id="sido" style="width:50px" readonly="true">
              					<input name="gugun" type="text" id="gugun" style="width:150px" readonly="true">
              					<input name="dong" type="text" id="dong" style="width:150px" readonly="true">
              					<input name="addr" type="text" id="addr" style="width:250px" notnull errname="����" onKeyUp="checklength(this,50)">
              					<input name="view_ok" type="hidden" id="view_ok" value="">
              					<a href="#" class="btnType03" onclick="javascript:pop_area()" >������ȸ</a>
                                </td>
							  <th>���CE</th>
                                <td class="left">
                                  <input name="mg_ce_id" type="text" id="mg_ce_id" size="10" readonly="true">
                                  <input name="mg_ce" type="text" class="ins_form" size="8" readonly="true">
                                  <input name="team" type="hidden" id="team" value="">
                                  <input name="reside_place" type="hidden" id="reside_place" value="">
                                </td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td class="left"><%
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
                                </select></td>
							  <th>������</th>
							  <td class="left"><%
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
                              </select></td>
								<th>�𵨸�</th>
								<td class="left"><input name="model_no" type="text" id="model_no" style="width:150px" notnull errname="�𵨸�" onKeyUp="checklength(this,20)"></td>
								<th>ó������</th>
								<td class="left"><select name="as_type" id="as_type" style="width:150px">
								  <option value="�湮ó��">�湮ó��</option>
								  <option value="�űԼ�ġ">�űԼ�ġ</option>
								  <option value="���ȸ��">���ȸ��</option>
								  <option value="��Ÿ">��Ÿ</option>
							    </select></td>
							</tr>
							<tr>
								<th class="first">��ֳ���</th>
								<td class="left" colspan="7">
                                <textarea name="as_memo" cols="115" rows="5" id="textarea"></textarea>
                                </td>
							</tr>
							<tr>
								<th class="first">��û��/�ð�</th>
								<td class="left">
                                <input name="request_date" type="text" size="10" readonly="true" id="datepicker" style="width:70px;">&nbsp;
                                <input name="request_hh" type="text" id="request_hh" value="<%=request_hh%>" size="2" maxlength="2">
                                <strong>��</strong>
                                <input name="request_mm" type="text" id="request_mm" value="<%=request_mm%>" size="2" maxlength="2"><strong>��</strong>
							  </td>
								<th>�ø����ȣ</th>
								<td class="left"><input name="serial_no" type="text" id="serial_no" style="width:150px" notnull errname="�ø����ȣ" onKeyUp="checklength(this,20)"></td>
								<th>�ڻ��ȣ</th>
								<td class="left"><input name="asets_no" type="text" id="asets_no" style="width:150px" notnull errname="�ڻ��ȣ" onKeyUp="checklength(this,20)"></td>
								<th>Ȯ�μ�����</th>
								<td class="left">
                                <input type="checkbox" name="doc_yn" id="doc_yn" value="Y"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:goBefore();"></span>
                </div>
				</form>
		</div>				
	</div>        				
	</body>
</html>

