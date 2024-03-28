<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
u_type = request("u_type")

mg_ce_id = user_id
mg_ce = user_name
start_point = ""
start_hh = ""
start_mm = ""
company = ""
end_point = ""
end_hh = ""
end_mm = ""
transit = ""
payment = ""
fare = 0
run_memo = ""
cancel_yn = "N"
end_yn = "N"

curr_date = mid(cstr(now()),1,10)
run_date = mid(cstr(now()),1,10)

strNowWeek = WeekDay(run_date)
Select Case (strNowWeek)
   Case 1
       week = "�Ͽ���"
   Case 2
       week = "������"
   Case 3
       week = "ȭ����"
   Case 4
       week = "������"
   Case 5
       week = "�����"
   Case 6
       week = "�ݿ���"
   Case 7
       week = "�����"
End Select

company = "����"

title_line = "���� ����� ���"
if u_type = "U" then

	run_date = request("run_date")
	mg_ce_id = request("mg_ce_id")
	run_seq = request("run_seq")

	sql = "select * from transit_cost where run_date ='"&run_date&"' and mg_ce_id ='"&mg_ce_id&"' and run_seq ='"&run_seq&"'"
	set rs = dbconn.execute(sql)

	if rs.eof or rs.bof then
		mg_ce = "ERROR"
	  else		
		sql = "select * from memb where user_id = '"&mg_ce_id&"'"
		set rs_memb=dbconn.execute(sql)
	
		if	rs_memb.eof or rs_memb.bof then
			mg_ce = "ERROR"
		  else
			mg_ce = rs_memb("user_name")
		end if
		rs_memb.close()						
	end if
	
	start_point = rs("start_point")
	start_hh = mid(rs("start_time"),1,2)
	start_mm = mid(rs("start_time"),3,2)
	company = rs("company")
	end_point = rs("end_point")
	end_hh = mid(rs("end_time"),1,2)
	end_mm = mid(rs("end_time"),3,2)
	transit = rs("transit")
	payment = rs("payment")
	fare = int(rs("fare"))
	run_memo = rs("run_memo")
	cancel_yn = rs("cancel_yn")
	end_yn = rs("end_yn")
	reg_id = rs("reg_id")
	reg_date = rs("reg_date")
	reg_user = rs("reg_user")
	mod_id = rs("mod_id")
	mod_date = rs("mod_date")
	mod_user = rs("mod_user")
	rs.close()

	title_line = "���� ����� ����"
end if
if end_yn = "Y" then
	end_view = "����"
  else
  	end_view = "����"
end if
if cancel_yn = "Y" then
	cancel_view = "���"
  else
  	cancel_view = "����"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=run_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.run_date.value <= document.frm.end_date.value) {
					alert('�̿����ڰ� ������ �Ǿ� �ִ� �����Դϴ�');
					frm.run_date.focus();
					return false;}
				if(document.frm.run_date.value > document.frm.curr_date.value) {
					alert('�̿����ڰ� �����Ϻ��� Ŭ���� �����ϴ�.');
					frm.run_date.focus();
					return false;}
				if(document.frm.company.value =="" ) {
					alert('��ü�� �����ϼ���');
					frm.company.focus();
					return false;}
				if(document.frm.mg_ce.value =="" ) {
					alert('�̿��� �����Դϴ�. �����ڿ��� ���� �ٶ��ϴ�');
					frm.mg_ce.focus();
					return false;}
				if(document.frm.start_point.value =="" ) {
					alert('����ּ��� �Է��ϼ���');
					frm.start_point.focus();
					return false;}
				if(document.frm.start_hh.value >"23"||document.frm.start_hh.value <"00") {
					alert('��߽ð��� �߸��Ǿ����ϴ�');
					frm.start_hh.focus();
					return false;}
				if(document.frm.start_mm.value >"59"||document.frm.start_mm.value <"00") {
					alert('��ߺ��� �߸��Ǿ����ϴ�');
					frm.start_mm.focus();
					return false;}
				if(document.frm.end_point.value =="" ) {
					alert('�����ּ��� �Է��ϼ���');
					frm.end_point.focus();
					return false;}
				if(document.frm.end_hh.value >"23"||document.frm.end_hh.value <"00") {
					alert('�����ð��� �߸��Ǿ����ϴ�');
					frm.end_hh.focus();
					return false;}
				if(document.frm.end_mm.value >"59"||document.frm.end_mm.value <"00") {
					alert('�������� �߸��Ǿ����ϴ�');
					frm.end_mm.focus();
					return false;}
				if(document.frm.start_hh.value > document.frm.end_hh.value) {
					alert('�����ð��� ��߽ð� ���� �����ϴ�');
					frm.end_hh.focus();
					return false;}
				if(document.frm.start_hh.value == document.frm.end_hh.value) {
					if(document.frm.start_mm.value > document.frm.end_mm.value) {
						alert('�����ð��� ��߽ð� ���� �����ϴ�');
						frm.end_mm.focus();
						return false;}}
				if(document.frm.transit.value =="" ) {
					alert('�������� �����ϼ���');
					frm.transit.focus();
					return false;}
				if(document.frm.fare.value <= 0 ) {
					alert('����� �Է��ϼ���');
					frm.fare.focus();
					return false;}
				if(document.frm.run_memo.value =="" ) {
					alert('�۾������� �����ϼ���');
					frm.run_memo.focus();
					return false;}
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
 			function week_check() {
			
			a = document.frm.run_date.value.substring(0,4);
			b = document.frm.run_date.value.substring(5,7);
			c = document.frm.run_date.value.substring(8,10);
			
			var newDate = new Date(a,b-1,c); 
			var s = newDate.getDay(); 
			
			switch(s) {
				case 0: str = "�Ͽ���" ; break;
				case 1: str = "������" ; break;
				case 2: str = "ȭ����" ; break;
				case 3: str = "������" ; break;
				case 4: str = "�����" ; break;
				case 5: str = "�ݿ���" ; break;
				case 6: str = "�����" ; break;
				}
			
				document.frm.week.value = str;			
			}
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
			function delcheck() 
				{
				a=confirm('���� �����Ͻðڽ��ϱ�?')
				if (a==true) {
					document.frm.action = "mass_transit_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
       </script>
	</head>
	<body onLoad="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="mass_transit_add_save.asp" method="post" name="frm">
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
								<th class="first">�̿�����</th>
								<td class="left">
                                <input name="run_date" type="text" id="datepicker" style="width:70px" value="<%=run_date%>" readonly="true">&nbsp;
                                �������� : <%=end_date%>
							<%  if u_type = "U" then	%>
                                <input name="old_date" type="hidden" value="<%=run_date%>">
                            <%	end if	%>
                                </td>
								<th>�̿���</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)
                                <input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=mg_ce_id%>">
                                <input name="mg_ce" type="hidden" id="mg_ce" value="<%=mg_ce%>">
                                </td>
							</tr>
							<tr>
								<th class="first">��ü</th>
								<td class="left">
							<% if reside_company = "" or isnull(reside_company)	Then	%>
								<%
                                 Sql="select * from trade where (trade_id ='����' or trade_id ='����') and use_sw = 'Y' order by trade_name asc"
                                 Rs_etc.Open Sql, Dbconn, 1
                                %>
                                  <select name="company" id="company" style="width:150px">
                                    <option value="">����</option>
                                    <option value='����' <%If company = "����" then %>selected<% end if %>>����</option>
                                    <% 
                                        do until rs_etc.eof 
                                    %>
                                    <option value='<%=rs_etc("trade_name")%>' <%If rs_etc("trade_name") = company then %>selected<% end if %>><%=rs_etc("trade_name")%></option>
                                    <%
                                        	rs_etc.movenext()  
                                        loop 
                                        rs_etc.Close()
                                    %>
                                </select>
							<%   else	%>
                                    <input name="company" type="text" id="company" style="width:100px" value="<%=reside_company%>" readonly="true" >
                            <% end if	%>
                                </td>
								<th>����ּ�</th>
								<td class="left"><input name="start_point" type="text" id="start_point" style="width:200px" onKeyUp="checklength(this,50)" value="<%=start_point%>"></td>
							</tr>
							<tr>
								<th class="first">��߽ð�</th>
								<td class="left">
                                <input name="start_hh" type="text" id="start_hh" size="2" maxlength="2" value="<%=start_hh%>">��
                                <input name="start_mm" type="text" id="start_mm" size="2" maxlength="2" value="<%=start_mm%>">��
                                </td>
								<th>�����ּ�</th>
								<td class="left"><input name="end_point" type="text" id="end_point" style="width:200px" onKeyUp="checklength(this,50)" value="<%=end_point%>"></td>
							</tr>
							<tr>
								<th class="first">�����ð�</th>
								<td class="left">
                                <input name="end_hh" type="text" id="end_hh" size="2" maxlength="2" value="<%=end_hh%>">��
                                <input name="end_mm" type="text" id="end_mm" size="2" maxlength="2" value="<%=end_mm%>">��
                                </td>
								<th>������</th>
								<td class="left">
                                <select name="transit" id="transit" style="width:80px">
                                    <option value="">����</option>
									<option value='����' <%If transit= "����" then %>selected<% end if %>>����</option>
								  	<option value='����ö' <%If transit= "����ö" then %>selected<% end if %>>����ö</option>
								  	<option value='�ý�' <%If transit= "�ý�" then %>selected<% end if %>>�ý�</option>
								  	<option value='����' <%If transit= "����" then %>selected<% end if %>>����</option>
								  	<option value='�����' <%If transit= "�����" then %>selected<% end if %>>�����</option>
								  	<option value='��' <%If transit= "��" then %>selected<% end if %>>��</option>
								  	<option value='��Ÿ' <%If transit= "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
							    </select></td>
							</tr>
							<tr>
								<th class="first">�����</th>
								<td class="left">���ҹ��
                                  <select name="payment" id="select" style="width:80px">
                                    <option value='����' <%If payment= "����" then %>selected<% end if %>>����</option>
                                </select>                                  
						<% if u_type = "U" then	%>
                                <input name="fare" type="text" id="far2" style="width:80px;text-align:right" value="<%=formatnumber(fare,0)%>" onKeyUp="plusComma(this);">
						<%   else	%>
                                <input name="fare" type="text" id="far2" style="width:80px;text-align:right" onKeyUp="plusComma(this);">
						<% end if	%>
                                </td>
								<th>�۾�����</th>
								<td class="left">
								  <%
                                        Sql="select * from etc_code where etc_type = '42' and used_sw = 'Y' order by etc_code asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                  <select name="run_memo" id="select" style="width:150px">
                                    <option value="">����</option>
                                    <% 
                                        do until rs_etc.eof 
                                    %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If rs_etc("etc_name") = run_memo then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                    <%
                                        	rs_etc.movenext()  
                                        loop 
                                        rs_etc.Close()
                                    %>
                                  </select>
                                </td>
							</tr>
    				  <tr id="cancel_col" style="display:none">
						<th class="first">��ҿ���</th>
						<td class="left"><%=cancel_view%><input type="hidden" name="cancel_yn" value="<%=cancel_yn%>" ID="Hidden1"></td>
                        <th>��������</th>
						<td class="left"><%=end_view%></td>
					</tr>
					<tr id="info_col" style="display:none">
						<th class="first">�������</th>
						<td class="left"><%=reg_user%>&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
                    	<th>��������</th>
						<td class="left"><%=mod_user%>&nbsp;<%=mod_id%>(<%=mod_date%>)</td>
					</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
				<%	
					if u_type = "U" and user_id = mg_ce_id then
						if end_yn = "N" or end_yn = "C" then	
				%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
						end if
					end if	
				%>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="curr_date" value="<%=curr_date%>" ID="Hidden1">
                <input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
				<input type="hidden" name="run_seq" value="<%=run_seq%>" ID="Hidden1">
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
                <input type="hidden" name="mod_id" value="<%=mod_id%>" ID="Hidden1">
                <input type="hidden" name="mod_user" value="<%=mod_user%>" ID="Hidden1">
                <input type="hidden" name="mod_date" value="<%=mod_date%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

