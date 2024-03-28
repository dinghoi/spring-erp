<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
acpt_no = request("acpt_no")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql="select * from as_acpt where acpt_no = " & int(acpt_no)
Set rs=DbConn.Execute(sql)
request_date_time = mid(rs("request_date"),1,10) + " " + mid(rs("request_time"),1,2) + ":" +  mid(rs("request_time"),3,2)
request_date_time = FormatDateTime(request_date_time, 0)

Sql_in="select * from as_into where acpt_no = " & int(acpt_no) & " order by in_seq asc"
Rs_in.Open Sql_in, Dbconn, 1

title_line = "�԰� ���� ����"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�԰� ���� ����</title>
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

			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.into_date.value == "") {
					alert('�԰����ڰ� �����ϴ�.');
					frm.into_date.focus();
					return false;}
				if(document.frm.in_place.value == "����") {
					alert('�԰�ó�� �����ϴ�.');
					frm.in_place.focus();
					return false;}
				if(document.frm.in_process.value == "����") {
					alert('�԰������� �����ϴ�.');
					frm.in_process.focus();
					return false;}
				if(document.frm.in_remark.value == "") {
					alert('���γ����� �����ϴ�.');
					frm.in_remark.focus();
					return false;}
							
				{
				a=confirm('����Ͻðڽ��ϱ�?');
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function insert_off() 
			{
				document.getElementById('into_tab').style.display = 'none';
			}
			function insert_on() 
			{
				document.getElementById('into_tab').style.display = '';
			}
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=into_date%>" );
			});	  
        </script>

	</head>
	<body>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="into_mg_ok.asp?acpt_no=<%=rs("acpt_no")%>">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
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
							  <th>����</th>
							  <td class="left"><%=rs("acpt_user")%></td>
							  <th>��������</th>
							  <td class="left"><%=rs("acpt_date")%></td>
							  <th>������</th>
							  <td class="left"><%=rs("maker")%></td>
					      	</tr>
							<tr>
							  <th>ȸ���</th>
							  <td class="left"><%=rs("company")%></td>
							  <th>������</th>
							  <td class="left" colspan="3"><%=rs("dept")%></td>
					      	</tr>
							<tr>
							  <th>�ּ�</th>
							  <td class="left" colspan="5"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>&nbsp;<%=rs("addr")%></td>
					      	</tr>
							<tr>
							  <th>���CE</th>
							  <td class="left"><%=rs("mg_ce")%>(<%=rs("mg_ce_id")%>)</td>
							  <th>��û����</th>
							  <td class="left"><%=rs("request_date")%></td>
							  <th>�԰����</th>
							  <td class="left"><%=rs("as_device")%></td>
					      	</tr>
						</tbody>
					</table>
					<h3 class="stit">* �԰� History</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="*" >
							<col width="23%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">��������</th>
								<th scope="col">�԰�ó</th>
								<th scope="col">�԰�����</th>
								<th scope="col" class="left">�԰��γ���</th>
								<th scope="col">�������</th>
							</tr>
						</thead>
						<tbody>
						<%
                        i = 0
                        in_end = "N"
                        do until rs_in.eof 
                        %>
							<tr>
								<td class="first"><%=rs_in("into_date")%></td>
								<td><%=rs_in("in_place")%></td>
								<td><%=rs_in("in_process")%></td>
								<td style="text-align:left" class="left"><%=rs_in("in_remark")%></td>
								<td><%=rs_in("reg_name")%>&nbsp;(<%=rs_in("reg_date")%>)</td>
							</tr>
						<%
                            i = i + 1
                            in_seq = rs_in("in_seq")
                            if rs_in("in_process") = "�����Ϸ�" then
                                in_end = "Y"
                            end if
                            rs_in.movenext()  
                        loop
                        rs_in.Close()
                        %>
						</tbody>
					</table>                    
					<br>
               		<div align=right>
						<a href="#" class="btnType04" onclick="javascript:insert_on()" >�߰��Է�</a>&nbsp;
						<a href="#" class="btnType04" onclick="javascript:goAction()" >����</a>&nbsp;&nbsp;
					</div>
                    <br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableView" id="into_tab" style="display:none">
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
							  <th>��������</th>
							  <td class="left">
                              <input name="in_seq" type="hidden" id="in_seq" value="<%=in_seq%>">
							  <input name="into_date" type="text" readonly="true" id="datepicker" style="width:70px;">
                              </td>
							  <th>�԰�ó</th>
							  <td class="left">
                              <select name="in_place" id="in_place">
            					<option value="����">����</option>
            					<option value="��ü�԰�">��ü�԰�</option>
            					<option value="�����԰�">�����԰�</option>
            					<option value="Repair Shop">Repair Shop</option>
                    		  </select>
                    		  </td>
							  <th>�԰�����</th>
							  <td class="left">
                              <select name="in_process" id="in_process">
            					<option value="����">����</option>
            					<option value="�����Ϸ�">�����Ϸ�</option>
            					<option value="����߼�">����߼�</option>
            					<option value="����߼�">����߼�</option>
            					<option value="�����԰�">�����԰�</option>
            					<option value="��ü">��ü</option>
            					<option value="��������">��������</option>
                    		  </select>
                              </td>
					      	</tr>
							<tr>
							  <th>���γ���</th>
							  <td class="left" colspan="5"><textarea name="in_remark" id="in_remark"></textarea></td>
					      	</tr>
							<tr>
							  <td class="center" colspan="6">
                                <div align=center>
                                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();"></span>            
                                    <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:insert_off();"></span>
                                    <% if c_grade = "0" and in_seq = 1 then %>
                                    	<a href="into_del_ok.asp?acpt_no=<%=rs("acpt_no")%>" class="btnType01">����</a>
									<% end if %>            
                                </div>
                              </td>
					      	</tr>
						</tbody>
					</table>
				</form>
				</div>
			</div>
	</body>
</html>

