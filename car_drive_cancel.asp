<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "���� �������� ���� ���"

run_date = request("run_date")
mg_ce_id = request("mg_ce_id")
run_seq = request("run_seq")

sql = "select * from transit_cost where run_date ='"&run_date&"' and mg_ce_id ='"&mg_ce_id&"' and run_seq ='"&run_seq&"'"
set rs = dbconn.execute(sql)

sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
set rs_memb=dbconn.execute(sql)

if	rs_memb.eof or rs_memb.bof then
	mg_ce = "ERROR"
  else
	mg_ce = rs_memb("user_name")
end if
rs_memb.close()						
car_owner = rs("car_owner")
car_no = rs("car_no")
car_name = rs("car_name")
oil_kind = rs("oil_kind")
start_company = rs("start_company")
start_point = rs("start_point")
start_time = rs("start_time")
start_km = int(rs("start_km"))
end_company = rs("end_company")
end_point = rs("end_point")
end_time = rs("end_time")
end_km = int(rs("end_km"))
far = int(rs("far"))
'	payment = rs("payment")
repair_pay = rs("repair_pay")
repair_cost = int(rs("repair_cost"))
run_memo = rs("run_memo")
oil_amt = int(rs("oil_amt"))
oil_pay = rs("oil_pay")
oil_price = int(rs("oil_price"))
parking_pay = rs("parking_pay")
parking = int(rs("parking"))
toll_pay = rs("toll_pay")
toll = int(rs("toll"))
cancel_yn = rs("cancel_yn")
end_yn = rs("end_yn")
reg_id = rs("reg_id")
reg_date = rs("reg_date")
reg_user = rs("reg_user")
mod_id = rs("mod_id")
mod_date = rs("mod_date")
mod_user = rs("mod_user")
rs.close()

if end_yn = "Y" then
	end_view = "����"
  else
  	end_view = "����"
end if
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
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="car_drive_cancel_save.asp" method="post" name="frm">
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
								<th class="first">������</th>
								<td class="left"><%=run_date%></td>
								<th>������</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)</td>
							</tr>
							<tr>
								<th class="first">��������</th>
								<td colspan="3" class="left"><strong>���� :</strong><%=car_owner%>&nbsp;<strong>������ȣ :</strong><%=car_no%>&nbsp;<strong>���� :</strong><%=car_name%>&nbsp;<strong>���� :</strong><%=oil_kind%></td>
						    </tr>
							<tr>
								<th class="first">���ȸ��</th>
								<td class="left"><%=start_company%></td>
								<th>����ּ�</th>
								<td class="left"><%=start_point%></td>
							</tr>
							<tr>
								<th class="first">���KM</th>
								<td class="left"><%=formatnumber(start_km,0)%></td>
								<th>��߽ð�</th>
								<td class="left"><%=start_time%></td>
							</tr>
							<tr>
								<th class="first">����ȸ��</th>
								<td class="left"><%=end_company%></td>
								<th>�����ּ�</th>
								<td class="left"><%=end_point%></td>
							</tr>
							<tr>
								<th class="first">����KM</th>
								<td class="left"><%=formatnumber(end_km,0)%></td>
								<th>�����ð�</th>
								<td class="left"><%=end_time%></td>
							</tr>
					    	<tr>
								<th class="first">����Ÿ�</th>
								<td class="left"><%=formatnumber(far,0)%></td>
								<th>�������</th>
								<td class="left"><%=run_memo%></td>
							</tr>
							<tr>
								<th class="first">������(L)</th>
								<td class="left"><%=formatnumber(oil_amt,0)%></td>
                                <th>�����ݾ�</th>
								<td class="left"><%=oil_pay%>&nbsp;<%=formatnumber(oil_price,0)%></td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td class="left"><%=parking_pay%>&nbsp;<%=formatnumber(parking,0)%></td>
                                <th>�����</th>
								<td class="left"><%=toll_pay%>&nbsp;<%=formatnumber(toll,0)%></td>
							</tr>
    				  <tr>
						<th class="first">��ҿ���</th>
						<td class="left">
						<input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:30px" ID="Radio1">���           
                        <input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:30px" ID="Radio2">����
                        </td>
                        <th>��������</th>
						<td class="left"><%=end_view%></td>
					</tr>
					<tr>
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
                </div>
				<input type="hidden" name="mg_ce_id" value="<%=mg_ce_id%>" ID="Hidden1">
				<input type="hidden" name="run_date" value="<%=run_date%>" ID="Hidden1">
				<input type="hidden" name="run_seq" value="<%=run_seq%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

