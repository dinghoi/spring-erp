<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw

slip_month = request("slip_month")
slip_gubun = request("slip_gubun")

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

title_line = slip_month + "�� �󰢺� ����"
savefilename = slip_month + "�� �󰢺� ����.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from general_cost where (slip_gubun = '�󰢺�') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') ORDER BY org_name, emp_name, slip_date ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">���ȸ��</th>
								<th scope="col">�������</th>
								<th scope="col">�����</th>
								<th scope="col">�ݾ�</th>
								<th scope="col">�󰢺�����</th>
								<th scope="col">�󰢺� ���γ���</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							if rs("end_yn") = "Y" then
								end_yn = "����"
								end_view = "N"
							  elseif rs("end_yn") = "I" then
								end_yn = "������"
								end_view = "N"
							  else
							  	end_yn = "����"
							end if
						%>
							<tr>
								<td class="first"><%=rs("emp_company")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("emp_name")%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

