<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

use_sw=Request("use_sw")
view_condi=Request("view_condi")
condi=Request("condi")

if condi = "" then
	condi_view = "����"
  else
  	condi_view = condi
end if

if use_sw = "Y" then
	use_view = "���"
  elseif use_sw = "N" then
  	use_view = "�̻��"
  else
  	use_view = "�Ѱ�"
end if 

title_line = "��뱸�� : " + use_view + " , ��ȸ���� : " + condi_view + " - �ŷ�ó����"
savefilename = cstr(now()) + " �ŷ�ó.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

if view_condi = "��ü" and use_sw = "T" then
	where_sql = " "
  else
  	where_sql = " where "
end if

if view_condi = "��ü" then
	condi_sql = " "
  else
	if condi = "" then
		condi_sql = view_condi + " = '" + condi + "'"
	  else
		condi_sql = view_condi + " like '%" + condi + "%'"
	end if
end if

if use_sw = "T" then
	use_sql = " "
  else
	if condi_sql = " " then
		use_sql = " use_sw = '" + use_sw + "'"
	  else
 		use_sql = " and use_sw = '" + use_sw + "'"
	end if
end if

Sql = "SELECT * FROM trade "&where_sql&condi_sql&use_sql&" ORDER BY trade_name ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<style type="text/css">
    <!--
    	.style10 {font-size: 10px; font-family: "����ü", "����ü", Seoul; }
        .style10B {font-size: 10px; font-weight: bold; font-family: "����ü", "����ü", Seoul; }
    -->
    </style>
		<title>��� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr class="style10B">
								<th class="first" scope="col">����</th>
								<th scope="col">�ڵ�</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">�ŷ�ó��(FULL)</th>
								<th scope="col">�ŷ�ó��</th>
								<th scope="col">�ŷ�ó����</th>
								<th scope="col">��ǥ��</th>
								<th scope="col">�ּ�</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">��ȭ</th>
								<th scope="col">�ѽ�</th>
								<th scope="col">�̸���</th>
								<th scope="col">�����</th>
								<th scope="col">�������ȭ</th>
								<th scope="col">�����׷�</th>
								<th scope="col">�׷��</th>
								<th scope="col">����ȸ��</th>
								<th scope="col">��꼭����ŷ�ó�ڵ�</th>
								<th scope="col">��꼭����ŷ�ó��</th>
								<th scope="col">�������</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
							trade_no = mid(rs("trade_no"),1,3) + "-" + mid(rs("trade_no"),4,2) + "-" + mid(rs("trade_no"),6) 
							sql_type="select * from type_code where etc_type='91' and etc_seq ='"+rs("mg_group")+"'"
							set rs_type=dbconn.execute(sql_type)
							mg_group = rs_type("type_name")
							rs_type.Close()		
							if rs("use_sw") = "Y" then
								view_use = "���"
							  else
							  	view_use = "�̻��"
							end if
						%>
							<tr class="style10">
								<td class="first"><%=i%></td>
								<td><%=rs("trade_code")%></td>
								<td><%=trade_no%></td>
								<td><%=rs("trade_full_name")%></td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("trade_id")%></td>
								<td><%=rs("trade_owner")%></td>
								<td><%=rs("trade_addr")%></td>
								<td><%=rs("trade_uptae")%></td>
								<td><%=rs("trade_upjong")%></td>
								<td><%=rs("trade_tel")%></td>
								<td><%=rs("trade_fax")%></td>
								<td><%=rs("trade_email")%></td>
								<td><%=rs("trade_person")%></td>
								<td><%=rs("trade_person_tel")%></td>
								<td><%=mg_group%></td>
								<td><%=rs("group_name")%></td>
								<td><%=rs("support_company")%></td>
								<td><%=rs("bill_trade_code")%></td>
								<td><%=rs("bill_trade_name")%></td>
								<td><%=use_view%></td>
							</tr>
					  	<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

