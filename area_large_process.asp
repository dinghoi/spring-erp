<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sido_tab
dim acpt_tab(17,2)
dim inst_tab(17,2)
dim ran_tab(17,2)
sido_tab = array("��","����","���","�λ�","�뱸","��õ","����","����","���","����","�泲","���","����","�泲","���","����","����","����")

large_paper_no=Request("large_paper_no")
company=Request("company")
as_type=Request("as_type")
acpt_cnt=Request("acpt_cnt")

for i = 0 to 17
	for j = 1 to 2
		acpt_tab(i,j) = 0
		inst_tab(i,j) = 0
		ran_tab(i,j) = 0
	next
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

' �Ϸ��
sql = "select sido,as_process,count(*) as pro_cnt,sum(dev_inst_cnt) as inst_cnt,sum(ran_cnt) as ran_cnt from as_acpt WHERE large_paper_no ='"&large_paper_no&"' GROUP BY sido, as_process"
Rs.Open Sql, Dbconn, 1

do until rs.eof
	select case rs("sido")
		case "����"
			i = 1
		case "���"
			i = 2
		case "�λ�"
			i = 3
		case "�뱸"
			i = 4
		case "��õ"
			i = 5
		case "����"
			i = 6
		case "����"
			i = 7
		case "���"
			i = 8
		case "����"
			i = 9
		case "�泲"
			i = 10
		case "���"
			i = 11
		case "����"
			i = 12
		case "�泲"
			i = 13
		case "���"
			i = 14
		case "����"
			i = 15
		case "����"
			i = 16
		case "����"
			i = 17
	end select	

	if rs("as_process") = "�Ϸ�" or rs("as_process") = "���" then
		acpt_tab(i,1) = acpt_tab(i,1) + cint(rs("pro_cnt"))
		inst_tab(i,1) = inst_tab(i,1) + cint(rs("inst_cnt"))
		ran_tab(i,1) = ran_tab(i,1) + cint(rs("ran_cnt"))
		acpt_tab(0,1) = acpt_tab(0,1) + cint(rs("pro_cnt"))
		inst_tab(0,1) = inst_tab(0,1) + cint(rs("inst_cnt"))
		ran_tab(0,1) = ran_tab(0,1) + cint(rs("ran_cnt"))
	  else
		acpt_tab(i,2) = acpt_tab(i,2) + cint(rs("pro_cnt"))
		inst_tab(i,2) = inst_tab(i,2) + cint(rs("inst_cnt"))
		ran_tab(i,2) = ran_tab(i,2) + cint(rs("ran_cnt"))
		acpt_tab(0,2) = acpt_tab(0,2) + cint(rs("pro_cnt"))
		inst_tab(0,2) = inst_tab(0,2) + cint(rs("inst_cnt"))
		ran_tab(0,2) = ran_tab(0,2) + cint(rs("ran_cnt"))
	end if
	rs.movenext()
loop
rs.close()

title_line = "������ �뷮�� ó�� ��Ȳ"
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
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="*" >
							<col width="35%" >
							<col width="15%" >
							<col width="35%" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first" scope="col">ȸ��</th>
								<td><%=company%></td>
								<th>������ȣ</th>
								<td><%=large_paper_no%></td>
							</tr>
							<tr>
								<th class="first" scope="col">ó������</th>
								<td><%=as_type%></td>
								<th>�ѰǼ�</th>
								<td><%=formatnumber(acpt_cnt,0)%></td>
							</tr>
						</tbody>
					</table>
					<h3 class="stit">* �õ��� ����</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">�õ�</th>
								<th scope="col" colspan="3" style=" border-bottom:1px solid #e3e3e3;">�����Ǽ�</th>
								<th scope="col" colspan="2" style=" border-bottom:1px solid #e3e3e3;">��ġ����</th>
								<th scope="col" colspan="2" style=" border-bottom:1px solid #e3e3e3;">���������</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">�Ϸ�Ǽ�</th>
								<th scope="col">��ó���Ǽ�</th>
								<th scope="col">��ô��</th>
								<th scope="col">�Ϸ����</th>
								<th scope="col">��ó������</th>
								<th scope="col">�Ϸ����</th>
								<th scope="col">��ó������</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                              <th>��</th>
                              <th class="right"><%=formatnumber(acpt_tab(0,1),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(acpt_tab(0,2),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(acpt_tab(0,1)/(acpt_tab(0,1)+acpt_tab(0,2))*100,2)%>%&nbsp;</th>
                              <th class="right"><%=formatnumber(inst_tab(0,1),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(inst_tab(0,2),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(ran_tab(0,1),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(ran_tab(0,2),0)%>&nbsp;</th>
							</tr>
						<% 	
                    	for i = 1 to 17
                		%>
							<tr>
                              <td><%=sido_tab(i)%></td>
                              <td class="right"><%=formatnumber(acpt_tab(i,1),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(acpt_tab(i,2),0)%>&nbsp;</td>
                              <td class="right">
							<% if acpt_tab(i,1) = 0 and acpt_tab(i,2) = 0 then	%>
                               0.00%
                            <%   else	%>							
							  <%=formatnumber(acpt_tab(i,1)/(acpt_tab(i,1)+acpt_tab(i,2))*100,2)%>%&nbsp;
							<% end if %>
                              </td>
                              <td class="right"><%=formatnumber(inst_tab(i,1),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(inst_tab(i,2),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(ran_tab(i,1),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(ran_tab(i,2),0)%>&nbsp;</td>
							</tr>
                		<% 	
						next 
						%>
						</tbody>
					</table>
				</div>
		</div>				
	</body>
</html>

