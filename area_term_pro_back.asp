<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_tab
dim com_sum(16)
dim ok_sum(16)
dim mi_sum(16)
dim com_cnt(16,9)
dim com_in(16,9)
dim sum_cnt(9)
dim sum_in(9)
dim company_tab(150)
dim end_tab(11)
dim mi_tab(11)
dim curr_mi_tab(11)
dim mi_in
com_tab = array("����","���","�λ�","�뱸","��õ","����","����","���","����","�泲","���","�泲","���","����","����","����","����")

from_date=Request.form("from_date")
to_date=Request.form("to_date")
as_type=Request.form("as_type")
company=Request.form("company")

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	as_type = "�湮ó��"
	company = "��ü"
End If

for i = 0 to 16
	com_sum(i) = 0
	ok_sum(i) = 0
	mi_sum(i) = 0
	for j = 0 to 9
		com_cnt(i,j) = 0
		com_in(i,j) = 0
		sum_cnt(j) = 0
		sum_in(j) = 0
	next
next
for i = 0 to 11
	end_tab(i) = 0
	mi_tab(i) = 0
	curr_mi_tab(i) = 0
next

curr_day = datevalue(mid(cstr(now()),1,10))
curr_date = datevalue(mid(dateadd("h",12,now()),1,10))

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

if company = "��ü" then
	com_sql0 = ""
	com_sql = ""
  else
  	com_sql0 = " (company ='"+company+"') and "
  	com_sql = " (as_acpt.company ='"+company+"') and "
end if
if as_type = "��ü" then
	type_sql0 = ""
	type_sql = ""
  else
  	type_sql0 = " (as_type ='"+as_type+"') and "
  	type_sql = " (as_acpt.as_type ='"+as_type+"') and "
end if

tot_cnt = 0
' ó���Ϸ�
sql = "select as_type, count(*) as end_cnt from as_acpt"
sql = sql + " where "+com_sql0+" (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"') and (as_process = '��ü' or as_process = '�Ϸ�' or as_process = '���') "
sql = sql + " GROUP BY as_type Order By as_type Asc"
Rs.Open Sql, Dbconn, 1 

do until rs.eof
	end_cnt = cint(rs("end_cnt"))
	end_tab(0) = end_tab(0) + end_cnt
	
	if rs("as_type") = "����ó��" then
		end_tab(1) = end_tab(1) + end_cnt
	end if
	if rs("as_type") = "�湮ó��" then
		end_tab(2) = end_tab(2) + end_cnt
	end if
	if rs("as_type") = "�űԼ�ġ" then
		end_tab(3) = end_tab(3) + end_cnt
	end if
	if rs("as_type") = "�űԼ�ġ����" then
		end_tab(4) = end_tab(4) + end_cnt
	end if
	if rs("as_type") = "������ġ" then
		end_tab(5) = end_tab(5) + end_cnt
	end if
	if rs("as_type") = "������ġ����" then
		end_tab(6) = end_tab(6) + end_cnt
	end if
	if rs("as_type") = "������" then
		end_tab(7) = end_tab(7) + end_cnt
	end if
	if rs("as_type") = "����������" then
		end_tab(8) = end_tab(8) + end_cnt
	end if
	if rs("as_type") = "���ȸ��" then
		end_tab(9) = end_tab(9) + end_cnt
	end if
	if rs("as_type") = "��������" then
		end_tab(10) = end_tab(10) + end_cnt
	end if
	if rs("as_type") = "��Ÿ" then
		end_tab(11) = end_tab(11) + end_cnt
	end if

	rs.movenext()
loop
rs.close()

'������� ��ó��
sql = "select as_type, count(*) as end_cnt from as_acpt"
sql = sql + " where "+com_sql0+" (Cast(acpt_date as date) <= '"+to_date+"') and (as_process = '����' or as_process = '�԰�' or as_process = '����') "
sql = sql + " GROUP BY as_type Order By as_type Asc"
Rs.Open Sql, Dbconn, 1 

do until rs.eof
	end_cnt = cint(rs("end_cnt"))
	curr_mi_tab(0) = curr_mi_tab(0) + end_cnt
	
	if rs("as_type") = "����ó��" then
		curr_mi_tab(1) = curr_mi_tab(1) + end_cnt
	end if
	if rs("as_type") = "�湮ó��" then
		curr_mi_tab(2) = curr_mi_tab(2) + end_cnt
	end if
	if rs("as_type") = "�űԼ�ġ" then
		curr_mi_tab(3) = curr_mi_tab(3) + end_cnt
	end if
	if rs("as_type") = "�űԼ�ġ����" then
		curr_mi_tab(4) = curr_mi_tab(4) + end_cnt
	end if
	if rs("as_type") = "������ġ" then
		curr_mi_tab(5) = curr_mi_tab(5) + end_cnt
	end if
	if rs("as_type") = "������ġ����" then
		curr_mi_tab(6) = curr_mi_tab(6) + end_cnt
	end if
	if rs("as_type") = "������" then
		curr_mi_tab(7) = curr_mi_tab(7) + end_cnt
	end if
	if rs("as_type") = "����������" then
		curr_mi_tab(8) = curr_mi_tab(8) + end_cnt
	end if
	if rs("as_type") = "���ȸ��" then
		curr_mi_tab(9) = curr_mi_tab(9) + end_cnt
	end if
	if rs("as_type") = "��������" then
		curr_mi_tab(10) = curr_mi_tab(10) + end_cnt
	end if
	if rs("as_type") = "��Ÿ" then
		curr_mi_tab(11) = curr_mi_tab(11) + end_cnt
	end if

	rs.movenext()
loop
rs.close()
'�Ⱓ���԰��
sql = "select as_type, count(*) as end_cnt from as_acpt"
sql = sql + " where "+com_sql0+" (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"') and (as_process = '�԰�') "
sql = sql + " GROUP BY as_type Order By as_type Asc"
Rs.Open Sql, Dbconn, 1 

mi_in = 0
do until rs.eof
	end_cnt = cint(rs("end_cnt"))
	mi_in = mi_in + end_cnt	
	rs.movenext()
loop
rs.close()
'������� �԰��
sql = "select as_type, count(*) as end_cnt from as_acpt"
sql = sql + " where "+com_sql0+" (as_process = '�԰�') "
sql = sql + " GROUP BY as_type Order By as_type Asc"
Rs.Open Sql, Dbconn, 1 

curr_mi_in = 0
do until rs.eof
	end_cnt = cint(rs("end_cnt"))
	curr_mi_in = curr_mi_in + end_cnt	
	rs.movenext()
loop
rs.close()

' �Ѱ� SQL -> ��ü SUM ���� ����
'sql = "select count(*) as err_cnt from as_acpt"
''sql = sql + " WHERE (mg_group='"+mg_group+"') and (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"') and (reside_place = '�ݼ���') and (as_type <> '����ó��')"
''sql = sql + " WHERE (mg_group='"+mg_group+"') and (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"') and (as_type = '�湮ó��')"
'sql = sql + " WHERE "+com_sql0+type_sql0+"(mg_group='"+mg_group+"') and (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"')"
'Rs.Open Sql, Dbconn, 1

'if rs.eof then
'	tot_cnt = 0
'  else
'  	tot_cnt = cint(rs("err_cnt"))
'end if
'rs.close()

' �Ϸ��
sql = "select as_acpt.sido, Cast(acpt_date as date) as acpt_day, CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) as com_date, visit_date, substring(visit_time,1,2) as visit_hh, count(*) as err_cnt from as_acpt"
'sql = sql + " WHERE (as_acpt.mg_group='"+mg_group+"') and (k1_etc_code.etc_type = '81') and (as_acpt.as_process = '���' or as_acpt.as_process = '�Ϸ�') and (as_acpt.as_type <> '����ó��') and (reside_place = '�ݼ���')"
sql = sql + " WHERE "+com_sql+type_sql+" (as_acpt.as_process = '��ü' or as_acpt.as_process = '�Ϸ�' or as_acpt.as_process = '���')"
sql = sql + " and (Cast(acpt_date as date) >= '" + from_date + "' AND Cast(acpt_date as date) <= '"+to_date+"')"
sql = sql + " GROUP BY as_acpt.sido, Cast(acpt_date as date), CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date), visit_date, substring(visit_time,1,2) Order By as_acpt.sido Asc"
Rs.Open Sql, Dbconn, 1

do until rs.eof
	select case rs("sido")
		case "����"
			i = 0
		case "���"
			i = 1
		case "�λ�"
			i = 2
		case "�뱸"
			i = 3
		case "��õ"
			i = 4
		case "����"
			i = 5
		case "����"
			i = 6
		case "���"
			i = 7
		case "����"
			i = 8
		case "�泲"
			i = 9
		case "���"
			i = 10
		case "�泲"
			i = 11
		case "���"
			i = 12
		case "����"
			i = 13
		case "����"
			i = 14
		case "����"
			i = 15
		case "����"
			i = 16
	end select	

  	visit_date = datevalue(rs("visit_date"))
' 1/19 �߰�
  	visit_day = datevalue(rs("visit_date"))
' 1/19 �߰� end

	if cstr(rs("visit_hh")) > "12" then
		visit_date = dateadd("d",1,visit_date)
	end if
	
	dd = datediff("d", rs("com_date"), visit_date)

	if cstr(rs("visit_date")) = cstr(rs("acpt_day")) then
		dd = 0
	end if

	if dd < 0 then
		dd = 0 
	end if

'���� ���
	if dd > 0 then
		a = datediff("d", rs("acpt_day"), visit_day)
		b = datepart("w",rs("acpt_day"))
		c = a + b
		d = a
		if a > 1 then
			if c > 7 then
				d = a - 2
			end if
		end if
		
'		visit_date = rs("visit_date")
		com_date = datevalue(rs("acpt_day"))
	
		do until com_date > visit_day
			sql_hol = "select * from holiday where holiday = '" + cstr(com_date) + "'"
			Set rs_hol=DbConn.Execute(SQL_hol)
			if rs_hol.eof or rs_hol.bof then
				d = d
			  else 
				d = d -1
			end if
			com_date = dateadd("d",1,com_date)
			rs_hol.close()
		loop
' 2012-02-06
		if d = 1 then
			visit_hh = int(rs("visit_hh"))
			if rs("acpt_day") <> rs("com_date") and visit_hh < 12 then
				d = 0
			end if
		end if
' 2012-02-06 end
		if d > 2 and d < 7 then
			d = 3
		end if
		if d > 6 then
			d = 4
		end if
' 1/19 �߰�
'		ddd = datediff("d", rs("acpt_day"), visit_day)
'		if d > ddd then
'			d = ddd
'		end if
' 1/19 �߰� end
		com_cnt(i,d) = com_cnt(i,d) + cint(rs("err_cnt"))	
	  else

' ���� ��� ��
		com_cnt(i,0) = com_cnt(i,0) + cint(rs("err_cnt"))
	end if
	tot_cnt = tot_cnt + cint(rs("err_cnt"))
	rs.movenext()
loop
rs.close()

' ��ó����
sql = "select as_acpt.sido, as_acpt.as_process, Cast(acpt_date as date) as acpt_day, CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) as com_date, count(*) as err_cnt from as_acpt"
sql = sql + " WHERE "+com_sql+type_sql+" (as_acpt.as_process = '����' or as_acpt.as_process = '�԰�' or as_acpt.as_process = '����')"
sql = sql + " and (Cast(acpt_date as date) <= '"+to_date+"')"
sql = sql + " GROUP BY as_acpt.sido, as_acpt.as_process, Cast(acpt_date as date), CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) Order By as_acpt.sido Asc"
Rs.Open Sql, Dbconn, 1

do until rs.eof
'	i = int(rs("etc_code")) - 8101
'	com_tab(i) = rs("sido")
	select case rs("sido")
		case "����"
			i = 0
		case "���"
			i = 1
		case "�λ�"
			i = 2
		case "�뱸"
			i = 3
		case "��õ"
			i = 4
		case "����"
			i = 5
		case "����"
			i = 6
		case "���"
			i = 7
		case "����"
			i = 8
		case "�泲"
			i = 9
		case "���"
			i = 10
		case "�泲"
			i = 11
		case "���"
			i = 12
		case "����"
			i = 13
		case "����"
			i = 14
		case "����"
			i = 15
		case "����"
			i = 16
	end select	

	dd = datediff("d", rs("com_date"), curr_date)

	if dd < 0 then
		dd = 0 
	end if
	
	if cstr(curr_day) = cstr(rs("acpt_day")) then
		dd = 0
	end if

'���� ���
	if dd > 0 then
		a = datediff("d", rs("acpt_day"), curr_day)
		b = datepart("w",rs("acpt_day"))
		bb = datepart("w", curr_day)
		if bb = 1 then
			a = a -1
		end if
		c = a + b
		d = a
		if a > 1 then
			if c > 7 then
				d = a - 2
			end if
		end if
		
'		visit_date = rs("visit_date")
		com_date = datevalue(rs("acpt_day"))
'		act_date = com_date
	
		do until com_date > curr_day
			sql_hol = "select * from holiday where holiday = '" + cstr(com_date) + "'"
			Set rs_hol=DbConn.Execute(SQL_hol)
			if rs_hol.eof or rs_hol.bof then
				d = d
			  else 
				d = d -1
			end if
			com_date = dateadd("d",1,com_date)
			rs_hol.close()
		loop
' 2012-02-06
		if d = 1 then
			curr_hh = int(datepart("h",now()))
			if rs("acpt_day") <> rs("com_date") and curr_hh < 12 then
				d = 0
			end if
		end if
' 2012-02-06 end
		if d = 0 then
			j = 5
		  elseif d = 1 then
			j = 6
		  elseif d = 2 then
			j = 7
		  elseif d > 2 and d < 7  then
			j = 8
		  else
			j = 9
		end if
		com_cnt(i,j) = com_cnt(i,j) + cint(rs("err_cnt"))	

		if rs("as_process") = "�԰�" then		
			com_in(i,j) = com_in(i,j) + cint(rs("err_cnt"))
		end if
	  else
' ���� ��� ��
		com_cnt(i,5) = com_cnt(i,5) + cint(rs("err_cnt"))

		if rs("as_process") = "�԰�" then		
			com_in(i,5) = com_in(i,5) + cint(rs("err_cnt"))
		end if
	end if
	tot_cnt = tot_cnt + cint(rs("err_cnt"))
	rs.movenext()
loop
rs.close()

title_line = "������ ó�� �Ⱓ�� ��Ȳ(������ ����)"
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
				return "3 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.from_date.value > document.frm.to_date.value) {
					alert ("�������� �����Ϻ��� Ŭ���� �����ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/sum_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=area_term_pro.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								<strong>������ : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>������ : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<strong>ȸ��</strong>
								<strong>ȸ��</strong>
							  	<%
									sql="select * from trade where use_sw = 'Y' and mg_group = '" + mg_group + "' order by trade_name asc"
                                    rs_trade.Open Sql, Dbconn, 1
                                %>
								<label>
        						<select name="company" id="company">
									<option value="��ü">��ü</option>
          					<% 
								While not rs_trade.eof 
							%>
          							<option value='<%=rs_trade("trade_name")%>' <%If rs_trade("trade_name") = company  then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
          					<%
									rs_trade.movenext()  
								Wend 
								rs_trade.Close()
							%>
        						</select>
                                </label>
								<strong>ó������</strong>
                                <select name="as_type" id="as_type" style="width:100px">
                                    <option value="��ü" <%If as_type = "��ü" then %>selected<% end if %>>��ü</option>
                                    <option value="����ó��" <%If as_type = "����ó��" then %>selected<% end if %>>����ó��</option>
                                    <option value="�湮ó��" <%If as_type = "�湮ó��" then %>selected<% end if %>>�湮ó��</option>
                                    <option value="�űԼ�ġ" <%If as_type = "�űԼ�ġ" then %>selected<% end if %>>�űԼ�ġ</option>
                                    <option value="�űԼ�ġ����" <%If as_type = "�űԼ�ġ����" then %>selected<% end if %>>�űԼ�ġ����</option>
                                    <option value="������ġ" <%If as_type = "������ġ" then %>selected<% end if %>>������ġ</option>
                                    <option value="������ġ����" <%If as_type = "������ġ����" then %>selected<% end if %>>������ġ����</option>
                                    <option value="������" <%If as_type = "������" then %>selected<% end if %>>������</option>
                                    <option value="����������" <%If as_type = "����������" then %>selected<% end if %>>����������</option>
                                    <option value="���ȸ��" <%If as_type = "���ȸ��" then %>selected<% end if %>>���ȸ��</option>
                                    <option value="��������" <%If as_type = "��������" then %>selected<% end if %>>��������</option>
                                    <option value="��Ÿ" <%If as_type = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                                </select>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">ó������</th>
								<th scope="col">�հ�</th>
								<th scope="col">����ó��</th>
								<th scope="col">�湮ó��</th>
								<th scope="col">�űԼ�ġ</th>
								<th scope="col">�űԼ�ġ.����</th>
								<th scope="col">������ġ</th>
								<th scope="col">������ġ.����</th>
								<th scope="col">������</th>
								<th scope="col">����������</th>
								<th scope="col">���ȸ��</th>
								<th scope="col">��������</th>
								<th scope="col">��Ÿ</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                                <td class="first">ó���Ϸ�</td>
                                <td><%=formatnumber(end_tab(0),0)%></td>
                                <td><%=formatnumber(end_tab(1),0)%></td>
                                <td><%=formatnumber(end_tab(2),0)%></td>
                                <td><%=formatnumber(end_tab(3),0)%></td>
                                <td><%=formatnumber(end_tab(4),0)%></td>
                                <td><%=formatnumber(end_tab(5),0)%></td>
                                <td><%=formatnumber(end_tab(6),0)%></td>
                                <td><%=formatnumber(end_tab(7),0)%></td>
                                <td><%=formatnumber(end_tab(8),0)%></td>
                                <td><%=formatnumber(end_tab(9),0)%></td>
                                <td><%=formatnumber(end_tab(10),0)%></td>
                                <td><%=formatnumber(end_tab(11),0)%></td>
							</tr>
							<tr>
                                <td class="first">��ü��ó��</td>
                                <td><%=formatnumber(curr_mi_tab(0),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(1),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(2),0)%>&nbsp;(<%=curr_mi_in%>)</td>
                                <td><%=formatnumber(curr_mi_tab(3),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(4),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(5),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(6),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(7),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(8),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(9),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(10),0)%></td>
                                <td><%=formatnumber(curr_mi_tab(11),0)%></td>
							</tr>
						</tbody>
					</table>
					<h3 class="stit">* �õ��� ����</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">�õ�</th>
								<th scope="col" colspan="6" style=" border-bottom:1px solid #e3e3e3;">ó���Ϸ�</th>
								<th scope="col" colspan="6" style=" border-bottom:1px solid #e3e3e3;">��ó�� * ��ȣ�� �԰��</th>
								<th scope="col" rowspan="2">�õ���</th>
								<th scope="col" rowspan="2">�����</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">����</th>
								<th scope="col">����</th>
								<th scope="col">2��</th>
								<th scope="col">3��~6��</th>
								<th scope="col">7���̻�</th>
								<th scope="col">�Ұ�</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">2��</th>
								<th scope="col">3��~6��</th>
								<th scope="col">7���̻�</th>
								<th scope="col">�Ұ�</th>
							</tr>
						</thead>
						<tbody>
						<% 	
                    	if tot_cnt > 0 then
                        	k = 0
                      	  else
                        	k = 16
                    	end if
        
                    	for i = k to 15 
                        	if	com_tab(i) <> "" then
        
								for j = 0 to 4
									ok_sum(i) = ok_sum(i) + com_cnt(i,j)
									sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)				
								next
								for j = 5 to 9
									mi_sum(i) = mi_sum(i) + com_cnt(i,j)
									sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)				
									sum_in(j) = sum_in(j) + com_in(i,j)				
								next
								com_sum(i) = ok_sum(i) + mi_sum(i)
				
								sido = com_tab(i)
							end if
						next
                		%>
							<tr>
                              <th>��</th>
                              <th><%=formatnumber(clng(sum_cnt(0)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(1)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(2)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(3)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(4)),0)%>&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)),0)%>&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%="�Ѱ�"%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=0%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(5)),0)%></a>(<%=sum_in(5)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%="�Ѱ�"%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(6)),0)%></a>(<%=sum_in(6)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%="�Ѱ�"%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(7)),0)%></a>(<%=sum_in(7)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%="�Ѱ�"%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(8)),0)%></a>(<%=sum_in(8)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%="�Ѱ�"%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=7%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(9)),0)%></a>(<%=sum_in(9)%>)&nbsp;</th>
                              <th><a  href="#" onClick="pop_Window('as_michulri_popup.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%="�Ѱ�"%>&company=<%=company%>&as_type=<%=as_type%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)),0)%>(<%=sum_in(5)+sum_in(6)+sum_in(7)+sum_in(8)+sum_in(9)%>)&nbsp;</th>
                              <th><%=formatnumber(clng(sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)),0)%>&nbsp;</th>
                              <th>
                              <% if tot_cnt = 0 then %>
                                    0%
                                <% else %>
                                    <%=formatnumber(((sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9))/tot_cnt * 100),2)%>%
                                <% end if %>
                              &nbsp;
                              </th>
							</tr>
						<% 	
                    	if tot_cnt > 0 then
                        	k = 0
                      	  else
                        	k = 16
                    	end if
        
                    	for i = k to 15 
                        	if	com_tab(i) <> "" then
                		%>
							<tr>
                              <td><%=com_tab(i)%></td>
                              <td><%=formatnumber(clng(com_cnt(i,0)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(com_cnt(i,1)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(com_cnt(i,2)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(com_cnt(i,3)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(com_cnt(i,4)),0)%>&nbsp;</td>
                              <td><%=formatnumber(clng(ok_sum(i)),0)%>&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=0%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,5)),0)%></a>(<%=com_in(i,5)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,6)),0)%></a>(<%=com_in(i,6)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,7)),0)%></a>(<%=com_in(i,7)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,8)),0)%></a>(<%=com_in(i,8)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('day_michulri.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&days=<%=7%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(com_cnt(i,9)),0)%></a>(<%=com_in(i,9)%>)&nbsp;</td>
                              <td><a  href="#" onClick="pop_Window('as_michulri_popup.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=formatnumber(clng(mi_sum(i)),0)%></a>(<%=com_in(i,5)+com_in(i,6)+com_in(i,7)+com_in(i,8)+com_in(i,9)%>)&nbsp;</td>
                              <td><%=formatnumber(clng(com_sum(i)),0)%>&nbsp;</td>
                              <td>
                              <% if tot_cnt = 0 then %>
                                    0%
                                <% else %>
                                    <%=formatnumber((com_sum(i)/tot_cnt * 100),2)%>%
                                <% end if %>
                              &nbsp;
                              </td>
							</tr>
                		<% 	
							end if
						next 
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

