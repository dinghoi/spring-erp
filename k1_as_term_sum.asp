<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/kwon2010.asp" -->
<%
dim pro_name(7)
dim pro_cnt(7)
dim err_name
dim company_tab(150)

for i = 0 to 7
	pro_cnt(i) = 0
next

pro_name(0) = "����ó��"
pro_name(1) = "����ó��"
pro_name(2) = "2�� ó��"
pro_name(3) = "3��~ 6��"
pro_name(4) = "7�� �̻�"
pro_name(5) = "ó������"
pro_name(6) = "�԰���"
pro_name(7) = "��ó��"


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_com = Server.CreateObject("ADODB.Recordset")
Set Rs_hol = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect
'ck_sw=Request("ck_sw")

c_name = "��ü"
c_grade = request.cookies("kwon_user")("coo_grade")
c_reside = request.cookies("kwon_user")("coo_reside")
user_id = request.cookies("kwon_user")("coo_id")
mg_group = request.cookies("kwon_user")("coo_mg_group")
user_name = request.cookies("kwon_user")("coo_name")

'If ck_sw = "n" Then
'	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	company = request.form("company")
'  Else
'	from_date=Request("from_date")
'	to_date=Request("to_date")
'	company = "��ü"
'End if

If to_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	company = "��ü"
End If
curr_dd = cstr(datepart("d",to_date))
from_date = mid(to_date,1,8) + "01"

last_year = mid(to_date,1,4) - 1
last_month = mid(to_date,6,2) - 1

curr_year = mid(to_date,1,4)
if last_month = 0 then
	last_month = 12
	curr_year = last_year
end if

curr_month = mid(to_date,6,2)

if	c_grade = "5" and c_reside = "0" then
	c_name = request.cookies("kwon_user")("coo_name")
	company = c_name
end if

if c_name = "��ü" then
	k = 0
	company_tab(0) = "��ü"
	if	( c_grade = "5" and c_reside = "1" ) then
		Sql="select * from k1_etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' and group_name = '"+user_name+"' order by etc_name asc"
		  else
		Sql="select * from k1_etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' order by etc_name asc"
	end if
	Rs_etc.Open Sql, Dbconn, 1
	while not rs_etc.eof
		k = k + 1
		company_tab(k) = rs_etc("etc_name")
		rs_etc.movenext()
	Wend
rs_etc.close()						
end if				

grade_sql = ""
if ( c_grade = "5" and c_reside = "1" ) then
	com_sql = "company = '" + company_tab(1) + "'"	
	for kk = 2 to k
		com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
	next
	grade_sql = " and (" + com_sql + ")"
end if

kkk = k

'��� ó�� ���� (������)
if company = "��ü" then
	if  ( c_grade = "5" and c_reside = "1" ) then
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') " + grade_sql
	  else 
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
	end if
  else
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
		sql = sql + " and company = '" + company + "'"
end if

Rs.Open Sql, Dbconn, 1
acpt_tot = cint(rs("acpt_tot"))
if rs.eof then
	acpt_tot = 0
end if
rs.close()

'���� ó�� ���� (������)
if company = "��ü" then
	if  ( c_grade = "5" and c_reside = "1" ) then
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and month(acpt_date) = "&last_month&" and year(acpt_date) ="&curr_year&grade_sql
	  else 
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and month(acpt_date) = "&last_month&" and year(acpt_date) ="&curr_year
	end if
  else
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and month(acpt_date) = "&last_month&" and year(acpt_date) ="&curr_year
		sql = sql + " and company = '" + company + "'"
end if

Rs.Open Sql, Dbconn, 1

if rs.eof then
	last_tot = 0
  else
 	last_tot =cint(rs("acpt_tot"))
end if
rs.close()

'���� ��� ó�� ���� (������)
if company = "��ü" then
	if  ( c_grade = "5" and c_reside = "1" ) then
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and month(acpt_date) = "&curr_month&" and year(acpt_date) ="&last_year&grade_sql
	  else 
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and month(acpt_date) = "&curr_month&" and year(acpt_date) ="&last_year
	end if
  else
		sql = "select count(*) as acpt_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and month(acpt_date) = "&curr_month&" and year(acpt_date) ="&last_year
		sql = sql + " and company = '" + company + "'"
end if

Rs.Open Sql, Dbconn, 1

if rs.eof then
	last_year = 0
  else
 	last_year =cint(rs("acpt_tot"))
end if
rs.close()

' ��� ó�� �Ϸ��
if company = "��ü" then
	if  ( c_grade = "5" and c_reside = "1" ) then
		sql = "select CAST(acpt_date as date) as acpt_day, CAST((acpt_date + interval 10 DAY_HOUR) as date) as com_date, visit_date, substring(visit_time,1,2) as visit_hh, count(*) as err_cnt from k1_as_acpt "
		sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '��ü' or as_process = '�Ϸ�' or as_process = '���')"
		sql = sql + " and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')" + grade_sql
		sql = sql + " GROUP BY CAST(acpt_date as date), CAST((acpt_date + interval 10 DAY_HOUR) as date), visit_date, substring(visit_time,1,2) Order By visit_date Asc"
	  else 
		sql = "select CAST(acpt_date as date) as acpt_day, CAST((acpt_date + interval 10 DAY_HOUR) as date) as com_date, visit_date, substring(visit_time,1,2) as visit_hh, count(*) as err_cnt from k1_as_acpt "
		sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '��ü' or as_process = '�Ϸ�' or as_process = '���')"
		sql = sql + " and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		sql = sql + " GROUP BY CAST(acpt_date as date), CAST((acpt_date + interval 10 DAY_HOUR) as date), visit_date, substring(visit_time,1,2) Order By visit_date Asc"
	end if
  else
	sql = "select CAST(acpt_date as date) as acpt_day, CAST((acpt_date + interval 10 DAY_HOUR) as date) as com_date, visit_date, substring(visit_time,1,2) as visit_hh, count(*) as err_cnt from k1_as_acpt "
	sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '��ü' or as_process = '�Ϸ�' or as_process = '���') and (company ='" + company + "')"
	sql = sql + " and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
	sql = sql + " GROUP BY CAST(acpt_date as date), CAST((acpt_date + interval 10 DAY_HOUR) as date), visit_date, substring(visit_time,1,2) Order By visit_date Asc"
end if  
Rs.Open Sql, Dbconn, 1

do until rs.eof

  	visit_date = datevalue(rs("visit_date"))
  	visit_day = datevalue(rs("visit_date"))

	if cstr(rs("visit_hh")) > "12" then
		visit_date = dateadd("d",1,visit_date)
	end if
	
	dd = datediff("d", rs("com_date"), visit_date)

	if cstr(visit_day) = cstr(rs("acpt_day")) then
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
			sql_hol = "select * from k1_holiday where holiday = '" + cstr(com_date) + "'"
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
		pro_cnt(d) = pro_cnt(d) + cint(rs("err_cnt"))	
	  else

' ���� ��� ��
		pro_cnt(0) = pro_cnt(0) + cint(rs("err_cnt"))	
	end if
	rs.movenext()
loop
rs.close()
end_tot = pro_cnt(0) + pro_cnt(1) + pro_cnt(2) + pro_cnt(3) + pro_cnt(4)
pro_cnt(7) = acpt_tot - end_tot


'��� ó�� ���� (ó������)
if company = "��ü" then
	if  ( c_grade = "5" and c_reside = "1" ) then
		sql = "select count(*) as end_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '����' or as_process = '����') and (request_date > '"+ to_date +"')" + grade_sql
	  else 
		sql = "select count(*) as end_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '����' or as_process = '����') and (request_date > '"+ to_date +"')"
	end if
  else
		sql = "select count(*) as end_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '����' or as_process = '����') and (request_date > '"+ to_date +"')"
		sql = sql + " and company = '" + company + "'"
end if

Rs.Open Sql, Dbconn, 1
pro_cnt(5) = cint(rs("end_tot"))
pro_cnt(7) = pro_cnt(7) - pro_cnt(5)
if rs.eof then
	pro_cnt(5) = 0
end if
rs.close()

'��� ó�� ���� (�԰�)
if company = "��ü" then
	if  ( c_grade = "5" and c_reside = "1" ) then
		sql = "select count(*) as end_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '�԰�')" + grade_sql
	  else 
		sql = "select count(*) as end_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '�԰�')"
	end if
  else
		sql = "select count(*) as end_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')  and (as_process = '�԰�')"
		sql = sql + " and company = '" + company + "'"
end if

Rs.Open Sql, Dbconn, 1
pro_cnt(6) = cint(rs("end_tot"))
pro_cnt(7) = pro_cnt(7) - pro_cnt(6)
if rs.eof then
	pro_cnt(6) = 0
end if
rs.close()


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="include/kwon_style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/java/PopupCalendar.js"></script>
<title></title>
</head>

<body>
<table width="900" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" height="30" bgcolor="#6699CC">&nbsp;<span class="style14BW">ó�� �Ⱓ�� ������Ȳ</span></td>
  </tr>
  <tr>
    <td height="29"><form name="form1" method="post" action="k1_waiting.asp?pg_name=k1_as_term_sum.asp">
        <table width="100%%"  border="0">
          <tr>
            <td><table width="100%%"  border="0">
                <tr>
                  <td><table width="100%" border="0" cellspacing="2" cellpadding="0">
                    <tr valign="middle" class="style12">
                      <td width="10%" height="20" bgcolor="#CCCCCC"><div align="center" class="style12">�����Ⱓ</div></td>
                      <td width="45%" height="20">&nbsp;
                          <input name="from_date" type="text" id="from_date" value="<%=from_date%>" size="10" readonly="true">
      ~
      <input name="to_date" type="text" id="to_date2" value="<%=to_date%>" size="10">
      <input name="button2" type="button" class="style12" onClick="popUpCalendar(this, to_date, 'yyyy-mm-dd')" value="�޷�">
                      </td>
                      <td width="10%" height="20" bgcolor="#CCCCCC"><div align="center" class="style12">ȸ��</div></td>
                      <td width="25%" height="20" class="style12">&nbsp;
                          <%
		if c_name = "��ü" then
		%>
                          <select name="company" id="company">
                            <% 
					for kk = 0 to kkk
			  	%>
                            <option value='<%=company_tab(kk)%>' <%If company_tab(kk) = company then %>selected<% end if %>><%=company_tab(kk)%></option>
                            <%
					next
				%>
                          </select>
                          <% else %>
                          <%=c_name%>
                          <% end if %>
                      </td>
                      <td width="10%" height="20"><div align="center">
                          <input name="imageField" type="image" src="image/burton/view01.gif" width="55" height="20">
                      </div></td>
                    </tr>
                  </table>
                    <table width="100%" border="1" cellspacing="0" cellpadding="0">
                      <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                        <td width="12%" height="40" rowspan="2"><div align="center" class="style12">��� ���� </div></td>
                        <td height="20" colspan="2"><div align="center" class="style12">����</div></td>
                        <td height="20" colspan="2"><div align="center" class="style12">����</div></td>
                        <td width="12%" height="40" rowspan="2" class="style12"><div align="center" class="style12">ó�� �Ϸ� </div></td>
                        <td width="12%" height="40" rowspan="2" class="style12"><div align="center" class="style12">�� ó �� </div></td>
                        <td width="14%" height="40" rowspan="2" class="style12"><div align="center" class="style12">ó �� �� </div></td>
                      </tr>
                      <tr valign="middle" bgcolor="#CCFFCC">
                        <td width="12%" height="20" class="style12"><div align="center">���� ���� </div></td>
                        <td width="13%" height="20" class="style12"><div align="center">������</div></td>
                        <td width="12%" height="20" class="style12"><div align="center">���� ���� </div></td>
                        <td width="13%" height="20" class="style12"><div align="center">������</div></td>
                      </tr>
                      <tr valign="middle" class="style12">
                        <td width="12%" height="25" bgcolor="#FFFFFF" class="style6"><div align="center" class="style12"><%=formatnumber(clng(acpt_tot),0)%></div></td>
                        <td width="12%" height="25" bgcolor="#FFFFFF" class="style12"><div align="center"><%=formatnumber(clng(last_tot),0)%></div></td>
                        <td width="13%" height="25" bgcolor="#FFFFFF" class="style12"><div align="center">
                            <% if last_tot = 0 then %>
        0%
        <% else %>
        <%=formatnumber(((acpt_tot/last_tot * 100)-100),2)%>%
        <% end if %>
                        </div></td>
                        <td width="12%" height="25" bgcolor="#FFFFFF" class="style12"><div align="center"><%=formatnumber(clng(last_year),0)%></div></td>
                        <td width="13%" height="25" bgcolor="#FFFFFF" class="style6"><div align="center" class="style12">
                            <% if last_year = 0 then %>
        0%
        <% else %>
        <%=formatnumber(((acpt_tot/last_year * 100)-100),2)%>%
        <% end if %>
                        </div></td>
                        <td width="12%" height="25" bgcolor="#FFFFFF" class="style6"><div align="center" class="style12"><%=formatnumber(clng(end_tot),0)%></div></td>
                        <td width="12%" height="25" bgcolor="#FFFFFF" class="style6"><div align="center" class="style12"><%=formatnumber(clng(acpt_tot-end_tot),0)%></div></td>
                        <td width="14%" height="25" class="style6"><div align="center" class="style12">
                            <% if acpt_tot = 0 then %>
        0%
        <% else %>
        <%=formatnumber((end_tot/acpt_tot * 100),2)%>%
        <% end if %>
                        </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td><table width="100%" border="1" cellpadding="0" cellspacing="0">
                    <tr bgcolor="#CCFFCC" class="style12">
                      <td width="12%" height="30"><div align="center" class="style12">ó���Ⱓ</div></td>
                      <td width="62%" height="30"><div align="center">�� �� �� </div></td>
                      <td width="12%" height="30" class="style12"><div align="center">ó���Ǽ�</div></td>
                      <td width="14%" height="30" class="style12"><div align="center">ó����(%)</div>
                          <div align="center"></div></td>
                    </tr>
                  </table>
                    <table width="100%" border="1" cellspacing="0" cellpadding="0">
                      <%
			for i = 0 to 7
			if	acpt_tot = 0 then
				pro_per = 0
			  else
			  	pro_per = formatnumber((pro_cnt(i)/acpt_tot * 100),2)
			end if
			%>
                      <tr>
                        <td width="12%" height="25" bgcolor="#FFFFCC" class="style6"><div align="center" class="style12"><%=pro_name(i)%></div></td>
                        <td width="62%" height="25"><span class="style7">&nbsp;<img src="image/graph02.gif" width="<%=pro_per*97/100%>%" height="13" align="center"></span></td>
                        <td width="12%" height="20" class="style12"><div align="center" class="style12"><%=formatnumber(clng(pro_cnt(i)),0)%></div></td>
                        <td width="14%" height="20" class="style12"><div align="center" class="style12"><%=pro_per%>%</div></td>
                      </tr>
                      <%
			next
			%>
                    </table>
                    <table width="100%%"  border="1" cellpadding="0" cellspacing="0">
                      <tr bgcolor="#CCCCCC" class="style12B">
                        <td width="12%" height="25"><div align="center">�Ѱ�</div></td>
                        <td width="62%"><div align="center">&nbsp;</div></td>
                        <td width="12%"><div align="center"><%=formatnumber(clng(acpt_tot),0)%></div></td>
                        <td width="14%"><div align="center">&nbsp;</div></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
</body>
</html>
<%
dbconn.Close()
Set dbconn = Nothing
%>
