<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/kwon2010.asp" -->
<%

'ck_sw=Request("ck_sw")
user_id = request.cookies("kwon_user")("coo_id")

'If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	as_type=Request("as_type")
'Else
'	from_date=Request.form("from_date")
'	to_date=Request.form("to_date")
'	as_type=Request.form("as_type")
'End if

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	as_type = "��ü"
End If
mg_group = request.cookies("kwon_user")("coo_mg_group")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_end = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if as_type = "��ü" then
	type_sql = ""
  else
  	type_sql = " (as_type ='"+as_type+"') and "
end if

sql = "select count(*) as err_tot from k1_as_acpt "
sql = sql + "WHERE "+type_sql+" (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "

Rs.Open Sql, Dbconn, 1
err_tot = cint(rs("err_tot"))
if rs.eof then
	err_tot = 0
end if

rs.close()

sql = "select company, count(*) as err_cnt from k1_as_acpt "
sql = sql + " WHERE "+type_sql+" (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
sql = sql + " GROUP BY company ORDER BY company ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="include/kwon_style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/java/PopupCalendar.js"></script>
<title></title>
<style type="text/css">
<!--
.style15 {font-size: 12px}
-->
</style>
</head>

<body>
<table width="900" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" height="30" bgcolor="#6699CC">&nbsp;<span class="style14BW">ȸ�纰 ���� �� ó����Ȳ</span></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellspacing="3">
      <tr>
        <td><form name="form1" method="post" action="k1_waiting.asp?pg_name=k1_company_per.asp">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr valign="middle" class="style12">
              <td width="10%" height="25" bgcolor="#CCCCCC"><div align="center" class="style12">�����Ⱓ</div></td>
              <td width="45%" height="25">&nbsp;                  <input name="from_date" type="text" id="from_date2" size="10" value=<%=from_date%>>                  <input name="button" type="button" onClick="popUpCalendar(this, from_date, 'yyyy-mm-dd')" value="�޷�">                  
                  ~
                  <input name="to_date" type="text" id="to_date2" size="10" value=<%=to_date%>>                  <input name="button2" type="button" class="style12" onClick="popUpCalendar(this, to_date, 'yyyy-mm-dd')" value="�޷�">                    
                  <div align="center" class="style6"></div></td>
              <td width="10%" bgcolor="#CCCCCC"><div align="center">ó������</div></td>
              <td width="25%" height="25" valign="middle" bgcolor="#FFFFFF"><div align="center" class="style6">
                <div align="left"><span class="style12">&nbsp;
                  <select name="as_type" id="as_type">
                      <option value="��ü" <%If as_type = "��ü" then %>selected<% end if %>>��ü</option>
                      <option value="����ó��" <%If as_type = "����ó��" then %>selected<% end if %>>����ó��</option>
                      <option value="�湮ó��" <%If as_type = "�湮ó��" then %>selected<% end if %>>�湮ó��</option>
                      <option value="�űԼ�ġ" <%If as_type = "�űԼ�ġ" then %>selected<% end if %>>�űԼ�ġ</option>
                      <option value="������ġ" <%If as_type = "������ġ" then %>selected<% end if %>>������ġ</option>
                      <option value="������" <%If as_type = "������" then %>selected<% end if %>>������</option>
                      <option value="���ȸ��" <%If as_type = "���ȸ��" then %>selected<% end if %>>���ȸ��</option>
                      <option value="��������" <%If as_type = "��������" then %>selected<% end if %>>��������</option>
                      <option value="��Ÿ" <%If as_type = "��Ÿ" then %>selected<% end if %>>��Ÿ</option>
                  </select>
                </span> </div>
              </div></td>
              <td width="10%" height="25"><div align="center">
                  <input name="imageField" type="image" src="image/burton/view01.gif" width="55" height="20">
              </div></td>
            </tr>
          </table>
          <table width="100%" border="1" cellspacing="0" cellpadding="0">
            <tr valign="middle" bgcolor="#CCFFCC" class="style12">
              <td width="17%" height="25"><div align="center">ȸ �� </div></td>
              <td width="46%" height="25"><div align="center">�� �� �� ( �� ���� ) </div></td>
              <td width="7%" height="25"><div align="center">������</div></td>
              <td width="7%" height="25"><div align="center">�Ϸ�</div></td>
              <td width="7%" height="25"><div align="center">��ó��</div></td>
              <td width="8%" height="25"><div align="center">ó����</div></td>
              <td width="8%"><div align="center">������</div></td>
            </tr>
    	    <% 
			sum_err = 0
			sum_end = 0
			sum_mi = 0
			do until rs.eof 
				err_per = formatnumber((cint(rs("err_cnt"))/err_tot * 100),2)
	
				sql = "select count(*) as end_cnt from k1_as_acpt "
	'			sql = sql + "WHERE "+type_sql+" (mg_group='"+mg_group+"') and (as_process='�Ϸ�' or as_process='��ü' or as_process='���') and (company='"+rs("company")+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
				sql = sql + "WHERE "+type_sql+" (as_process='�Ϸ�' or as_process='���') and (company='"+rs("company")+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
				
				Rs_end.Open Sql, Dbconn, 1
				end_cnt = cint(rs_end("end_cnt"))
				if rs_end.eof then
					end_cnt = 0
				end if
				mi_cnt = cint(rs("err_cnt")) - cint(rs_end("end_cnt"))
				sum_err = sum_err + cint(rs("err_cnt"))
				sum_end = sum_end + cint(rs_end("end_cnt"))
				sum_mi = sum_mi + mi_cnt
				if end_cnt = 0 then
					pro_per = 0 
				  else
				  	pro_per = formatnumber((cint(end_cnt) / cint(rs("err_cnt")) * 100),2)
				end if
			%>
            <tr valign="middle" bgcolor="#FFFFFF" class="style12">
              <td width="17%" height="20" bgcolor="#FFFFCC"><div align="center" class="style12"><%=rs("company")%></div></td>
              <td width="46%" height="10">&nbsp;<img src="image/graph02.gif" width="<%=err_per*97/100%>%" height="13" align="center"></td>
              <td width="7%" class="style12"><div align="right"><%=formatnumber(clng(rs("err_cnt")),0)%>&nbsp;</div></td>
              <td width="7%" class="style12"><div align="right"><%=formatnumber(clng(rs_end("end_cnt")),0)%>&nbsp;</div></td>
              <td width="7%" class="style12"><div align="right"><%=formatnumber(clng(mi_cnt),0)%>&nbsp;</div></td>
              <td width="8%" height="10" class="style12"><div align="right"><%=pro_per%>%&nbsp;</div></td>
              <td width="8%" height="10" class="style12"><div align="right"><%=err_per%>%&nbsp;</div></td>
            </tr>
    		<%
				rs_end.close()
				rs.movenext()
			loop
			rs.close()
			%>
          </table>
          <table width="100%"  border="1" cellpadding="0" cellspacing="0">
            <tr bgcolor="#CCCCCC" class="style12B">
              <td width="17%" height="20" class="style12"><div align="center">��</div></td>
              <td width="46%" height="20" class="style12"><div align="left">&nbsp;</div></td>
              <td width="7%" class="style12"><div align="right"><%=formatnumber(clng(sum_err),0)%>&nbsp;</div></td>
              <td width="7%" class="style12"><div align="right"><%=formatnumber(clng(sum_end),0)%>&nbsp;</div></td>
              <td width="7%" class="style12"><div align="right"><%=formatnumber(clng(sum_mi),0)%>&nbsp;</div></td>
              <td width="8%" class="style12"><div align="right"><%=formatnumber((clng(sum_end)/clng(sum_err)*100),2)%>%&nbsp;</div></td>
              <td width="8%" class="style12">&nbsp;</td>
            </tr>
          </table>
        </form></td>
      </tr>
    </table></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
dbconn.Close()
Set dbconn = Nothing
%>
