<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

to_date = Request("to_date")

savefilename = "CE�� ������ ��ó�� ��Ȳ" + to_date + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select memb.user_id,memb.team,memb.user_name,memb.reside,memb.reside_place from as_acpt inner join memb on as_acpt.mg_ce_id "
sql = sql + "= memb.user_id Where (as_acpt.as_process='����' or as_acpt.as_process='����' or as_acpt.as_process='�԰�')"
sql = sql + " GROUP BY memb.user_id,memb.team,memb.user_name,memb.reside,memb.reside_place Order By memb.team, memb.user_name Asc"

Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title></title>
<style type="text/css">
<!--
.style5 {font-size: 12}
.style6 {
	font-size: 12px;
	font-family: "����ü", "����ü", Seoul;
}
.style7 {font-size: 12px}
.style8 {font-family: "����ü", "����ü", Seoul}
-->
</style>
</head>

<body>
<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%"  border="0" cellspacing="3">
      <tr>
        <td>
          <table width="800" border="0" cellspacing="0" cellpadding="0">
            <tr valign="middle" class="style6">
              <td width="100" height="25" bgcolor="#CCCCCC"><div align="center" class="style6">������</div></td>
              <td height="25">&nbsp;<%=to_date%></td>
              </tr>
          </table>
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="64"><table width="100%" border="1" cellspacing="0" cellpadding="0">
      <tr valign="middle" bgcolor="#CCFFCC" class="style6">
        <td width="75" rowspan="2" bgcolor="#FFFF99"><div align="center" class="style6">�Ҽ�</div></td>
        <td width="55" rowspan="2" bgcolor="#FFFF99"><div align="center">CE</div></td>
        <td rowspan="2" bgcolor="#FFFF99"><div align="center">����</div></td>
        <td height="20" colspan="13"><div align="center">���ϱ��� ��ó��</div></td>
        <td height="20" colspan="13" bgcolor="#FFCCFF"><div align="center">��ü ��ó��</div></td>
        </tr>
      <tr valign="middle" bgcolor="#CCFFCC" class="style6">
        <td width="30" height="20"><div align="center" class="style6">��</div></td>
        <td width="30" height="20"><div align="center">����</div></td>
        <td width="30" height="20"><div align="center">�湮</div></td>
        <td width="30" height="20"><div align="center">�԰�</div></td>
        <td width="30" height="20"><div align="center">�űԼ�ġ</div></td>
        <td width="30" height="20"><div align="center">�ż�����</div></td>
        <td width="30" height="20"><div align="center">������ġ</div></td>
        <td width="30" height="20"><div align="center">�̼�����</div></td>
        <td width="30" height="20"><div align="center">��</div></td>
        <td width="30" height="20"><div align="center">������</div></td>
        <td width="30" height="20"><div align="center">ȸ��</div></td>
        <td width="30" height="20"><div align="center">����</div></td>
        <td width="30" height="20"><div align="center">��Ÿ</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center" class="style6">��</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">����</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">�湮</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">�԰�</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">�űԼ�ġ</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">�ż�����</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">������ġ</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">�̼�����</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">������</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">��</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">ȸ��</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">����</div></td>
        <td width="30" height="20" bgcolor="#FFCCFF"><div align="center">��Ÿ</div></td>
      </tr>
<% 
		dim day_sum(12)
		dim month_sum(12)
		dim day_tot(12)
		dim month_tot(12)
		for i = 0 to 12
			day_sum(i) = 0
			month_sum(i) = 0
			day_tot(i) = 0
			month_tot(i) = 0
		next

		do until rs.eof 
' ���� ��ó�� �԰�
			sql = "select count(*) as end_cnt from as_acpt "
			sql = sql + "WHERE (as_process='�԰�') and (mg_ce_id='"+rs("user_id")+"')"
			set rs_in=dbconn.execute(sql)
			if rs_in.eof then
				month_sum(3) = 0
			  else
			  	month_sum(3) = cint(rs_in("end_cnt"))
			end if
			rs_in.close()

' ���� ��ó�� �԰�
			sql = "select count(*) as end_cnt from as_acpt "
			sql = sql + "WHERE (as_process='�԰�') and (mg_ce_id='"+rs("user_id")+"') and (request_date <= '"+to_date+"')"
			set rs_in=dbconn.execute(sql)
			if rs_in.eof then
				day_sum(3) = 0
			  else
			  	day_sum(3) = cint(rs_in("end_cnt"))
			end if
			rs_in.close()
' ���� ������ ��ó��
			sql = "select as_type, count(*) as end_cnt from as_acpt "
			sql = sql + "WHERE (as_process='����' or as_process='����') and (mg_ce_id='"+rs("user_id")+"') GROUP BY as_type"		
			rs_as.Open Sql, Dbconn, 1
			do until rs_as.eof
				select case rs_as("as_type")
                	case "����ó��"
                    	month_sum(1) = cint(rs_as("end_cnt"))	
                    case "�湮ó��"
                        month_sum(2) = cint(rs_as("end_cnt"))	
                    case "�űԼ�ġ"
                        month_sum(4) = cint(rs_as("end_cnt"))	
                    case "�űԼ�ġ����"
                        month_sum(5) = cint(rs_as("end_cnt"))	
                    case "������ġ"
                        month_sum(6) = cint(rs_as("end_cnt"))	
                    case "������ġ����"
                        month_sum(7) = cint(rs_as("end_cnt"))	
                    case "������"
                        month_sum(8) = cint(rs_as("end_cnt"))	
                    case "����������"
                        month_sum(9) = cint(rs_as("end_cnt"))	
                    case "���ȸ��"
                        month_sum(10) = cint(rs_as("end_cnt"))	
                    case "��������"
                        month_sum(11) = cint(rs_as("end_cnt"))	
                    case "��Ÿ"
                        month_sum(12) = cint(rs_as("end_cnt"))	
				end select												
				rs_as.movenext()
			loop
			rs_as.close()
' ���� ������ ��ó��
			sql = "select as_type, count(*) as end_cnt from as_acpt "
			sql = sql + "WHERE (as_process='����' or as_process='����') and (mg_ce_id='"+rs("user_id")+"') and (request_date <= '"+to_date+"') GROUP BY as_type"		
			rs_as.Open Sql, Dbconn, 1
			do until rs_as.eof
				select case rs_as("as_type")
                	case "����ó��"
                    	day_sum(1) = cint(rs_as("end_cnt"))	
                    case "�湮ó��"
                        day_sum(2) = cint(rs_as("end_cnt"))	
                    case "�űԼ�ġ"
                        day_sum(4) = cint(rs_as("end_cnt"))	
                    case "�űԼ�ġ����"
                        day_sum(5) = cint(rs_as("end_cnt"))	
                    case "������ġ"
                        day_sum(6) = cint(rs_as("end_cnt"))	
                    case "������ġ����"
                        day_sum(7) = cint(rs_as("end_cnt"))	
                    case "������"
                        day_sum(8) = cint(rs_as("end_cnt"))	
                    case "����������"
                        day_sum(9) = cint(rs_as("end_cnt"))	
                    case "���ȸ��"
                        day_sum(10) = cint(rs_as("end_cnt"))	
                    case "��������"
                        day_sum(11) = cint(rs_as("end_cnt"))	
                    case "��Ÿ"
                        day_sum(12) = cint(rs_as("end_cnt"))	
				end select												
				rs_as.movenext()
			loop
			rs_as.close()

			for i = 1 to 12
				day_sum(0) = day_sum(0) + day_sum(i)
				month_sum(0) = month_sum(0) + month_sum(i)
				day_tot(0) = day_tot(0) + day_tot(i)
				month_tot(0) = month_tot(0) + month_tot(i)			
			next
			for i = 1 to 12
				day_tot(i) = day_tot(i) + day_sum(i)
				month_tot(i) = month_tot(i) + month_sum(i)			
			next

			if day_sum(0) <> 0 or month_sum(0) <> 0 then
	%>
	      <tr class="style6">
        <td height="20"><div align="center"><%=rs("team")%></div></td>
        <td height="20"><div align="center"><%=rs("user_name")%></div></td>
        <td height="20"><div align="center"><%=rs("reside_place")%></div></td>
        <td bgcolor="#CCFFCC"><div align="center"><%=day_sum(0)%></div></td>
        <td><div align="center"><%=day_sum(1)%></div></td>
        <td><div align="center"><%=day_sum(2)%></div></td>
        <td><div align="center"><%=day_sum(3)%></div></td>
        <td><div align="center"><%=day_sum(4)%></div></td>
        <td><div align="center"><%=day_sum(5)%></div></td>
        <td><div align="center"><%=day_sum(6)%></div></td>
        <td><div align="center"><%=day_sum(7)%></div></td>
        <td><div align="center"><%=day_sum(8)%></div></td>
        <td><div align="center"><%=day_sum(9)%></div></td>
        <td><div align="center"><%=day_sum(10)%></div></td>
        <td><div align="center"><%=day_sum(11)%></div></td>
        <td><div align="center"><%=day_sum(12)%></div></td>
        <td bgcolor="#FFCCFF"><div align="center"><%=month_sum(0)%></div></td>
        <td><div align="center"><%=month_sum(1)%></div></td>
        <td><div align="center"><%=month_sum(2)%></div></td>
        <td><div align="center"><%=month_sum(3)%></div></td>
        <td><div align="center"><%=month_sum(4)%></div></td>
        <td><div align="center"><%=month_sum(5)%></div></td>
        <td><div align="center"><%=month_sum(6)%></div></td>
        <td><div align="center"><%=month_sum(7)%></div></td>
        <td><div align="center"><%=month_sum(8)%></div></td>
        <td><div align="center"><%=month_sum(9)%></div></td>
        <td><div align="center"><%=month_sum(10)%></div></td>
        <td><div align="center"><%=month_sum(11)%></div></td>
        <td><div align="center"><%=month_sum(12)%></div></td>
      </tr>
		<%
			end if
			
			for i = 0 to 12
				day_sum(i) = 0
				month_sum(i) = 0
			next

			rs.movenext()
		loop
		rs.close()
		day_tot(0) = day_tot(1) + day_tot(2) + day_tot(3) + day_tot(4) + day_tot(5) + day_tot(6) + day_tot(7) + day_tot(8) + day_tot(9) + day_tot(10) + day_tot(11) + day_tot(12)
		month_tot(0) = month_tot(1) + month_tot(2) + month_tot(3) + month_tot(4) + month_tot(5) + month_tot(6) + month_tot(7) + month_tot(8) + month_tot(9) + month_tot(10) + month_tot(11) + month_tot(12)
		%>
      <tr valign="middle" bgcolor="#FFFFFF" class="style6">
        <td height="20" colspan="3" bgcolor="#CCCCCC"><div align="center">��</div></td>
        <td height="20" bgcolor="#CCFFCC"><div align="center"><%=day_tot(0)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(1)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(2)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(3)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(4)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(5)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(6)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(7)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(8)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(9)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(10)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(11)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=day_tot(12)%></div></td>
        <td height="20" bgcolor="#FFCCFF"><div align="center"><%=month_tot(0)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(1)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(2)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(3)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(4)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(5)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(6)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(7)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(8)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(9)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(10)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(11)%></div></td>
        <td height="20" bgcolor="#CCCCCC"><div align="center"><%=month_tot(12)%></div></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
dbconn.Close()
Set dbconn = Nothing
%>
