<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
from_date=request("from_date")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "������ ��Ȳ -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "��ü" then
   Sql = "select * from emp_master where emp_end_date >= '"+from_date+"' and emp_end_date <= '"+to_date+"' ORDER BY emp_no,emp_name ASC"
   else  
   Sql = "select * from emp_master where emp_company = '"+view_condi+"' and emp_end_date >= '"+from_date+"' and emp_end_date <= '"+to_date+"' ORDER BY emp_no,emp_name ASC"
end if
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=from_date%>&nbsp;��&nbsp;<%=to_date%> &nbsp;������ ��Ȳ>&nbsp;(<%=view_condi%>)</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">���</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">�ֹι�ȣ</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��å</div></td>
    <td><div align="center" class="style1">�Ҽ�</div></td>
    <td><div align="center" class="style1">ȸ��</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">�����</div></td>
    <td><div align="center" class="style1">��</div></td>
    <td><div align="center" class="style1">����ó</div></td>
    <td><div align="center" class="style1">�Ǳٹ���</div></td>
    <td><div align="center" class="style1">�����Ի���</div></td>
    <td><div align="center" class="style1">�Ի���</div></td>
    <td><div align="center" class="style1">�ټӱ����</div></td>
    <td><div align="center" class="style1">���������</div></td>
    <td><div align="center" class="style1">���������</div></td>
    <td><div align="center" class="style1">������</div></td>
    <td><div align="center" class="style1">�������</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">�����ּ�</div></td>
    <td><div align="center" class="style1">���ּ�</div></td>
    <td><div align="center" class="style1">��ȭ��ȣ</div></td> 
    <td><div align="center" class="style1">�ڵ���</div></td>
    <td><div align="center" class="style1">e����</div></td>
    <td><div align="center" class="style1">��󿬶���</div></td>
    <td><div align="center" class="style1">����ȸ</div></td>
    <td><div align="center" class="style1">��ֿ���</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">���</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��ȥ�����</div></td>
  </tr>
    <%
		do until rs.eof 
		
		stay_name = ""
		stay_code = rs("emp_stay_code")
        if stay_code <> "" then
		   Sql="select * from emp_stay where stay_code = '"&stay_code&"'"
		   Rs_stay.Open Sql, Dbconn, 1
			  
		  if not rs_stay.eof then
             stay_name = rs_stay("stay_name")
		  end if
		  rs_stay.Close()
		end if
		emp_person2 = rs("emp_person2")
        if emp_person2 <> "" then
	       sex_id = mid(cstr(emp_person2),1,1)
	       if sex_id = "1" then
	             emp_sex = "��"
	     	  else
	    	     emp_sex = "��"
	       end if
	    end if
		
	if rs("emp_birthday") = "1900-01-01" then
		   emp_birthday = ""
	   else 
		   emp_birthday = rs("emp_birthday")
	end if
	if rs("emp_military_date1") = "1900-01-01" then
           emp_military_date1 = ""
           emp_military_date2 = ""
	   else 
	       emp_military_date1 = rs("emp_military_date1")
           emp_military_date2 = rs("emp_military_date2")
    end if
    if rs("emp_marry_date") = "1900-01-01" then
           emp_marry_date = ""
		else
		   emp_marry_date = rs("emp_marry_date")
    end if
	if rs("emp_grade_date") = "1900-01-01" then
           emp_grade_date = ""
	   else
	       emp_grade_date = rs("emp_grade_date")
    end if
	if rs("emp_end_date") = "1900-01-01" then
           emp_end_date = ""
		else
		   emp_end_date = rs("emp_end_date")
    end if
	if rs("emp_org_baldate") = "1900-01-01" then
           emp_org_baldate = ""
	   else
	       emp_org_baldate = rs("emp_org_baldate")
    end if
	if rs("emp_sawo_date") = "1900-01-01" then
           emp_sawo_date = ""
	   else
	       emp_sawo_date = rs("emp_sawo_date")
    end if
 
    emp_email = rs("emp_email") + "@k-won.co.kr"

	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_sex%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_type")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_person1")%>-<%=rs("emp_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_reside_place")%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
    <td width="145"><div align="center" class="style1"><%=stay_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_gunsok_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_end_gisan")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_yuncha_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_end_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=emp_birthday%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_jikmu")%></div></td>
    <td width="350"><div align="center" class="style1"><%=rs("emp_family_zip")%>&nbsp;<%=rs("emp_family_sido")%>&nbsp;<%=rs("emp_family_gugun")%>&nbsp;<%=rs("emp_family_dong")%>&nbsp;<%=rs("emp_family_addr")%></div></td>

    <td width="350"><div align="center" class="style1"><%=rs("emp_zipcode")%>&nbsp;<%=rs("emp_sido")%>&nbsp;<%=rs("emp_gugun")%>&nbsp;<%=rs("emp_dong")%>&nbsp;<%=rs("emp_addr")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_tel_ddd")%>-<%=rs("emp_tel_no1")%>-<%=rs("emp_tel_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_email%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_emergency_tel")%></div></td>
    <% 'response.write(rs("emp_emergency_tel"))
	   'response.End %>
    <%
	if rs("emp_sawo_id") = "Y" then
	   emp_sawo_id = "����"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%>-<%=emp_sawo_date%></div></td>
    <% 
	   else
	   emp_sawo_id = "����"
	 %>
       <td width="145"><div align="center" class="style1"><%=emp_sawo_id%></div></td>
    <%
	end if
	%>
    <td width="145"><div align="center" class="style1"><%=rs("emp_disabled")%>&nbsp;<%=rs("emp_disab_grade")%></div></td>
    <td width="200"><div align="center" class="style1"><%=rs("emp_military_id")%>&nbsp;<%=emp_military_date1%>&nbsp;<%=emp_military_date2%>&nbsp;<%=rs("emp_military_grade")%>&nbsp;<%=rs("emp_military_comm")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_hobby")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_faith")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_marry_date%></div></td>
  </tr>
	<%
	Rs.MoveNext()
	loop
	%>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
