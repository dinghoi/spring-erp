<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

view_condi=Request("view_condi")
condi=Request("condi")
owner_view=Request("owner_view")

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = condi + "~ " + view_condi + " " + " ���� ���� ������Ȳ"

savefilename = title_line +".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if condi = "" then
      Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"')  and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC"
   else  
      if owner_view = "C" then 
            Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"') and (emp_name like '%"+condi+"%') ORDER BY emp_in_date,emp_no ASC"
         else
            Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"') and (emp_no = '"+condi+"') ORDER BY emp_in_date,emp_no ASC"
	  end if
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
    <td colspan="10" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">���</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��å</div></td>
    <td><div align="center" class="style1">�Ի���</div></td>
    <td><div align="center" class="style1">�Ҽ�</div></td>
    <td><div align="center" class="style1">�ŷ�����</div></td>
    <td><div align="center" class="style1">���¹�ȣ</div></td>
    <td><div align="center" class="style1">������</div></td>
    <td><div align="center" class="style1">����</div></td>
    <%' �Ʒ��κ��� �ϴ� ���Ƴ���... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">�԰� ���γ��� </div> %>
    <%'</div></td> %>
  </tr>
    <%
		if  view_condi <> "" then 
		do until rs.eof 
          bank_name = ""
		  account_no = ""
		  account_holder = ""
		  emp_no = rs("emp_no")
		  emp_person1 = rs("emp_person1")
		  emp_person2 = rs("emp_person2")
						  
		  Sql = "SELECT * FROM pay_bank_account where emp_no = '"&emp_no&"'"
          Set rs_bnk = DbConn.Execute(SQL)
		  if not rs_bnk.eof then
                bank_name = rs_bnk("bank_name")
			    account_no = rs_bnk("account_no")
				account_holder = rs_bnk("account_holder")
	         else
                bank_name = ""
			    account_no = ""
				account_holder = ""
          end if
          rs_bnk.close()
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=bank_name%></div></td>
    <td width="145"><div align="center" class="style1"><%=account_no%></div></td>
    <td width="145"><div align="center" class="style1"><%=account_holder%></div></td>
    <td width="400"><div align="center" class="style1"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
  </tr>
	<%
	Rs.MoveNext()
	loop
	
	end if
	%>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
