<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/itft2005_db.asp" -->

<%
Dim Rs
Dim Repeat_Rows

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_wait = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "SELECT count(acpt_no) as wait_no  FROM as_acpt WHERE ( reside_place = '�ݼ���' ) and ( as_type = '����ó��' ) and ( as_process = '����' )"
Rs_wait.Open Sql, Dbconn, 1
wait_no = rs_wait("wait_no") + 1

Sql = "SELECT * FROM as_acpt WHERE ( reside_place = '�ݼ���' ) and ( as_type = '����ó��' ) and ( as_process = '����' or as_process = 'Ȯ��' ) ORDER BY acpt_no DESC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="include/itft_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style13 {color: #003366}
.style14 {font-family: "����", "����", Seoul, "�Ѱ�ü"}
.style15 {font-family: "����ü", "����ü", Seoul}
.style22 {font-family: Arial, Seoul}
-->
</style>
</head>

<body>
<table width="650" border="0">
  <tr> 
    <td width="800" height="41"><img src="image/as_acpt_view_title.gif" width="650" height="40"></td>
  </tr>
  <tr> 
    <td width="800"><form action="" method="post" name="form1">
      <table width="650"  border="1" cellpadding="0" cellspacing="0">
        <tr valign="middle" bgcolor="#EFEFEF" class="style11B">
          <td width="120" height="28"><div align="center">��������</div></td>
          <td width="55"><div align="center">����NO</div></td>
          <td width="50"><div align="center">����</div></td>
          <td width="50"><div align="center">����</div></td>
          <td width="60"><div align="center">�����</div></td>
          <td width="70"><div align="center">ȸ��</div></td>
          <td width="135"><div align="center">������</div></td>
          <td width="40"><div align="center">���</div></td>
          <td width="50"><div align="center">FAQ</div></td>
        </tr>
        <%
  	Repeat_rows = 10 
	Repeat_index = 0
	While ((Repeat_Rows <> 0) AND (NOT Rs.EOF)) 
    %>
        <tr valign="middle" class="style11">
    <%
	int date_len 
	date_len=len(rs("acpt_date"))
	as_memo = rs("as_memo")
	if rs("acpt_man") = "���ͳ�" then
		acpt_type = "���ͳ�"
	  else
	  	acpt_type = "��ȭ"
	end if
	if rs("as_process") = "Ȯ��" then
		wait_no_view = 0
		as_process = "������"
	  else
		wait_no = wait_no - 1
		wait_no_view = wait_no
	  	as_process = "����"
	end if
	%>
          <td width="120" height="27"><div align="center"><%=mid(cstr(rs("acpt_date")),3,date_len-5)%></div></td>
          <td width="55" height="27"><div align="center"><%=rs("acpt_no")%></div></td>
          <td width="50" height="27"><div align="center"><%=acpt_type%></div></td>
          <td width="50" height="27"><div align="center"><%=as_process%></div></td>
          <td width="60" height="27" class="style13"><div align="center">
            <p style="cursor:pointer"><span title="<%=as_memo%>"><%=rs("acpt_user")%></span></p>
          </div></td>
          <td width="70" height="27"><div align="center"><%=rs("company")%></div></td>
          <td width="135" height="27"><div align="center"><%=rs("dept")%></div></td>
          <td width="40" height="27"><div align="center"><%=wait_no_view%></div></td>
          <td width="50" height="27"><div align="center">1 / 1 </div></td>
        </tr>
        <% 
  	Repeat_index=Repeat_index+1
'	Repeat1__numRows=Repeat1__numRows-1
	Repeat_rows = Repeat_rows - 1
	Rs.MoveNext()
	Wend
%>
        <%
intstart = (int((page-1)/10)*10) + 1
intend = intstart + 9
first_page = 1

if intend > total_page then
	intend = total_page
end if
%>
      </table>
      <table width="650">
        <tr>
          <td height="33"><span class="style1"></span>
              <div align="center" class="style12"></div></td>
        </tr>
      </table>
    </form>	</td>
  </tr>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
