<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

view_condi=Request("view_condi")
goods_type=Request("goods_type")
field_check=Request("field_check")

If field_check = "stock_jj_amt" Then
	    field_view = " �ݾ׼� "
   elseif field_check = "stock_JJ_qty" Then
              field_view = " ������ "
End If

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = field_view + " �����Ȳ" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY " + field_check + " DESC"

if goods_type = "��ü" then
     sql = "select * from met_stock_gmaster where (stock_company = '"&view_condi&"') " + order_sql
   else
     sql = "select * from met_stock_gmaster where (stock_company = '"&view_condi&"') and (stock_goods_type = '"&goods_type&"') " + order_sql
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
    <td colspan="17" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=field_view%> &nbsp;(<%=goods_type%>)&nbsp;��� ��Ȳ&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">ȸ��</div></td>
    <td><div align="center" class="style1">�뵵����</div></td>
    <td><div align="center" class="style1">ǰ���ڵ�</div></td>
    <td><div align="center" class="style1">ǰ�񱸺�</div></td>
    <td><div align="center" class="style1">ǰ���</div></td>
    <td><div align="center" class="style1">�԰�</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">â��</div></td>
    <td><div align="right" class="style1">�����̿�����</div></td>
    <td><div align="right" class="style1">�����̿��ݾ�</div></td>
    <td><div align="right" class="style1">�԰����</div></td>
    <td><div align="right" class="style1">�԰�ݾ�</div></td>
    <td><div align="right" class="style1">������</div></td>
    <td><div align="right" class="style1">���ݾ�</div></td>
    <td><div align="right" class="style1">��������</div></td>
    <td><div align="right" class="style1">�����ݾ�</div></td>
    <td><div align="center" class="style1">���</div></td>
  </tr>
    <%
		do until rs.eof 
		
	%>
  <tr valign="middle" class="style11">
    <td width="145"><div align="center" class="style1"><%=view_condi%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_type")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_code")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_gubun")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_name")%></div></td>
    <td width="160"><div align="center" class="style1"><%=rs("stock_goods_standard")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_grade")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_name")%>(<%=rs("stock_code")%>)</div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_last_qty"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_last_amt"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_in_qty"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_in_amt"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_go_qty"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_go_amt"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_JJ_qty"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_JJ_amt"),0)%></div></td>
    <td width="145"><div align="center" class="style1">&nbsp;</div></td>
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
