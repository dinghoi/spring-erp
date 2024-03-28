<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

view_condi=Request("view_condi")
goods_type=Request("goods_type")
goods_gubun=Request("goods_gubun")
owner_view=Request("owner_view")
condi=Request("condi")
stock=Request("stock")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = view_condi + goods_type + "재고현황" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY stock_company,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_code ASC"

if goods_type = "전체" then 
   if condi = "" then
         where_sql = " WHERE (stock_company = '"&view_condi&"')" 
      else  
         if owner_view = "C" then 
                where_sql = " WHERE (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
            else
		        where_sql = " WHERE (stock_company = '"&view_condi&"') and (stock_goods_code like '%"+condi+"%')"
   	      end if
   end if
  else
   if condi = "" then
         where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"')" 
      else  
         if owner_view = "C" then 
                where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
            else
		        where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_code like '%"+condi+"%')"
	      end if
   end if   
end if

'if stock = "" then
'       stock_sql = ""
'   else
'       stock_sql = " and (stock_code = '"&stock&"') "
'end if

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (stock_name like '%"&stock&"%') "
end if

if goods_gubun = "" then
       gubun_sql = ""
   else
       gubun_sql = " and (stock_goods_gubun like '%"&goods_gubun&"%') "
end if

' 입출고 및 재고수량이 없는것
cnt_sql = " and (stock_in_qty > 0 or stock_go_qty > 0 or stock_JJ_qty > 0) "

sql = "select * from met_stock_gmaster " + where_sql + stock_sql + gubun_sql + cnt_sql + order_sql 
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
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=stock%> &nbsp;(<%=goods_type%>)&nbsp;재고 현황&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">용도구분</div></td>
    <td><div align="center" class="style1">품목코드</div></td>
    <td><div align="center" class="style1">품목구분</div></td>
    <td><div align="center" class="style1">품목명</div></td>
    <td><div align="center" class="style1">규격</div></td>
    <td><div align="center" class="style1">상태</div></td>
    <td><div align="center" class="style1">창고</div></td>
    <td><div align="right" class="style1">전년이월수량</div></td>
    <td><div align="right" class="style1">입고수량</div></td>
    <td><div align="right" class="style1">출고수량</div></td>
    <td><div align="right" class="style1">현재고수량</div></td>
    <td><div align="center" class="style1">비고</div></td>
  </tr>
    <%
		do until rs.eof 
		
	%>
  <tr valign="middle" class="style11">
    <td width="145"><div align="center" class="style1"><%=view_condi%></div></td>
    <td width="145"><div align="center" class="style1"><%=goods_type%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_code")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_gubun")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_name")%></div></td>
    <td width="160"><div align="center" class="style1"><%=rs("stock_goods_standard")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_goods_grade")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("stock_name")%>(<%=rs("stock_code")%>)</div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_last_qty"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_in_qty"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_go_qty"),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("stock_JJ_qty"),0)%></div></td>
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
