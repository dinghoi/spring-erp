<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
stock_level=Request("stock_level")

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = " 창고코드 현황 -- " + stock_level 

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY stock_level,stock_code ASC"
if condi = "" then
      where_sql = " WHERE (stock_level = '"&stock_level&"')" 
   else  
      if owner_view = "C" then 
             where_sql = " WHERE (stock_level = '"&stock_level&"') and (stock_name like '%"+condi+"%')"
         else
		     where_sql = " WHERE (stock_level = '"&stock_level&"') and (stock_code = '"+condi+"')"
	   end if
end if   

sql = "select * from met_stock_code " + where_sql + order_sql
Rs.Open Sql, Dbconn, 1
	

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>상품자재관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">창고코드</th>
				                <th scope="col">창고명</th>
                                <th scope="col">창고유형</th>
                                <th scope="col">창고장</th>
                                <th scope="col">회사</th>
                                <th scope="col">생성일</th>
                                <th scope="col">폐쇄일</th>
                                <th scope="col">출고담당</th>
                                <th scope="col">입고담당</th>
                                <th scope="col">소속조직</th>
							</tr>
						</thead>
						<tbody>
			<%
						do until rs.eof
						   stock_end_date = rs("stock_end_date")
						   if stock_end_date = "1900-01-01" then
	                            stock_end_date = ""
	                       end if
		    %>
                                 <tr>
								    <td class="first"><%=rs("stock_code")%>&nbsp;</td>
                                    <td><%=rs("stock_name")%>&nbsp;</td>
                                    <td><%=rs("stock_level")%>&nbsp;</td>
                                    <td><%=rs("stock_manager_name")%>(<%=rs("stock_manager_code")%>)&nbsp;</td>
                                    <td><%=rs("stock_company")%>&nbsp;</td>
                                    <td><%=rs("stock_open_date")%>&nbsp;</td>
                                    <td><%=stock_end_date%>&nbsp;</td>
                                    <td><%=rs("stock_go_name")%>&nbsp;</td>
                                    <td><%=rs("stock_in_name")%>&nbsp;</td>
                                    <td class="left"><%=rs("stock_bonbu")%>-<%=rs("stock_saupbu")%>-<%=rs("stock_team")%>&nbsp;</td>
						         </tr>            
            <%            							
						   rs.movenext()
						loop
						rs.close()
		    %>						
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
