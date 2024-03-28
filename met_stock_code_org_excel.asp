<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

view_condi=Request("view_condi")
sel_company=Request("sel_company")
sel_bonbu=Request("sel_bonbu")
sel_saupbu=Request("sel_saupbu")
sel_team=Request("sel_team")

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = " 창고코드 현황(조직별) -- " + sel_company 

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_tab = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "1" then
   condi_Sql = " and (stock_company = '" + sel_company + "')"
end if

if view_condi = "2" then
   condi_Sql = " and (stock_company = '"+sel_company+"') and (stock_bonbu = '" + sel_bonbu + "')"
end if

if view_condi = "3" then
   condi_Sql = " and (stock_company = '"+sel_company+"') and (stock_bonbu = '" + sel_bonbu + "') and (stock_saupbu = '" + sel_saupbu + "')"
end if

if view_condi = "4" then
   condi_Sql = " and (stock_company = '"+sel_company+"') and (stock_bonbu = '" + sel_bonbu + "') and (stock_saupbu = '" + sel_saupbu + "') and (stock_team = '" + sel_team + "')"
end if

order_Sql = " ORDER BY stock_level,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_name DESC" 
'order_Sql = " ORDER BY stock_code " + view_sort
where_sql = " WHERE (isNull(stock_end_date) or stock_end_date = '1900-01-01')"

sql = "select * from met_stock_code " + where_sql + condi_sql + order_sql
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
