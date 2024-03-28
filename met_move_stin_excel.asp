<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
view_condi=Request("view_condi")
stock=request("stock")
goods_type=request("goods_type")
from_date=request("from_date")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = " 창고 입고 현황 -- "+ goods_type +" (" + from_date + " ∼ " + to_date + ")"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'mvin_id = "본사출고"

order_Sql = " ORDER BY mvin_in_date,mvin_in_stock,mvin_in_seq DESC"

if view_condi = "전체" then
   if goods_type = "전체" then
          where_sql = " WHERE (mvin_in_date >= '"+from_date+"' and mvin_in_date <= '"+to_date+"')" 
	  else
	      where_sql = " WHERE (mvin_goods_type = '"&goods_type&"') and (mvin_in_date >= '"+from_date+"' and mvin_in_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "전체" then
          where_sql = " WHERE (mvin_stock_company = '"&view_condi&"') and (mvin_in_date >= '"+from_date+"' and mvin_in_date <= '"+to_date+"')"
	  else
	      where_sql = " WHERE (mvin_goods_type = '"&goods_type&"') and (mvin_stock_company = '"&view_condi&"') and (mvin_in_date >= '"+from_date+"' and mvin_in_date <= '"+to_date+"')"
   end if
end if   

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (mvin_stock_name like '%"&stock&"%') "
end if

sql = "select * from met_mv_in " + where_sql + stock_sql + order_sql
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
                                 <th scope="col">용도구분</th>
                                 <th scope="col">출고일자</th>
                                 <th scope="col">출고번호</th>
                                 <th scope="col">출고구분</th>
                                 <th scope="col">출고창고</th>
                                 <th scope="col">입고일자</th>
                                 <th scope="col">입고번호</th>
                                 <th scope="col">입고창고</th>
                                 <th scope="col">적요</th>

                                 <th scope="col">품목구분</th>
                                 <th scope="col">품목코드</th>
                                 <th scope="col">품목명</th>
                                 <th scope="col">규격</th>
                                 <th scope="col">상태</th>
                                 <th scope="col">입고수량</th>
							</tr>
						</thead>
						<tbody>
			<%
						i = 0
						do until rs.eof
                           i = i + 1
						   mvin_in_date = rs("mvin_in_date")
						   mvin_in_stock = rs("mvin_in_stock")
						   mvin_in_seq = rs("mvin_in_seq")
						   
						   chulgo_date = rs("chulgo_date")
						   chulgo_stock = rs("chulgo_stock")
						   chulgo_seq = rs("chulgo_seq")
				   
						   mvin_no = mid(cstr(rs("mvin_in_date")),3,2) + mid(cstr(rs("mvin_in_date")),6,2) + mid(cstr(rs("mvin_in_date")),9,2) 

						   k = 0
                           sql = "select * from met_mv_in_goods where (mvin_in_date = '"&mvin_in_date&"') and (mvin_in_stock = '"&mvin_in_stock&"') and (mvin_in_seq = '"&mvin_in_seq&"')  ORDER BY in_goods_seq,in_goods_code ASC"
	                       Rs_reg.Open Sql, Dbconn, 1	
	                       while not Rs_reg.eof
		                     k = k + 1
							 if k = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("mvin_goods_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=chulgo_date%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_no")%>&nbsp;<%=rs("rele_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("mvin_id")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_stock_name")%>(<%=rs("chulgo_stock")%>)</td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("mvin_in_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=mvin_no%><%=rs("mvin_in_stock")%>&nbsp;<%=rs("mvin_in_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("mvin_stock_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_memo")%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("in_goods_gubun")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("in_goods_code")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("in_goods_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("in_standard")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("in_goods_grade")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_reg("in_qty"),0)%></td>
						         </tr>
            <%
			                    else
		    %>		
                                 <tr>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    
								    <td class="left" ><%=Rs_reg("in_goods_gubun")%></td>
                                    <td class="left" ><%=Rs_reg("in_goods_code")%></td>
                                    <td class="left" ><%=Rs_reg("in_goods_name")%></td>
                                    <td class="left" ><%=Rs_reg("in_standard")%></td>
                                    <td class="left" ><%=Rs_reg("in_goods_grade")%></td>
                                    <td align="right"><%=formatnumber(Rs_reg("in_qty"),0)%></td>
						         </tr>            
            <%            							
							 end if
		                         Rs_reg.movenext()
	                       Wend
                           Rs_reg.close()
							  
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
