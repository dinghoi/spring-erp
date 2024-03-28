<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
view_condi=Request("view_condi")
stock = request("stock")
goods_type=request("goods_type")
goods_name = request("goods_name")
from_date=request("from_date")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

if goods_name = "" then
       title_line =  " 품목별 입고 현황 -- "+ stock +" (" + from_date + " ∼ " + to_date + ")"
   else
       title_line = goods_name + " 입고 현황 -- "+ stock +" (" + from_date + " ∼ " + to_date + ")"
end if

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
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stin = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY stin_date DESC"  

if view_condi = "전체" then
   if goods_type = "전체" then
      where_sql = " WHERE (stin_id = '구매') and (stin_date >= '"+from_date+"' and stin_date <= '"+to_date+"')" 
	  else
	  where_sql = " WHERE (stin_id = '구매') and (stin_goods_type = '"&goods_type&"') and (stin_date >= '"+from_date+"' and stin_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "전체" then
      where_sql = " WHERE (stin_id = '구매') and (stin_stock_company = '"&view_condi&"') and (stin_date >= '"+from_date+"' and stin_date <= '"+to_date+"')"
	  else
	  where_sql = " WHERE (stin_id = '구매') and (stin_goods_type = '"&goods_type&"') and (stin_stock_company = '"&view_condi&"') and (stin_date >= '"+from_date+"' and stin_date <= '"+to_date+"')"
   end if
end if   

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (stin_stock_name like '%"&stock&"%') "
end if

if goods_name = "" then
       goods_name_sql = ""
   else
       goods_name_sql = " and (stin_goods_name like '%"&goods_name&"%') "
end if

sql = "select * from met_stin_goods " + where_sql + stock_sql + goods_name_sql + order_sql
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
								<th class="first" scope="col">입고일자</th>
                                <th scope="col">용도구분</th>
                                <th scope="col">입고번호</th>
                                <th scope="col">입고구분</th>
                                <th scope="col">그룹사</th>
                                <th scope="col">사업부</th>
                                <th scope="col">입고창고</th>

                                <th scope="col">구매거래처</th>
                                <th scope="col">사업자번호</th>
                                
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
                                <th scope="col">품목명</th>
                                <th scope="col">규격</th>
                                <th scope="col">입고수량</th>
                                <th scope="col">입고단가</th>
                                <th scope="col">입고액</th>
                                
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>
			<%
						i = 0
						do until rs.eof
                           i = i + 1
						   stin_in_date = rs("stin_date")
						   stin_order_no = rs("stin_order_no")
						   stin_order_seq = rs("stin_order_seq")
						   
						   sql = "select * from met_stin where (stin_in_date = '"&stin_in_date&"') and (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"')"
						   Set Rs_stin=DbConn.Execute(Sql)
						   if Rs_stin.eof or Rs_stin.bof then
								stin_buy_company = ""
								stin_buy_saupbu = ""
								stin_trade_name = ""
								stin_trade_no = ""
								stin_memo = ""
							  else
							  	stin_buy_company = Rs_stin("stin_buy_company")
								stin_buy_saupbu = Rs_stin("stin_buy_saupbu")
								stin_trade_name = Rs_stin("stin_trade_name")
								stin_trade_no = Rs_stin("stin_trade_no")
								stin_memo = Rs_stin("stin_memo")
						   end if
						   Rs_stin.close()
						   
						   stin_trade_no = mid(stin_trade_no,1,3) + "-" + mid(stin_trade_no,4,2) + "-" + mid(stin_trade_no,6)
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_goods_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_order_no")%>-<%=rs("stin_order_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_id")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=stin_buy_company%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=stin_buy_saupbu%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_stock_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=stin_trade_name%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=stin_trade_no%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_goods_gubun")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_goods_code")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_goods_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_standard")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("stin_qty"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("stin_unit_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("stin_amt"),0)%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=stin_memo%></td>
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
