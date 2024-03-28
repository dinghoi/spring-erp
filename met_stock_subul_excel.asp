<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

view_condi=Request("view_condi")
goods_type=request("goods_type")
owner_view=request("owner_view")
condi=request("condi")
stock=request("stock")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

savefilename = curr_date + "품목별 수불현황.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename


If view_condi = "" Then
	view_condi = "케이원정보통신"
	stock = ""
	goods_type = "상품"
	owner_view = "C"
	ck_sw = "n"
	condi = ""
End If

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY stock_company,stock_goods_grade,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_code ASC"
if goods_type = "전체" then
   if condi = "" then
         where_sql = " WHERE (stock_company = '"&view_condi&"')" 
      else  
      if owner_view = "C" then 
             where_sql = " WHERE (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
         else
		     where_sql = " WHERE (stock_company = '"&view_condi&"') and (stock_goods_code = '"+condi+"')"
	   end if
   end if   
  else
   if condi = "" then
         where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"')" 
      else  
      if owner_view = "C" then 
             where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
         else
		     where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_code = '"+condi+"')"
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

sql = "select * from met_stock_gmaster " + where_sql + stock_sql + order_sql
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 품목별 수불현황 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
				              <th rowspan="2" class="first" scope="col">코드</th>
				              <th rowspan="2" scope="col">품목구분</th>
                              <th rowspan="2" scope="col">품목명</th>
                              <th rowspan="2" scope="col">규격</th>
                              <th rowspan="2" scope="col">상태</th>
                              <th colspan="2" scope="col">기초</th>
				              <th colspan="2" scope="col">입고</th>
                              <th colspan="2" scope="col">출고</th>
                              <th colspan="2" scope="col">기말</th>
			                </tr>
                            <tr>
				              <th scope="col">수량</th>
                              <th scope="col">금액</th>
                              <th scope="col">수량</th>
                              <th scope="col">금액</th>
                              <th scope="col">수량</th>
                              <th scope="col">금액</th>
                              <th scope="col">수량</th>
                              <th scope="col">금액</th>
                            </tr>
						</thead>
                        <tbody>
					<%
						do until rs.eof
							  
	           		%>
							<tr>
				              <td class="first"><%=rs("stock_goods_code")%>&nbsp;</td>
                              <td><%=rs("stock_goods_gubun")%>&nbsp;</td>
                              <td><%=rs("stock_goods_name")%>&nbsp;</td>
                              <td><%=rs("stock_goods_standard")%>&nbsp;</td>
                              <td><%=rs("stock_goods_grade")%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("stock_last_qty"),0)%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("stock_last_amt"),0)%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("stock_in_amt"),0)%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("stock_go_amt"),0)%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("stock_jj_qty"),0)%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("stock_jj_amt"),0)%>&nbsp;</td>
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

