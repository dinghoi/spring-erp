<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
view_condi=Request("view_condi")
goods_type=request("goods_type")
from_date=request("from_date")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = goods_type + " 구매품의 현황 -- "+ view_condi +" (" + from_date + " ∼ " + to_date + ")"

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

order_Sql = " ORDER BY buy_company,buy_date,buy_no,buy_seq DESC"
if view_condi = "전체" then
   if goods_type = "전체" then
      where_sql = " WHERE (buy_date >= '"+from_date+"' and buy_date <= '"+to_date+"')" 
	  else
	  where_sql = " WHERE (buy_goods_type = '"&goods_type&"') and (buy_date >= '"+from_date+"' and buy_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "전체" then
      where_sql = " WHERE (buy_company = '"&view_condi&"') and (buy_date >= '"+from_date+"' and buy_date <= '"+to_date+"')"
	  else
	  where_sql = " WHERE (buy_goods_type = '"&goods_type&"') and (buy_company = '"&view_condi&"') and (buy_date >= '"+from_date+"' and buy_date <= '"+to_date+"')"
   end if
end if   

sql = "select * from met_buy " + where_sql + order_sql
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
								<th class="first" scope="col">구매번호</th>
                                <th scope="col">용도구분</th>
				                <th scope="col">구매품의일</th>
                                <th scope="col">회사</th>
                                <th scope="col">사업부</th>
                                <th scope="col">부서</th>
                                <th scope="col">구매담당</th>
                                <th scope="col">구매거래처</th>
                                <th scope="col">사업자번호</th>
                                <th scope="col">거래처담당자</th>
                                <th scope="col">이메일</th>
                                <th scope="col">구매총액</th>
                                <th scope="col">구매금액</th>
                                <th scope="col">부가세</th>
                                <th scope="col">진행상태</th>
                                <th scope="col">적요</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
                                <th scope="col">품목명</th>
                                <th scope="col">규격</th>
                                <th scope="col">수량</th>
                                <th scope="col">구입단가</th>
                                <th scope="col">구입액</th>
							</tr>
						</thead>
						<tbody>
			<%
						i = 0
						do until rs.eof
                           i = i + 1
						   buy_no = rs("buy_no")
		                   buy_seq = rs("buy_seq")
						   buy_date = rs("buy_date")
						   buy_ing = rs("buy_ing")
						   buy_ing_gubun = ""
						   if buy_ing = "0" then
						         buy_ing_gubun = "구매의뢰"
						      elseif buy_ing = "1" then
							            buy_ing_gubun = "발주등록"
									 elseif buy_ing = "2" then
							                   buy_ing_gubun = "발주"
										    elseif buy_ing = "3" then
							                          buy_ing_gubun = "입고"
						   end if
					   
						   k = 0
                           sql = "select * from met_buy_goods where (bg_no = '"&buy_no&"') and (bg_date = '"&buy_date&"') and (buy_seq = '"&buy_seq&"') ORDER BY bg_seq,bg_goods_code ASC"
	                       Rs_buy.Open Sql, Dbconn, 1	
	                       while not Rs_buy.eof
		                     k = k + 1
							 if k = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_no")%>-<%=rs("buy_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_goods_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_saupbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_org_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_emp_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_trade_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_trade_no")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_trade_person")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_trade_email")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("buy_price"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("buy_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("buy_cost_vat"),0)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=buy_ing_gubun%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("buy_memo")%></td>
                                    
								    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("bg_goods_gubun")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("bg_goods_code")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("bg_goods_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("bg_standard")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("bg_qty"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("bg_unit_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("bg_buy_amt"),0)%></td>
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
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td align="right">&nbsp;</td>
                                    <td align="right">&nbsp;</td>
                                    <td align="right">&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    
								    <td class="left" ><%=Rs_buy("bg_goods_gubun")%></td>
                                    <td class="left" ><%=Rs_buy("bg_goods_code")%></td>
                                    <td class="left" ><%=Rs_buy("bg_goods_name")%></td>
                                    <td class="left" ><%=Rs_buy("bg_standard")%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("bg_qty"),0)%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("bg_unit_cost"),0)%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("bg_buy_amt"),0)%></td>
						         </tr>            
            <%            							
							 end if
		                         Rs_buy.movenext()
	                       Wend
                           Rs_buy.close()
							  
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
