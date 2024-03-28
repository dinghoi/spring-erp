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

if goods_type = "" then
   goods_type = "전체"
end if

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = " 본사 출고의뢰 현황 -- "+ goods_type +" (" + from_date + " ∼ " + to_date + ")"

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
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

rele_id = "본사출고"

order_Sql = " ORDER BY rele_date,rele_no,rele_seq DESC"
if view_condi = "전체" then
   if goods_type = "전체" then
      where_sql = " WHERE (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')" 
	  else
	  where_sql = " WHERE (rele_goods_type = '"&goods_type&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "전체" then
      where_sql = " WHERE (rele_company = '"&view_condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')"
	  else
	  where_sql = " WHERE (rele_goods_type = '"&goods_type&"') and (rele_company = '"&view_condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')"
   end if
end if   

sql = "select * from met_chulgo_reg " + where_sql + order_sql 
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
                                <th scope="col">의뢰일자</th>
                                <th scope="col">의뢰번호</th>
                                <th scope="col">용도구분</th>
                                <th scope="col">신청자</th>
                                <th scope="col">소속</th>
                                <th scope="col">의뢰창고</th>

                                <th scope="col">출고(예정)창고</th>
                                <th scope="col">출고(예정)일자</th>
                                <th scope="col">출고상태</th>
                                <th scope="col">적요</th>

                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
                                <th scope="col">품목명</th>
                                <th scope="col">규격</th>
                                <th scope="col">상태</th>
                                <th scope="col">의뢰수량</th>
                                
                                <th scope="col">서비스No</th>
                                <th scope="col">고객사</th>
                                <th scope="col">점소명</th>
                                <th scope="col">비고(사유)</th>
							</tr>
						</thead>
						<tbody>
			<%
						i = 0
						do until rs.eof
                           i = i + 1
						   rele_no = rs("rele_no")
						   rele_seq = rs("rele_seq")
						   rele_date = rs("rele_date")

						   k = 0
                           sql = "select * from met_chulgo_reg_goods where (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"') and (rele_date = '"&rele_date&"')  ORDER BY rl_goods_seq,rl_goods_code ASC"
	                       Rs_buy.Open Sql, Dbconn, 1	
	                       while not Rs_buy.eof
		                     k = k + 1
							 if k = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_no")%>-<%=rs("rele_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_goods_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_emp_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_org_name")%>-<%=rs("rele_saupbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_stock_company")%>-<%=rs("rele_stock_name")%></td>

                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_stock_company")%>-<%=rs("chulgo_stock_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_ing")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_memo")%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_goods_gubun")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_goods_code")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_goods_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_standard")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_goods_grade")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("rl_qty"),0)%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_service_no")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_trade_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_trade_dept")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("rl_bigo")%></td>
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
                                    
								    <td class="left" ><%=Rs_buy("rl_goods_gubun")%></td>
                                    <td class="left" ><%=Rs_buy("rl_goods_code")%></td>
                                    <td class="left" ><%=Rs_buy("rl_goods_name")%></td>
                                    <td class="left" ><%=Rs_buy("rl_standard")%></td>
                                    <td class="left" ><%=Rs_buy("rl_goods_grade")%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("rl_qty"),0)%></td>
                                    
                                    <td class="left" ><%=Rs_buy("rl_service_no")%></td>
                                    <td class="left" ><%=Rs_buy("rl_trade_name")%></td>
                                    <td class="left" ><%=Rs_buy("rl_trade_dept")%></td>
                                    <td class="left" ><%=Rs_buy("rl_bigo")%></td>
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
