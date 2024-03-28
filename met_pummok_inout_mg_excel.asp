<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim rs
Dim rs_numRows

view_condi = request("view_condi")
goods_type = request("goods_type")
from_date = request("from_date")
to_date = request("to_date")
stock = request("stock")

title_line = stock + " < " + goods_type + " > 입★출고 현황 "

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_jae = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

order_Sql = " ORDER BY stock_code,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_code,stock_date,id_seq,inout_no,inout_seq ASC"
where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_date >= '"+from_date+"' and stock_date <= '"+to_date+"')" 

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (stock_name like '%"&stock&"%') "
end if

sql = "select * from met_stock_inout_goods " + where_sql + stock_sql + order_sql 

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
								<th class="first" scope="col">대분류</th>
                                <th scope="col">중분류</th>
                                <th scope="col">규격</th>
				                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
                                <th scope="col">품목명</th>
                                <th scope="col">상태</th>
                                <th scope="col">창고</th>
                                <th scope="col">입고창고</th>
                                <th scope="col">입출고일자</th>
                                <th scope="col">입출고구분</th>
				                <th scope="col">입고</th>
                                <th scope="col">출고</th>
                                <th scope="col">현재고</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>     
						<%
						    sum_in_cnt = 0	 
						    sum_out_cnt = 0
						    sum_jae_cnt = 0
						
						    if rs.eof or rs.bof then
							    bi_goods_code = ""
								bi_stock = ""
							    bi_goods_type = ""
					          else						  
							    if isnull(rs("stock_goods_code")) or rs("stock_goods_code") = "" then	
								    bi_goods_code = ""
									bi_stock = ""
							        bi_goods_type = ""
							      else
								    bi_goods_code = rs("stock_goods_code")
									bi_stock = rs("stock_code")
							        bi_goods_type = rs("stock_goods_type")
							    end if
						    end if
						
							do until rs.eof or rs.bof
                               if isnull(rs("stock_goods_code")) or rs("stock_goods_code") = "" then
								         goods_goods_code = ""
							      else
							  	         goods_goods_code = rs("stock_goods_code")
						       end if

						       if bi_goods_code <> goods_goods_code then
							         Sql = "SELECT * FROM met_goods_code where goods_code = '"&bi_goods_code&"'"
                                     Set Rs_good = DbConn.Execute(SQL)
							         if not Rs_good.eof then
								          goods_goods_name = Rs_good("goods_name")
							         end if
							         Rs_good.close()
								 
								     sql="select * from met_stock_gmaster where stock_code='"&bi_stock&"' and stock_goods_code='"&bi_goods_code&"' and stock_goods_type='"&bi_goods_type&"'"
	                                 set Rs_jae=dbconn.execute(sql)
                                     if not Rs_jae.eof then
								            jj_a_qty = Rs_jae("stock_JJ_qty")
									    else
									        jj_a_qty = 0
								     end if
								     Rs_jae.close()
								 
								     sum_jjj_cnt = jj_a_qty - sum_in_cnt + sum_out_cnt
								     sum_jae_cnt = sum_jjj_cnt + sum_in_cnt - sum_out_cnt
					  %>
                                 <tr>
								    <td colspan="9" bgcolor="#EEFFFF" align="center"><%=goods_goods_name%>&nbsp;(<%=bi_goods_code%>)&nbsp;&nbsp;계</td>
							        <td bgcolor="#EEFFFF" >재고</th>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_jjj_cnt,0)%>&nbsp;</td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_in_cnt,0)%>&nbsp;</td>
							        <td bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_out_cnt,0)%>&nbsp;</td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_jae_cnt,0)%>&nbsp;</td>
							        <td bgcolor="#EEFFFF" >&nbsp;</th>
						         </tr>
                      <%
							         sum_in_cnt = 0	 
						             sum_out_cnt = 0
						             sum_jae_cnt = 0
								     bi_goods_code = goods_goods_code
									 bi_stock = rs("stock_code")
							         bi_goods_type = rs("stock_goods_type")
						        end if
								
								stock_goods_type = rs("stock_goods_type")
						        stock_goods_code = rs("stock_goods_code")

                                sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
                                Set Rs_good = DbConn.Execute(SQL)
                                if not Rs_good.eof then
    	                               goods_level1 = Rs_good("goods_level1")
		                               goods_level2 = Rs_good("goods_level2")
                                   else
		                               goods_level1 = ""
		                               goods_level2 = ""
                                end if
                                Rs_good.close()

						        sum_in_cnt = sum_in_cnt + int(rs("stock_in_qty"))
						        sum_out_cnt = sum_out_cnt + int(rs("stock_go_qty"))
						        sum_jae_cnt = sum_jae_cnt + 0								
								
						%>
							<tr>
                                <td class="first"><%=goods_level1%>&nbsp;</td>
                                <td><%=goods_level2%>&nbsp;</td>
                                <td><%=rs("stock_goods_standard")%>&nbsp;</td>
                                <td><%=rs("stock_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("stock_goods_code")%>&nbsp;</td>
                                <td><%=rs("stock_goods_name")%>&nbsp;</td>
                                <td><%=rs("stock_goods_grade")%>&nbsp;</td>
                                <td><%=rs("stock_name")%>&nbsp;</td>
                                <td><%=rs("rele_stock_name")%>&nbsp;</td>
                                <td><%=rs("stock_date")%>&nbsp;</td>
                                <td><%=rs("stock_id")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</td>
                                <td align="right">&nbsp;</td>
                                <td>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
							
						    Sql = "SELECT * FROM met_goods_code where goods_code = '"&bi_goods_code&"'"
                            Set Rs_good = DbConn.Execute(SQL)
						    if not Rs_good.eof then
						          goods_goods_name = Rs_good("goods_name")
						    end if
				            Rs_good.close()
						
						    sql="select * from met_stock_gmaster where stock_code='"&bi_stock&"' and stock_goods_code='"&bi_goods_code&"' and stock_goods_type='"&bi_goods_type&"'"
	                        set Rs_jae=dbconn.execute(sql)
                            if not Rs_jae.eof then
						            jj_a_qty = Rs_jae("stock_JJ_qty")
							    else
							        jj_a_qty = 0
						    end if
						    Rs_jae.close()
								 
						    sum_jjj_cnt = jj_a_qty - sum_in_cnt + sum_out_cnt
						    sum_jae_cnt = sum_jjj_cnt + sum_in_cnt - sum_out_cnt				
						%>
                            <tr>
						       <td colspan="9" bgcolor="#EEFFFF" align="center"><%=goods_goods_name%>&nbsp;(<%=bi_goods_code%>)&nbsp;&nbsp;계</td>
						       <td bgcolor="#EEFFFF" >재고</th>
                               <td bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_jjj_cnt,0)%>&nbsp;</td>
                               <td bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_in_cnt,0)%>&nbsp;</td>
						       <td bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_out_cnt,0)%>&nbsp;</td>
                               <td bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_jae_cnt,0)%>&nbsp;</td>
						       <td bgcolor="#EEFFFF" >&nbsp;</th>
						    </tr>                        
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
