<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

view_condi = request("view_condi")
goods_type = request("goods_type")
goods_gubun = request("goods_gubun")
from_date = request("from_date")
to_date = request("to_date")
stock = request("stock")

curr_date = datevalue(mid(cstr(now()),1,10))
'st_date = cstr(mid(curr_date,1,4)) + "-" + "01" + "-" + "01"
st_date = "2015" + "-" + "01" + "-" + "01" '자재관리 초기재고 등록 기준..추후 매년 이월을 하면 매년 1월1일 기준으로 바꿀것
be_from_date = cstr(mid(from_date,1,4)) + cstr(mid(from_date,6,2)) + cstr(mid(from_date,9,2))
be_to_date = cstr(mid(to_date,1,4)) + cstr(mid(to_date,6,2)) + cstr(mid(to_date,9,2))

title_line = stock + " < " + goods_type + " > 입출고 현황 " + "(" + from_date + " ~ " + to_date + ")"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_out = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY stock_company,stock_goods_type,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_code ASC"

if goods_type = "전체" then 
         where_sql = " WHERE (a.stock_company = '"&view_condi&"')" 
   else
         where_sql = " WHERE (a.stock_company = '"&view_condi&"') and (a.stock_goods_type = '"&goods_type&"')" 
end if

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (a.stock_name like '%"&stock&"%') "
end if

if goods_gubun = "" then
       gubun_sql = ""
   else
       gubun_sql = " and (a.stock_goods_gubun like '%"&goods_gubun&"%') "
end if

' 입출고 및 재고수량이 없는것
cnt_sql = " and (a.stock_in_qty > 0 or a.stock_go_qty > 0 or a.stock_JJ_qty > 0) "

yes_goods_sql = " and (a.stock_goods_code = b.goods_code AND b.goods_used_sw = 'Y') "

sql = "select * from met_stock_gmaster a, met_goods_code b " + where_sql + stock_sql + gubun_sql + cnt_sql + yes_goods_sql + order_sql 
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>상품자재관리 시스템</title>
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
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
            	<table border="0" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr bgcolor="#EFEFEF" class="style1">
								<th class="first" scope="col">회사</th>
                                <th scope="col">용도구분</th>
                                <th scope="col">품목코드</th>
				                <th scope="col">품목구분</th>
                                <th scope="col">품목명</th>
                                <th scope="col">규격</th>
                                <th scope="col">Part_No.</th>
                                <th scope="col">상태</th>
                                <th scope="col">창고</th>
                                <th scope="col">전 재고</th>
                                <th scope="col">입고</th>
                                <th scope="col">출고</th>
                                <th scope="col">현 재고</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>     
                    <%
						do until rs.eof
                           stock_level = rs("stock_level")
						   stock_company = rs("stock_company")
						   stock_code = rs("stock_code")
						   stock_goods_type = rs("stock_goods_type")
						   stock_goods_code = rs("stock_goods_code")
						   stock_last_qty = rs("stock_last_qty")
						   stock_jj_qty = rs("stock_JJ_qty")
						   
						   sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
                           Set Rs_good = DbConn.Execute(SQL)
                           if not Rs_good.eof then
    	                         part_number = Rs_good("part_number")
                              else
		                         goods_model = ""
                           end if
                           Rs_good.close()

                           be_in_cnt = 0	 
						   be_out_cnt = 0
						   be_jae_cnt = 0
						   
						   sum_in_cnt = 0	 
						   sum_out_cnt = 0
						   sum_jae_cnt = 0
						   
'입고찾기- 본사구매입고	/ 직접 영업및ce쪽으로 입고시킨것									   
						   sql = "select * from met_stin_goods where (stin_stock_company = '"&stock_company&"') and (stin_stock_code = '"&stock_code&"') and (stin_goods_type = '"&stock_goods_type&"') and (stin_goods_code = '"&stock_goods_code&"') and (stin_date >= '"+st_date+"' and stin_date <= '"+to_date+"') ORDER BY stin_date ASC"
			   
'			               response.write sql
			   
	                       Rs_in.Open Sql, Dbconn, 1
                           while not Rs_in.eof
							  
							  be_stin_date = cstr(mid(Rs_in("stin_date"),1,4)) + cstr(mid(Rs_in("stin_date"),6,2)) + cstr(mid(Rs_in("stin_date"),9,2))
							  if be_stin_date < be_from_date then
							         be_in_cnt = be_in_cnt + int(Rs_in("stin_qty"))
							  end if
							  
							  if be_stin_date >= be_from_date and be_stin_date <= be_to_date then
							         sum_in_cnt = sum_in_cnt + int(Rs_in("stin_qty"))
							  end if
							  
						   	  Rs_in.movenext()
                           Wend
                           Rs_in.close()	
						   
'입고찾기- 본사 또는 개인이 ce쪽으로 출고한것(창고이동 입고건)							   
						   sql = "select * from met_mv_in_goods where (mvin_in_stock = '"&stock_code&"') and (in_goods_type = '"&stock_goods_type&"') and (in_goods_code = '"&stock_goods_code&"') and (mvin_in_date >= '"+st_date+"' and mvin_in_date <= '"+to_date+"') ORDER BY mvin_in_date ASC"
						   
						   Rs_in.Open Sql, Dbconn, 1
                           while not Rs_in.eof
							  
							  be_stin_date = cstr(mid(Rs_in("mvin_in_date"),1,4)) + cstr(mid(Rs_in("mvin_in_date"),6,2)) + cstr(mid(Rs_in("mvin_in_date"),9,2))
							  if be_stin_date < be_from_date then
							         be_in_cnt = be_in_cnt + int(Rs_in("in_qty"))
							  end if
							  
							  if be_stin_date >= be_from_date and be_stin_date <= be_to_date then
							         sum_in_cnt = sum_in_cnt + int(Rs_in("in_qty"))
							  end if
							  
						   	  Rs_in.movenext()
                           Wend
                           Rs_in.close()						   
						   
'출고찾기						   
						   sql = "select * from met_chulgo_goods where  (chulgo_stock_company = '"&stock_company&"') and (chulgo_stock = '"&stock_code&"') and (cg_goods_type = '"&stock_goods_type&"') and (cg_goods_code = '"&stock_goods_code&"') and (chulgo_date >= '"&st_date&"' and chulgo_date <= '"&to_date&"') and (cg_type <> '고객출고') ORDER BY chulgo_date ASC"
						   
						   Rs_out.Open Sql, Dbconn, 1
                           while not Rs_out.eof
						      
							  be_chulgo_date = cstr(mid(Rs_out("chulgo_date"),1,4)) + cstr(mid(Rs_out("chulgo_date"),6,2)) + cstr(mid(Rs_out("chulgo_date"),9,2))
							  
							  if be_chulgo_date < be_from_date then
							         be_out_cnt = be_out_cnt + int(Rs_out("cg_qty"))
							  end if
							  
							  if be_chulgo_date >= be_from_date and be_chulgo_date <= be_to_date then
							         sum_out_cnt = sum_out_cnt + int(Rs_out("cg_qty"))
							  end if

						   	  Rs_out.movenext()
                           Wend
                           Rs_out.close()
						   
						   be_jae_cnt = stock_last_qty + be_in_cnt - be_out_cnt
						   sum_jae_cnt = be_jae_cnt + sum_in_cnt - sum_out_cnt

					%>
				      <tr valign="middle" class="style1">
				        <td align="center"><%=rs("stock_company")%>&nbsp;</td>
                        <td align="center"><%=rs("stock_goods_type")%>&nbsp;</td>
                        <td align="center"><%=rs("stock_goods_code")%>&nbsp;</td>
                        <td align="center"><%=rs("stock_goods_gubun")%>&nbsp;</td>
                        <td align="center"><%=rs("stock_goods_name")%>&nbsp;</td>
                        <td align="center"><%=rs("stock_goods_standard")%>&nbsp;</td>
                        <td align="center"><%=part_number%>&nbsp;</td>
                        <td align="center"><%=rs("stock_goods_grade")%>&nbsp;</td>
                        <td align="center"><%=rs("stock_name")%>&nbsp;</td>
                        <td align="right"><%=formatnumber(be_jae_cnt,0)%>&nbsp;</td>
                        <td align="right"><%=formatnumber(sum_in_cnt,0)%>&nbsp;</td>
                        <td align="right"><%=formatnumber(sum_out_cnt,0)%>&nbsp;</td>
                        <td align="right"><%=formatnumber(sum_jae_cnt,0)%>&nbsp;</td>
                        <td>&nbsp;</td>
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