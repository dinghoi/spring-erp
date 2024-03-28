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

title_line = " 창고이동 이동재고 현황 -- "+ stock +" (" + from_date + " ∼ " + to_date + ")"

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
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

mvin_id = "창고이동"

order_Sql = " ORDER BY chulgo_date,chulgo_stock,chulgo_seq DESC"

if view_condi = "전체" then
   if goods_type = "전체" then
          where_sql = " WHERE (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')" 
	  else
	      where_sql = " WHERE (chulgo_goods_type = '"&goods_type&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "전체" then
          where_sql = " WHERE (rele_stock_company = '"&view_condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')"
	  else
	      where_sql = " WHERE (chulgo_goods_type = '"&goods_type&"') and (rele_stock_company = '"&view_condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')"
   end if
end if   


if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (rele_stock_name like '%"&stock&"%') "
end if

sql = "select * from met_not_enter " + where_sql + stock_sql + order_sql

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
                               <th class="first" scope="col">순번</th>
                               <th scope="col">용도구분</th>
                               <th scope="col">출고일자</th>
                               <th scope="col">출고번호</th>
                               <th scope="col">출고상태</th>
                               <th scope="col">출고창고</th>
                               <th scope="col">출고담당</th>
                               <th scope="col">출고품목</th>
                               <th scope="col">신청일(No)</th>
                               <th scope="col">신청창고</th>
                               <th scope="col">신청담당</th>
                               <th scope="col">의뢰수량</th>
                               <th scope="col">출고수량</th>
							</tr>
						</thead>
						<tbody>
			<%
						i = 0
						do until rs.eof
                           i = i + 1
						   chulgo_date = rs("chulgo_date")
						   chulgo_stock = rs("chulgo_stock")
						   chulgo_seq = rs("chulgo_seq")
						   
						   rele_date = rs("rele_date")
						   rele_stock = rs("rele_stock")
						   rele_seq = rs("rele_seq")
					       
						   chulgo_no = mid(cstr(rs("chulgo_date")),3,2) + mid(cstr(rs("chulgo_date")),6,2) + mid(cstr(rs("chulgo_date")),9,2) 
						   rele_no = mid(cstr(rs("rele_date")),3,2) + mid(cstr(rs("rele_date")),6,2) + mid(cstr(rs("rele_date")),9,2) 
		    %>

                           <tr>
				              <td class="first"><%=i%></td>
                              <td><%=rs("chulgo_goods_type")%>&nbsp;</td>
                              <td><%=rs("chulgo_date")%>&nbsp;</td>
                              <td><%=chulgo_no%>&nbsp;<%=rs("chulgo_stock")%><%=rs("chulgo_seq")%></td>
                              <td><%=rs("chulgo_type")%>&nbsp;</td>
                              <td><%=rs("chulgo_stock_name")%>(<%=rs("chulgo_stock")%>)&nbsp;</td>
                              <td><%=rs("chulgo_emp_name")%>(<%=rs("chulgo_emp_no")%>)&nbsp;</td>
                              <td><%=rs("cg_goods_name")%>&nbsp;</td>
                              <td><%=rele_no%>&nbsp;<%=rs("rele_stock")%><%=rs("rele_seq")%></td>
                              <td><%=rs("rele_stock_name")%>&nbsp;</td>
                              <td><%=rs("rele_emp_name")%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("rl_qty"),0)%>&nbsp;</td>
                              <td align="right"><%=formatnumber(rs("cg_qty"),0)%>&nbsp;</td>
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
