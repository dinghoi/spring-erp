<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 

goods_type=request("goods_type")
field_check=request("field_check")
field_gubun=Request("field_gubun")
field_grade=Request("field_grade")
field_group=Request("field_group")
field_name=Request("field_name")
field_code=Request("field_code")
field_stand=Request("field_stand")
view_c = Request("view_c")

curr_date = datevalue(mid(cstr(now()),1,10))

If view_c = "" Then
	ck_sw = "n"
	field_check = "total"
	view_c = "gubun"
End If

if field_check = "total" then
     field_name = "전체"
	 field_view = "전체"
   elseif view_c = "gubun" then
            field_name = "품목구분"
			if field_gubun = "" then
			       field_view = "전체"
			   else
			       field_view = field_gubun
		    end if
		elseif view_c = "pono" then
               field_name = "PO_No."
			   field_view = field_grade
			elseif view_c = "name" then
                   field_name = "품목명"
				   field_view = field_name
			   elseif view_c = "partno" then
                      field_name = "Part_NO."
					  field_view = field_stand
				   elseif view_c = "code" then
                          field_name = "품목코드"
						  field_view = field_code
					   elseif view_c = "serial" then
                              field_name = "품목serial"
							  field_view = field_group
end if   
  

title_line = " Serial 관리 현황 -- "+ field_name +" (" + field_view + ")"

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

order_Sql = " ORDER BY goods_gubun,in_date,goods_code,serial_no,serial_seq ASC"

If field_check = "total" Then
       owner_sql = "select * FROM met_goods_serial"
	   field_check = ""
   else
       if view_c = "gubun" Then
              owner_sql = "select * FROM met_goods_serial  where goods_gubun like '%" + field_gubun + "%'"
       end if
	   if view_c = "pono" Then
              owner_sql = "select * FROM met_goods_serial  where po_number like '%" + field_grade + "%'"
       end if
	   if view_c = "serial" Then
              owner_sql = "select * FROM met_goods_serial  where serial_no like '%" + field_group + "%'"
       end if
	   if view_c = "name" Then
              owner_sql = "select * FROM met_goods_serial  where goods_name like '%" + field_name + "%'"
       end if
	   if view_c = "code" Then
              owner_sql = "select * FROM met_goods_serial  where goods_code like '%" + field_code + "%'"
       end if
	   if view_c = "partno" Then
              owner_sql = "select * FROM met_goods_serial  where part_number like '%" + field_stand + "%'"
       end if
End If


sql = owner_sql + order_sql
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
								<th class="first" scope="col">품목구분</th>
                                <th scope="col">Serial No</th>
                                <th scope="col">품목코드</th>
                                <th scope="col">품목명</th>
                                <th scope="col">Part_No./규격</th>
                                <th scope="col">입고일자</th>
                                <th scope="col">PO_No.</th>
                                <th scope="col">출고일자</th>
                                <th scope="col">납품처</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>
			<%
						do until rs.eof

		    %>
                                 <tr>
								    <td class="first"><%=rs("goods_gubun")%>&nbsp;</td>
                                    <td class="left"><%=rs("serial_no")%>&nbsp;<%=rs("serial_seq")%></td>
                                    <td><%=rs("goods_code")%>&nbsp;</td>
                                    <td class="left"><%=rs("goods_name")%>&nbsp;</td>
                                    <td class="left"><%=rs("part_number")%>&nbsp;</td>
                                    <td><%=rs("in_date")%>&nbsp;</td>
                                    <td class="left"><%=rs("po_number")%>&nbsp;</td>
                                    <td><%=rs("chulgo_date")%>&nbsp;</td>
                                    <td class="left"><%=rs("chulgo_trade")%>&nbsp;<%=rs("chulgo_trade_dept")%></td>
                                    <td class="left"><%=rs("chulgo_bigo")%>&nbsp;</td>
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
