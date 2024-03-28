<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 

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
		elseif view_c = "grade" then
               field_name = "상태"
			   field_view = field_grade
			elseif view_c = "name" then
                   field_name = "품목명"
				   field_view = field_name
			   elseif view_c = "stand" then
                      field_name = "규격"
					  field_view = field_stand
				   elseif view_c = "code" then
                          field_name = "품목코드"
						  field_view = field_code
					   elseif view_c = "group" then
                              field_name = "품목분류"
							  field_view = field_group
end if   
  

title_line = " 품목코드 현황 -- "+ field_name +" (" + field_view + ")"

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

order_Sql = " ORDER BY goods_gubun,goods_name,goods_code ASC"
If field_check = "total" Then
       owner_sql = "select * FROM met_goods_code"
	   field_check = ""
   else
       if view_c = "gubun" Then
              owner_sql = "select * FROM met_goods_code  where goods_gubun like '%" + field_gubun + "%'"
       end if
	   if view_c = "grade" Then
              owner_sql = "select * FROM met_goods_code  where goods_grade like '%" + field_grade + "%'"
       end if
	   if view_c = "group" Then
              owner_sql = "select * FROM met_goods_code  where goods_group like '%" + field_group + "%'"
       end if
	   if view_c = "name" Then
              owner_sql = "select * FROM met_goods_code  where goods_name like '%" + field_name + "%'"
       end if
	   if view_c = "code" Then
              owner_sql = "select * FROM met_goods_code  where goods_code like '%" + field_code + "%'"
       end if
	   if view_c = "stand" Then
              owner_sql = "select * FROM met_goods_code  where goods_standard like '%" + field_stand + "%'"
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
								<th class="first" scope="col">품목코드</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목명</th>
                                <th scope="col">규격</th>
                                <th scope="col">모델</th>
                                <th scope="col">상태</th>
                                <th scope="col">Serial No.</th>
                                <th scope="col">Part_Number</th>
                                <th scope="col">PO_Number</th>
                                <th scope="col">상세설명</th>
                                <th scope="col">비고</th>
                                <th scope="col">등록일</th>
                                <th scope="col">적정재고</th>
                                <th scope="col">분류</th>
                                <th scope="col">사용</th>
							</tr>
						</thead>
						<tbody>
			<%
						do until rs.eof
						   if rs("goods_security_yn") = "Y" then 
						          goods_security = "대상" 
							  else
							      goods_security = "비대상"
						   end if
						   goods_end_date = rs("goods_end_date")
						   if goods_end_date = "1900-01-01" then
	                            goods_end_date = ""
	                       end if
		    %>
                                 <tr>
								    <td class="first"><%=rs("goods_code")%>&nbsp;</td>
                                    <td><%=rs("goods_gubun")%>&nbsp;</td>
                                    <td><%=rs("goods_name")%>&nbsp;</td>
                                    <td><%=rs("goods_standard")%>&nbsp;</td>
                                    <td><%=rs("goods_model")%>&nbsp;</td>
                                    <td><%=rs("goods_grade")%>&nbsp;</td>
                                    <td><%=rs("goods_serial_no")%>&nbsp;</td>
                                    <td><%=rs("part_number")%>&nbsp;</td>
                                    <td><%=rs("po_number")%>&nbsp;</td>
                                    <td><%=rs("goods_comment")%>&nbsp;</td>
                                    <td><%=rs("goods_bigo")%>&nbsp;</td>
                                    <td><%=rs("goods_date")%>&nbsp;</td>
              <% if rs("goods_security_yn") = "Y" then  %>
                                    <td align="right"><%=formatnumber(rs("goods_security_cnt"),0)%>&nbsp;</td>
              <%      else %>
                                    <td>&nbsp;</td>
              <% end if %>
                                    <td><%=rs("goods_group")%>&nbsp;</td>
                                    <td align="center"><%=rs("goods_used_sw")%>&nbsp;</td>
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
