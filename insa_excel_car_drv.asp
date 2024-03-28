<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

view_condi=Request("view_condi")
from_date=Request("from_date")
to_date=Request("to_date")
	
curr_date = datevalue(mid(cstr(now()),1,10))

title_line = cstr(from_date) + "~ " + cstr(to_date) + " " + " 차량 운행현황"

savefilename = title_line +".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_drv = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
   Sql = "select * from transit_cost where run_date >= '"+from_date+"' and run_date <= '"+to_date+"' "
   else  
   Sql = "select * from transit_cost where car_no = '"+view_condi+"' and run_date >= '"+from_date+"' and run_date <= '"+to_date+"' "
end If
'//2017-09-07 정렬순서 변경
'Sql = Sql & " ORDER BY car_no,run_date,run_seq DESC"
Sql = Sql & " ORDER BY car_no,run_date,run_seq ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
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
<table  border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="17" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">차량번호</div></td>
    <td><div align="center" class="style1">차종</div></td>
    
    <td><div align="center" class="style1">운행일자</div></td>
    <td><div align="center" class="style1">운행자</div></td>
    <td><div align="center" class="style1">구분</div></td>
    <td><div align="center" class="style1">유종/대중교통</div></td>
    <td><div align="center" class="style1">출발업체명</div></td>
    <td><div align="center" class="style1">출발지</div></td>
    <td><div align="center" class="style1">출발KM</div></td>
    <td><div align="center" class="style1">도착업체명</div></td>
    <td><div align="center" class="style1">도착지</div></td>
    <td><div align="center" class="style1">도착KM</div></td>
    <td><div align="center" class="style1">운행목적</div></td>
    <td><div align="center" class="style1">대중교통경비</div></td>
    <td><div align="center" class="style1">주유금액</div></td>
    <td><div align="center" class="style1">주차비</div></td>
    <td><div align="center" class="style1">통행료</div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
  </tr>
    <%
		do until rs.eof 
          
		if rs("car_owner") <> "대중교통" then 
		   car_no = rs("car_no")
							  
		   Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
           Set rs_car = DbConn.Execute(SQL)
		   if not rs_car.eof then
				car_name = rs_car("car_name")
	          else
			    car_name = ""
           end if
           rs_car.close()
		   
		   emp_no = rs("mg_ce_id")
		   Sql = "select * from emp_master where emp_no = '"+emp_no+"'"
	       Set Rs_emp = DbConn.Execute(SQL)
	       if not Rs_emp.EOF or not Rs_emp.BOF then
			      drv_owner_emp_name = rs_emp("emp_name")
              else
                  drv_owner_emp_name = emp_no
			end if
			Rs_emp.close()
						
			if rs("start_km") = "" or isnull(rs("start_km")) then
					start_view = 0
			  else
				  	start_view = rs("start_km")
			end if
			if rs("end_km") = "" or isnull(rs("end_km")) then
					end_view = 0
			  else
				  	end_view = rs("end_km")
			end if
			run_km = rs("far")
	%>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=rs("car_no")%></div></td>
    <td width="115"><div align="center" class="style1"><%=car_name%></div></td>
<% 'response.write(rs("run_date"))
	   'response.End %>    
    <td width="115"><div align="center" class="style1"><%=rs("run_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=drv_owner_emp_name%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_owner")%></div></td>
    <td width="115"><div align="center" class="style1">
<% if rs("car_owner") = "대중교통" then %>
	       <%=rs("transit")%>
<%   else	%>                                
	       <%=rs("oil_kind")%>
<% end if %>
    </td>  
<% 'response.write(rs("transit"))
	   'response.End %>          
    <td width="115"><div align="center" class="style1"><%=rs("start_company")%></div></td>
    <td width="200"><div align="left" class="style1"><%=rs("start_point")%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(start_view,0)%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("end_company")%></div></td>
    <td width="200"><div align="left" class="style1"><%=rs("end_point")%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(end_view,0)%></div></td>
    <td width="200"><div align="left" class="style1"><%=rs("run_memo")%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(rs("fare"),0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(rs("oil_price"),0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(rs("parking"),0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(rs("toll"),0)%></div></td>
<% 'response.write(rs("toll"))
	   'response.End %>
  </tr>
<%
      end if
	    Rs.MoveNext()
	loop
%>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
