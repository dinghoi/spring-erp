<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim Rs_stay
Dim stay_name

car_no=Request("car_no")
from_date=Request("from_date")
to_date=Request("to_date")
view_condi = car_no
	
curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = " 차량 비용 현황" + cstr(from_date) + " ~ " + cstr(to_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_drv = Server.CreateObject("ADODB.Recordset")
Set Rs_insu = Server.CreateObject("ADODB.Recordset")
Set Rs_pen = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'sql = "select * from car_info where car_no = '"&view_condi&"'"
sql = "select * from car_info ORDER BY car_owner,car_no ASC"
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
<table  border="1" cellpadding="0" cellspacing="0">
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;차량 현황&nbsp;<%=curr_date%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">차량번호</div></td>
    <td><div align="center" class="style1">차종</div></td>
    <td><div align="center" class="style1">연식</div></td>
    <td><div align="center" class="style1">유류종류</div></td>
    <td><div align="center" class="style1">소유구분</div></td>
    <td><div align="center" class="style1">차량소유회사</div></td>
    <td><div align="center" class="style1">사용부서</div></td>
    <td><div align="center" class="style1">용도</div></td>
    <td><div align="center" class="style1">운행자</div></td>
    <td><div align="center" class="style1">차량등록일</div></td>
    <td><div align="center" class="style1">운행Km</div></td>
    
    <td><div align="center" class="style1">보험가입일</div></td>
    <td><div align="center" class="style1">보험만기일</div></td>
    <td><div align="center" class="style1">보험회사</div></td>
    <td><div align="center" class="style1">보험료</div></td>
    <td><div align="center" class="style1">대인1</div></td>
    <td><div align="center" class="style1">대인2</div></td>
    <td><div align="center" class="style1">대물</div></td>
    <td><div align="center" class="style1">자기보험</div></td>
    <td><div align="center" class="style1">무상해</div></td>
    <td><div align="center" class="style1">자차</div></td>
    <td><div align="center" class="style1">연령</div></td>
    <td><div align="center" class="style1">긴급출동</div></td>
    <td><div align="center" class="style1">계약내용</div></td>
    
    <td><div align="center" class="style1">주유금액</div></td>
    <td><div align="center" class="style1">주차비</div></td>
    <td><div align="center" class="style1">통행료</div></td>
    
    <td><div align="center" class="style1">A/S금액</div></td>
    
    <td><div align="center" class="style1">과태료</div></td>
    <td><div align="center" class="style1">과태료납부</div></td>
    <td><div align="center" class="style1">과태료미납</div></td>
    <td><div align="center" class="style1">처분일자</div></td>
    <td><div align="center" class="style1">차량정보</div></td>
    <%' 아래부분은 일단 막아놓구... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">입고 세부내역 </div> %>
    <%'</div></td> %>
  </tr>
    <%
		do until rs.eof 
           view_condi = rs("car_no")
		   
		   sql="select max(ins_date) as max_ins_date,max(ins_last_date) as max_ins_last_date,max(ins_company) as max_ins_company,max(ins_amount) as max_ins_amount,max(ins_man1) as max_ins_man1,max(ins_man2) as max_ins_man2,max(ins_object) as max_ins_object,max(ins_self) as max_ins_self,max(ins_injury) as max_ins_injury,max(ins_self_car) as max_ins_self_car,max(ins_age) as max_ins_age,max(ins_comment) as max_ins_comment,max(ins_contract_yn) as max_ins_contract_yn,max(ins_scramble) as max_ins_scramble from car_insurance where ins_car_no = '"+view_condi+"'"
	       set rs_max=dbconn.execute(sql)
		
			if	isnull(rs_max("max_ins_date"))  then
			        ins_date = ""
					ins_last_date = ""
					ins_company = ""
			        ins_amount = 0
			        ins_man1 = ""
					ins_man2 = ""
					ins_object = ""
					ins_self = ""
					ins_injury = ""
					ins_self_car = ""
					ins_age = ""
					ins_scramble = ""
					ins_contract_yn = ""
					ins_comment = ""
		        else
					ins_date = rs_max("max_ins_date")
			        ins_last_date = rs_max("max_ins_last_date")
					ins_company = rs_max("max_ins_company")
			        ins_amount = clng(rs_max("max_ins_amount"))
			        ins_man1 = rs_max("max_ins_man1")
					ins_man2 = rs_max("max_ins_man2")
					ins_object = rs_max("max_ins_object")
					ins_self = rs_max("max_ins_self")
					ins_injury = rs_max("max_ins_injury")
					ins_self_car = rs_max("max_ins_self_car")
					ins_age = rs_max("max_ins_age")
					ins_scramble = rs_max("max_ins_scramble")
					ins_contract_yn = rs_max("max_ins_contract_yn")
					ins_comment = rs_max("max_ins_comment")
		    end if
			rs_max.close()	
			
					tot_fare = 0
                    tot_oil_price = 0
					tot_parking = 0
                    tot_toll = 0
                    sql = "select * from transit_cost where car_no = '"&view_condi&"' and run_date >= '"+from_date+"' and run_date <= '"+to_date+"' ORDER BY car_no,run_date,run_seq ASC"
					Rs_drv.Open Sql, Dbconn, 1
                    do until Rs_drv.eof
                              tot_fare = tot_fare + int(Rs_drv("fare"))
	                          tot_oil_price = tot_oil_price + int(Rs_drv("oil_price"))
							  tot_parking = tot_parking + int(Rs_drv("parking"))
							  tot_toll = tot_toll + int(Rs_drv("toll"))
	                    Rs_drv.movenext()
                    loop
                    Rs_drv.close()	
					
					tot_as = 0
                    sql = "select * from car_as where as_car_no = '"&view_condi&"' and as_date >= '"+from_date+"' and as_date <= '"+to_date+"' ORDER BY as_car_no,as_date,as_seq ASC"
					Rs_as.Open Sql, Dbconn, 1
                    do until Rs_as.eof
                             tot_as = tot_as + int(Rs_as("as_amount"))
	                   Rs_as.movenext()
                    loop
                    Rs_as.close()	
					
					tot_pe_amount = 0
                    tot_in_amt = 0
                    sql = "select * from car_penalty where pe_car_no = '"&view_condi&"' and pe_date >= '"+from_date+"' and pe_date <= '"+to_date+"' ORDER BY pe_car_no,pe_date,pe_seq ASC"
					Rs_pen.Open Sql, Dbconn, 1
                    do until Rs_pen.eof
                             tot_pe_amount = tot_pe_amount + int(Rs_pen("pe_amount"))
	                         tot_in_amt = tot_in_amt + int(Rs_pen("pe_in_amt"))
	                   Rs_pen.movenext()
                    loop
                    Rs_pen.close()	
					jan_amount = tot_pe_amount - tot_in_amt
	
	%>
  <tr valign="middle" class="style11">
      <% 'response.write(ins_man1)
	   'response.End %>
    <td width="115"><div align="center" class="style1"><%=rs("car_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("car_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_year")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("oil_kind")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_owner")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_company")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_use_dept")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("car_use")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("owner_emp_name")%>(<%=rs("owner_emp_no")%>)&nbsp;</div></td>
    <td width="145"><div align="center" class="style1"><%=rs("car_reg_date")%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(rs("last_km"),0)%></div></td>
    
    <td width="145"><div align="center" class="style1"><%=ins_date%></div></td>
    <td width="145"><div align="center" class="style1"><%=ins_last_date%></div></td>
    <td width="145"><div align="center" class="style1"><%=ins_company%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(ins_amount,0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=ins_man1%></div></td>
    <td width="115"><div align="center" class="style1"><%=ins_man2%></div></td>
    <td width="115"><div align="center" class="style1"><%=ins_object%></div></td>
    <td width="145"><div align="center" class="style1"><%=ins_self%></div></td>
    <td width="145"><div align="center" class="style1"><%=ins_injury%></div></td>
    <td width="145"><div align="center" class="style1"><%=ins_self_car%></div></td>
    <td width="145"><div align="center" class="style1"><%=ins_age%></div></td>
    <td width="145"><div align="center" class="style1"><%=ins_scramble%></div></td>
<% if ins_date = "" then %>
    <td width="145"><div align="left" class="style1">&nbsp;</div></td>
<%    else	
      if ins_contract_yn = "Y" then %>
   <td width="145"><div align="left" class="style1">계약내용포함</div></td>
<%    else %>
   <td width="145"><div align="left" class="style1">계약내용미포함(<%=ins_comment%>)</div></td>
<%    end if 
   end if %>   
   <td width="115"><div align="right" class="style1"><%=formatnumber(tot_oil_price,0)%></div></td>
   <td width="115"><div align="right" class="style1"><%=formatnumber(tot_parking,0)%></div></td>
   <td width="115"><div align="right" class="style1"><%=formatnumber(tot_toll,0)%></div></td>
   
   <td width="115"><div align="right" class="style1"><%=formatnumber(tot_as,0)%></div></td>
   
   <td width="115"><div align="right" class="style1"><%=formatnumber(tot_pe_amount,0)%></div></td>
   <td width="115"><div align="right" class="style1"><%=formatnumber(tot_in_amt,0)%></div></td>
   <td width="115"><div align="right" class="style1"><%=formatnumber(jan_amount,0)%></div></td>
   
   <td width="145"><div align="center" class="style1"><%=rs("end_date")%></div></td>
   <td width="145"><div align="center" class="style1"><%=rs("car_comment")%></div></td>
    <% 'response.write(ins_man1)
	   'response.End %>
  </tr>
	<%
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
