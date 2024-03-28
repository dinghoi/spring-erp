<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw

bill_month = request("bill_month")
slip_gubun = request("slip_gubun")
'view_c = Request("view_c")
view_d = Request("view_d")

from_date = mid(bill_month,1,4) + "-" + mid(bill_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

savefilename = bill_month + "월 세금계산서 내역.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

posi_sql = " and (emp_no = '"&user_id&"' or reg_id = '"&user_id&"') "
	
if position = "팀원" then
	view_condi = "본인"
end if
	
if position = "파트장" then
	if org_name = "한화생명호남" then
		posi_sql = " and (org_name = '한화생명호남' or org_name = '한화생명전북') "
	  else
		posi_sql = " and org_name = '"&org_name&"'"
	end if
end if
	
if position = "팀장" then
	posi_sql = " and team = '"&team&"'"
end if
	
if position = "사업부장" or cost_grade = "2" then
	posi_sql = " and saupbu = '"&saupbu&"'"
end if
	
if position = "본부장" or cost_grade = "1" then 
	posi_sql = " and bonbu = '"&bonbu&"'"
end if
	
view_grade = position
	
if cost_grade = "0" then
	posi_sql = ""
end if

if slip_gubun = "전체" then
	gubun_sql = ""
  else
  	gubun_sql = " and slip_gubun = '"&slip_gubun&"' "
end if

base_sql = "select * from general_cost where (tax_bill_yn = 'Y') "
if view_d = "slip" then
    base_sql = base_sql & " and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
    order_sql = " ORDER BY org_name, emp_name, slip_date ASC"
end if
if view_d = "reg" then
    base_sql = base_sql & " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59')"
    order_sql = " ORDER BY org_name, emp_name, reg_date ASC"
end if	

sql = base_sql + posi_sql + gubun_sql + order_sql
Rs.Open Sql, Dbconn, 1

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">회사</th>
								<th scope="col">본부</th>
								<th scope="col">사업부</th>
								<th scope="col">팀</th>
								<th scope="col">조직명</th>
								<th scope="col">상주처</th>
								<th scope="col">담당자</th>
								<th scope="col">발행일자</th>
								<th scope="col">발급일자</th>
								<th scope="col">고객사</th>
								<th scope="col">외주업체</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">비용유형</th>
								<th scope="col">비용구분</th>
								<th scope="col">세부유형</th>
								<th scope="col">발행내역</th>
								<th scope="col">손익</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							if rs("end_yn") = "Y" then
								end_yn = "마감"
								end_view = "N"
							  elseif rs("end_yn") = "I" then
								end_yn = "결재중"
								end_view = "N"
							  else
							  	end_yn = "진행"
							end if
							org_name = rs("emp_company") + "/" + rs("org_name")
							customer_no = mid(rs("customer_no"),1,3) + "-" + mid(rs("customer_no"),4,2) + "-" + mid(rs("customer_no"),6)
                            %>
                            <%
                            ' 5일 이후 지연 입력건 검출...
                            chk_slip_month = mid(rs("slip_date"),1,7)
                            chk_reg_month = mid(rs("reg_date"),1,7)
                            chk_reg_day = mid(rs("reg_date"),9,2)

                            if ((chk_slip_month < chk_reg_month) and (chk_reg_day > "05")) then
                                bgcolor = "burlywood"
                            else
                                bgcolor = "#f8f8f8"
                            end if
                            %>
                            <tr style="background-color: <%=bgcolor%>;">
                                <td class="first"><%=rs("emp_company")%></td>
                                <td><%=rs("bonbu")%></td>
                                <td><%=rs("saupbu")%></td>
                                <td><%=rs("team")%></td>
                                <td><%=rs("org_name")%></td>
                                <td><%=rs("reside_place")%></td>
                                <td><%=rs("emp_name")%></td>
                                <td><%=rs("slip_date")%></td>
                                <td><%=mid(rs("slip_date"),1,10)%></td>
                                <td><%=rs("company")%></td>
                                <td><%=rs("customer")%></td>
                                <td class="right"><%=formatnumber(rs("price"),0)%></td>
                                <td class="right"><%=formatnumber(rs("cost"),0)%></td>
                                <td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
                                <td><%=rs("cost_center")%></td>
                                <td><%=rs("slip_gubun")%></td>
                                <td><%=rs("account")%></td>
                                <td><%=rs("slip_memo")%></td>
                                <td><%=rs("pl_yn")%></td>
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

