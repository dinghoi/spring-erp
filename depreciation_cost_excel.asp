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

slip_month = request("slip_month")
slip_gubun = request("slip_gubun")

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

title_line = slip_month + "월 상각비 내역"
savefilename = slip_month + "월 상각비 내역.xls"

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

sql = "select * from general_cost where (slip_gubun = '상각비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') ORDER BY org_name, emp_name, slip_date ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
								<th class="first" scope="col">비용회사</th>
								<th scope="col">비용일자</th>
								<th scope="col">담당자</th>
								<th scope="col">금액</th>
								<th scope="col">상각비유형</th>
								<th scope="col">상각비 세부내역</th>
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
						%>
							<tr>
								<td class="first"><%=rs("emp_company")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("emp_name")%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
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

