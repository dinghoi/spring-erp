<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

slip_month=Request("slip_month")
		
from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
owner_company = Request("owner_company")
card_type = Request("card_type")
field_check = Request("field_check")
field_view = Request("field_view")

title_line = slip_month + " 카드 전표 관리"
savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select * from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"

if owner_company = "전체" then
	owner_company_sql = " "
  else
	owner_company_sql = " and ( owner_company = '" + owner_company + "' ) "
end if
if card_type = "전체" then
	card_type_sql = " "
  else
	card_type_sql = " and ( card_slip.card_type = '" + card_type + "' ) "
end if

if field_check <> "total" then
	field_sql = " and ( card_slip." + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if
order_sql = " ORDER BY slip_date ASC"

sql = base_sql + owner_company_sql + card_type_sql + field_sql + order_sql
response.write(sql)
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
	<style type="text/css">
    <!--
    	.style10 {font-size: 10px; font-family: "굴림체", "굴림체", Seoul; }
        .style10B {font-size: 10px; font-weight: bold; font-family: "굴림체", "굴림체", Seoul; }
    -->
    </style>
		<title>관리 회계 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="6%" >
							<col width="12%" >
							<col width="6%" >
							<col width="8%" >
							<col width="*" >
							<col width="2%" >
							<col width="7%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr class="style10B">
								<th class="first" scope="col">종류</th>
								<th scope="col">카드사명</th>
								<th scope="col">카드번호</th>
								<th scope="col">사용자</th>
								<th scope="col">승인일자</th>
								<th scope="col">사업자번호</th>
								<th scope="col">거래처명</th>
								<th scope="col">거래처유형</th>
								<th scope="col">공급가액</th>
								<th scope="col">세액</th>
								<th scope="col">봉사료</th>
								<th scope="col">합계금액</th>
								<th scope="col">부가세공제여부</th>
								<th scope="col">부가세유형</th>
								<th scope="col">계정과목</th>
								<th scope="col">품목</th>
								<th scope="col">수량</th>
								<th scope="col">단가</th>
								<th scope="col">대표자</th>
								<th scope="col">업태</th>
								<th scope="col">업종</th>
								<th scope="col">사업장주소</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						j = 0
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0
						err_cnt = 0
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							i = i + 1
							if rs("cost_vat") > 0 then
								vat_yn = "공제"
								vat_type = "57"
							  else
								vat_yn = "불공제"
								vat_type = "  "
							end if

							Sql="select * from account where account_name = '" + rs("account") + "'"
							Set rs_acc=DbConn.Execute(Sql)
							if rs_acc.eof or rs_acc.bof then
								account_code = "ERROR"
								err_cnt = err_cnt + 1
							  else
							  	account_code = rs_acc("account_code")
							end if
						%>
							<tr class="style10">
								<td class="first">3</td>
								<td><%=rs("card_type")%></td>
								<td><%=rs("card_no")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("customer_no")%></td>
								<td><%=rs("customer")%></td>
								<td></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td>&nbsp;</td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
								<td><%=vat_yn%></td>
								<td><%=vat_type%>&nbsp;</td>
								<td><%=account_code%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
							</tr>
					  <%
							rs.movenext()
						loop
						rs.close()
						if price_sum <> ( cost_sum + cost_vat_sum ) then
							err_msg = "금액확인 요망"
						  else
						  	err_msg = " "
						end if
						%>
							<tr class="style10B">
								<th colspan="2" class="first">총계</th>
								<th colspan="4"><%=i%>&nbsp;건</th>
								<td></td>
							  	<th><%=formatnumber(cost_sum,0)%></th>
								<th><%=formatnumber(cost_vat_sum,0)%></th>
							  	<th>0</th>
							  	<th><%=formatnumber(price_sum,0)%></th>
								<th colspan="3">
					<% if err_cnt > 0 then	%>
                    		계정과목 미분류가 <%=err_cnt%> 건 있습니다.
					<%   else	%>
                    			&nbsp;
                    <% end if	%>					
                                </th>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td>
                                </td>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

