<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

run_month=Request("run_month")
transit_id = Request("transit_id")
view_c = Request("view_c")
view_d = Request("view_d")
use_man = Request("use_man")

if run_month = "" then
	run_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	view_c = "total"
	emp_name = ""
end If

from_date = mid(run_month,1,4) + "-" + mid(run_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
sign_month = run_month

savefilename = run_month + "월 " + transit_id + " 교통비 현황.xls"


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 포지션별
posi_sql = " and transit_cost.mg_ce_id = '" + user_id + "'"
Response.write position

if position = "파트장" then
	if view_c = "total" then
		if org_name = "한화생명호남" then
			posi_sql = " and (transit_cost.org_name = '한화생명호남' or transit_cost.org_name = '한화생명전북') "
		  else
			posi_sql = " and transit_cost.org_name = '"&org_name&"'"
		end if
	  else
		if org_name = "한화생명호남" then
			posi_sql = " and (transit_cost.org_name = '한화생명호남' or transit_cost.org_name = '한화생명전북') and memb.user_name like '%"&use_man&"%'"
		  else
			posi_sql = " and transit_cost.org_name = '"&org_name&"' and memb.user_name like '%"&use_man&"%'"
		end if
	end if
end if

if position = "팀장" then
	if view_c = "total" then
		posi_sql = " and transit_cost.team = '"&team&"'"
	  else
		posi_sql = " and transit_cost.team = '"&team&"' and memb.user_name like '%"&use_man&"%'"
	end if
end if

if position = "사업부장" or cost_grade = "2" then
	if view_c = "total" then
		'posi_sql = " and transit_cost.saupbu = '"&saupbu&"'"
        posi_sql = " and transit_cost.saupbu = emp_master.emp_saupbu "&chr(13)
	else
        'posi_sql = " and transit_cost.saupbu = '"&saupbu&"' and memb.user_name like '%"&use_man&"%'"
        posi_sql = " and transit_cost.saupbu = emp_master.emp_saupbu and memb.user_name like '%"&use_man&"%' "&chr(13)
	end if
end if

if position = "본부장" or cost_grade = "1" then
  	if view_c = "total" then
		posi_sql = " and transit_cost.bonbu = '"&bonbu&"'"
 	  else
		posi_sql = " and transit_cost.bonbu = '"&bonbu&"' and memb.user_name like '%"&use_man&"%'"
	end if	 
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "전체"
  	if view_c = "total" then
		posi_sql = ""
 	  else
		posi_sql = " and memb.user_name like '%"&use_man&"%'"
	end if	 
end if

if transit_id = "차량" then
	transit_sql = " and (transit_cost.car_owner = '개인' or transit_cost.car_owner = '회사') "
  else
	transit_sql = " and (transit_cost.car_owner = '대중교통') "
end if

' 조건별 조회.........
base_sql = "    select *                                          "&chr(13)&_
           "      from transit_cost                               "&chr(13)&_
           "inner join (SELECT user_id, user_name FROM memb) memb "&chr(13)&_
           "        on transit_cost.mg_ce_id = memb.user_id       "&chr(13)&_
           "inner join emp_master                                 "&chr(13)&_           
           "        ON emp_master.emp_no = memb.user_id           "&chr(13)

if view_d = "run" then
    date_sql = " where (run_date >= '" + from_date  + "' and run_date <= '" + to_date  + "') "&chr(13)
    order_sql = " ORDER BY memb.user_name asc, run_date desc, run_seq desc"
end if
if view_d = "reg" then
    date_sql = " where (transit_cost.reg_date >= '" + from_date  + " 00:00:00' and transit_cost.reg_date <='" + to_date  + " 23:59:59') "&chr(13)
    order_sql = " ORDER BY memb.user_name asc, transit_cost.reg_date desc, run_seq desc"
end if

sql = base_sql + date_sql + posi_sql + transit_sql + order_sql
Rs.Open Sql, Dbconn, 1

'Response.write "<pre>"&Sql & "</pre><br>"

base_CntSql = "    select count(*)                                   "&chr(13)&_
              "      from transit_cost                               "&chr(13)&_
              "inner join (SELECT user_id, user_name FROM memb) memb "&chr(13)&_
              "        on transit_cost.mg_ce_id = memb.user_id       "&chr(13)&_
              "inner join emp_master                                 "&chr(13)&_           
              "        ON emp_master.emp_no = memb.user_id           "&chr(13)
base_CntSql = base_CntSql + date_sql + posi_sql + transit_sql + order_sql
Set RsCount = Dbconn.Execute (base_CntSql)

'Response.write "<pre>"&base_CntSql & "</pre><br>"

tottal_record = cint(RsCount(0))


if (tottal_record > 0) then
    Response.Buffer = True
    Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
    Response.CacheControl = "public"
    Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

    'Response.write tottal_record & "<br>"
    'Response.End
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
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th rowspan="2" scope="col" class="first">회사</th>
								<th rowspan="2" scope="col">본부</th>
								<th rowspan="2" scope="col">사업부</th>
								<th rowspan="2" scope="col">팀</th>
								<th rowspan="2" scope="col">조직명</th>
								<th rowspan="2" scope="col">상주처</th>
								<th rowspan="2" scope="col">사용회사</th>
								<th rowspan="2" scope="col">운행자</th>
								<th rowspan="2" scope="col">사번</th>
								<th rowspan="2" scope="col">운행일자</th>
								<th rowspan="2" scope="col">발급일자</th>
								<th rowspan="2" scope="col">비용유형</th>
								<th rowspan="2" scope="col">차량구분</th>
								<th rowspan="2" scope="col">출발지</th>
								<th rowspan="2" scope="col">도착지</th>
								<th rowspan="2" scope="col">운행목적</th>
								<th rowspan="2" scope="col">시작KM</th>
								<th rowspan="2" scope="col">도착KM</th>
								<th rowspan="2" scope="col">거리</th>
								<th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">경 비 </th>
								<th rowspan="2" scope="col">지급</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">수리비</th>
								<th scope="col">대중교통</th>
								<th scope="col">주유금액</th>
								<th scope="col">주차료</th>
								<th scope="col">통행료</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							if rs("car_owner") = "대중교통" then
								car_gubun = rs("transit")
							  else
								car_gubun = rs("car_owner") + " " + rs("oil_kind")
							end if
							run_km = rs("far")

							if rs("cancel_yn") = "Y" then
								cancel_view = "취소"
							  else
							  	cancel_view = "지급"
							end if														  
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
                            %>
                            <%
                            ' 5일 이후 지연 입력건 검출...
                            chk_slip_month = mid(rs("run_date"),1,7)
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
                                <td><%=rs("company")%></td>
                                <td><%=rs("user_name")%></td>
                                <td><%=rs("mg_ce_id")%></td>
                                <td><%=rs("run_date")%></td>
                                <td><%=mid(rs("reg_date"),1,10)%></td>
                                <td><%=rs("cost_center")%></td>
                                <td><%=car_gubun%></td>
                                <td><%=rs("start_company")%>-<%=rs("start_point")%></td>
                                <td><%=rs("end_company")%>-<%=rs("end_point")%></td>
                                <td><%=rs("run_memo")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(start_view,0)%></td>
                                <td class="right"><%=formatnumber(end_view,0)%></td>
                                <td class="right"><%=formatnumber(run_km,0)%></td>
                                <td class="right"><%=formatnumber(rs("repair_cost"),0)%></td>
                                <td class="right"><%=formatnumber(rs("fare"),0)%></td>
                                <td class="right"><%=formatnumber(rs("oil_price"),0)%></td>
                                <td class="right"><%=formatnumber(rs("parking"),0)%></td>
                                <td class="right"><%=formatnumber(rs("toll"),0)%></td>
                                <td><%=cancel_view%></td>
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

    <%
else
    %>
    <script>alert("데이터가 존재하지 않습니다.");</script>
    <%
end if
%>