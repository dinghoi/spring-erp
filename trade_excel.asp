<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

use_sw=Request("use_sw")
view_condi=Request("view_condi")
condi=Request("condi")

if condi = "" then
	condi_view = "없음"
  else
  	condi_view = condi
end if

if use_sw = "Y" then
	use_view = "사용"
  elseif use_sw = "N" then
  	use_view = "미사용"
  else
  	use_view = "총괄"
end if 

title_line = "사용구분 : " + use_view + " , 조회조건 : " + condi_view + " - 거래처내역"
savefilename = cstr(now()) + " 거래처.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

if view_condi = "전체" and use_sw = "T" then
	where_sql = " "
  else
  	where_sql = " where "
end if

if view_condi = "전체" then
	condi_sql = " "
  else
	if condi = "" then
		condi_sql = view_condi + " = '" + condi + "'"
	  else
		condi_sql = view_condi + " like '%" + condi + "%'"
	end if
end if

if use_sw = "T" then
	use_sql = " "
  else
	if condi_sql = " " then
		use_sql = " use_sw = '" + use_sw + "'"
	  else
 		use_sql = " and use_sw = '" + use_sw + "'"
	end if
end if

Sql = "SELECT * FROM trade "&where_sql&condi_sql&use_sql&" ORDER BY trade_name ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<style type="text/css">
    <!--
    	.style10 {font-size: 10px; font-family: "굴림체", "굴림체", Seoul; }
        .style10B {font-size: 10px; font-weight: bold; font-family: "굴림체", "굴림체", Seoul; }
    -->
    </style>
		<title>비용 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr class="style10B">
								<th class="first" scope="col">순번</th>
								<th scope="col">코드</th>
								<th scope="col">사업자번호</th>
								<th scope="col">거래처명(FULL)</th>
								<th scope="col">거래처명</th>
								<th scope="col">거래처유형</th>
								<th scope="col">대표자</th>
								<th scope="col">주소</th>
								<th scope="col">업태</th>
								<th scope="col">업종</th>
								<th scope="col">전화</th>
								<th scope="col">팩스</th>
								<th scope="col">이메일</th>
								<th scope="col">담당자</th>
								<th scope="col">담당자전화</th>
								<th scope="col">관리그룹</th>
								<th scope="col">그룹명</th>
								<th scope="col">지원회사</th>
								<th scope="col">계산서발행거래처코드</th>
								<th scope="col">계산서발행거래처명</th>
								<th scope="col">사용유무</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
							trade_no = mid(rs("trade_no"),1,3) + "-" + mid(rs("trade_no"),4,2) + "-" + mid(rs("trade_no"),6) 
							sql_type="select * from type_code where etc_type='91' and etc_seq ='"+rs("mg_group")+"'"
							set rs_type=dbconn.execute(sql_type)
							mg_group = rs_type("type_name")
							rs_type.Close()		
							if rs("use_sw") = "Y" then
								view_use = "사용"
							  else
							  	view_use = "미사용"
							end if
						%>
							<tr class="style10">
								<td class="first"><%=i%></td>
								<td><%=rs("trade_code")%></td>
								<td><%=trade_no%></td>
								<td><%=rs("trade_full_name")%></td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("trade_id")%></td>
								<td><%=rs("trade_owner")%></td>
								<td><%=rs("trade_addr")%></td>
								<td><%=rs("trade_uptae")%></td>
								<td><%=rs("trade_upjong")%></td>
								<td><%=rs("trade_tel")%></td>
								<td><%=rs("trade_fax")%></td>
								<td><%=rs("trade_email")%></td>
								<td><%=rs("trade_person")%></td>
								<td><%=rs("trade_person_tel")%></td>
								<td><%=mg_group%></td>
								<td><%=rs("group_name")%></td>
								<td><%=rs("support_company")%></td>
								<td><%=rs("bill_trade_code")%></td>
								<td><%=rs("bill_trade_name")%></td>
								<td><%=use_view%></td>
							</tr>
					  	<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

