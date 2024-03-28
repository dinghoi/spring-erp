<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

from_date=Request("from_date")
to_date=Request("to_date")

savefilename = "CE별 실적 세부 내역 " + from_date + "_" + to_date + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select ce_work.*,as_acpt.acpt_date,as_acpt.acpt_user from as_acpt inner join ce_work on as_acpt.acpt_no=ce_work.acpt_no where (ce_work.work_id='2') and (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') Order By as_acpt.acpt_no, as_acpt.acpt_date Asc"
Rs.Open Sql, Dbconn, 1

title_line = from_date + " ~ " + to_date + " CE별 실적 세부 내역"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title></title>
<style type="text/css">
<!--
.style12 {font-size: 12px; font-family: "굴림체", "돋움체", Seoul; }
.style12B {font-size: 12px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; }
.style12BW {font-size: 12px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; color: #FFFFFF; }
.style14 {font-size: 14px; font-family: "굴림체", "돋움체", Seoul; }
.style14B {font-size: 14px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; }
.style14BW {font-size: 14px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; color: #FFFFFF; }
.style11 {font-size: 11px; font-family: "굴림체", "돋움체", Seoul; }
.style11B {font-size: 11px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; }
.style_red {color: #FF0000; font-weight: bold}
.style_green {color: #006600; font-weight: bold}
.style_blue {color: #000099; font-weight: bold}
-->
</style>
</head>

<body>
<span class="style14B"><%=title_line%></span>
<table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
    <tr valign="middle" class="style12">
    	<td><div align="center">순번</div></td>
    	<td><div align="center">접수번호</div></td>
    	<td><div align="center">접수일</div></td>
    	<td><div align="center">고객명</div></td>
    	<td><div align="center">처리일</div></td>
    	<td><div align="center">CE</div></td>
    	<td><div align="center">CE명</div></td>
    	<td><div align="center">처리회사</div></td>
    	<td><div align="center">처리지원회사</div></td>
    	<td><div align="center">처리유형</div></td>
    	<td><div align="center">처리건수</div></td>
    	<td><div align="center">팀</div></td>
    	<td><div align="center">상주구분</div></td>
    	<td><div align="center">상주처</div></td>
    	<td><div align="center">상주회사</div></td>
    	<td><div align="center">상주지원회사</div></td>
    	<td><div align="center">휴일근무</div></td>
    	<td><div align="center">타사지원</div></td>
	</tr>
<%
i = 0
do until rs.eof
	i = i + 1
	if rs("reside") = "1" then
		reside_view = "상주"
	  else
		reside_view = ""
	end if
	if rs("team") = "" or isnull(rs("team")) then
		org_view = rs("org_name")
	  else
	  	org_view = rs("team")
	end if
	if rs("holiday_yn") = "Y" then
		holiday_yn_view = "휴일근무"
	  else
	  	holiday_yn_view = ""
	end if

	sql_emp = "select * from memb where user_id ='"&rs("mg_ce_id")&"'"
	Set rs_emp = Dbconn.Execute (sql_emp)
	if rs_emp.eof or rs_emp.bof then
		user_name = "ERROR"
	  else
	  	user_name = rs_emp("user_name")
	end if

	tasa_pro = ""
	if rs("reside") = "1" and rs("reside_company") > "" and rs("as_type") <> "원격처리" then
		sql_trade = "select support_company from trade where trade_id ='매출' and trade_name = '"&rs("company")&"'"
		Set rs_trade = Dbconn.Execute (sql_trade)
		if rs_trade.eof or rs_trade.bof then
			company1 = rs("company")
			support_view1 = "ERROR"
		  else
			support_view1 = rs_trade("support_company")
			if rs_trade("support_company") = "없음" then
				company1 = rs("company")
			  else												
				company1 = rs_trade("support_company")
			end if
		end if
		rs_trade.close()
	
		sql_trade = "select support_company from trade where trade_id ='매출' and trade_name = '"&rs("reside_company")&"'"
		Set rs_trade = Dbconn.Execute (sql_trade)
		if rs_trade.eof or rs_trade.bof then
			company2 = rs("reside_company")
			support_view2 = "ERROR"
		  else
			support_view2 = rs_trade("support_company")
			if rs_trade("support_company") = "없음" then
				company2 = rs("reside_company")
			  else												
				company2 = rs_trade("support_company")
			end if
		end if
		rs_trade.close()
	
		if company1 <> company2 then
			tasa_pro = "타사처리"
		end if
	end if
%>
    <tr valign="middle" class="style12">
    	<td><div align="center"><%=i%></div></td>
    	<td><div align="center"><%=rs("acpt_no")%></div></td>
    	<td><div align="center"><%=rs("acpt_date")%></div></td>
    	<td><div align="center"><%=rs("acpt_user")%></div></td>
    	<td><div align="center"><%=rs("work_date")%></div></td>
    	<td><div align="center"><%=rs("mg_ce_id")%></div></td>
    	<td><div align="center"><%=user_name%></div></td>
    	<td><div align="center"><%=rs("company")%></div></td>
    	<td><div align="center"><%=support_view1%></div></td>
    	<td><div align="center"><%=rs("as_type")%></div></td>
    	<td><div align="center"><%=rs("person_amt")%></div></td>
    	<td><div align="center"><%=org_view%></div></td>
    	<td><div align="center"><%=reside_view%>&nbsp;</div></td>
    	<td><div align="center"><%=rs("reside_place")%>&nbsp;</div></td>
    	<td><div align="center"><%=rs("reside_company")%>&nbsp;</div></td>
    	<td><div align="center"><%=support_view1%></div></td>
    	<td><div align="center"><%=holiday_yn_view%>&nbsp;</div></td>
    	<td><div align="center"><%=tasa_pro%>&nbsp;</div></td>
	</tr>
<%
	rs.movenext()
loop
%>
</table>
</body>
</html>
<%
dbconn.Close()
Set dbconn = Nothing
%>
