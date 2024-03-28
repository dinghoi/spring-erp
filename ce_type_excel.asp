<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

from_date=Request("from_date")
to_date=Request("to_date")

savefilename = "CE별 실적 현황" + from_date + "_" + to_date + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select ce_work.mg_ce_id,memb.team,memb.org_name,memb.reside,memb.reside_place,memb.user_name from ce_work inner join memb on ce_work.mg_ce_id=memb.user_id where (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') GROUP BY ce_work.mg_ce_id,memb.team,memb.org_name,memb.reside,memb.reside_place,memb.user_name Order By memb.team, memb.user_name Asc"

Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" bgcolor="#6699CC">&nbsp;<span class="style14BW">CE별 실적 현황</span></td>
  </tr>
  <tr>
    <td><table border="0" cellspacing="3">
      <tr>
        <td>
          <table width="100%"  border="0">
            <tr>
              <td>&nbsp;<%=from_date%>&nbsp;~&nbsp;<%=to_date%></td>
            </tr>
            <tr>
              <td><table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                  <td rowspan="4"><div align="center">소속</div></td>
                  <td rowspan="4"><div align="center">CE</div></td>
                  <td rowspan="4"><div align="center">상주처</div></td>
				  <td colspan="52"><div align="center">당월 처리 유형 ( 전체수량/휴일 근무수량 )</div></td>
                </tr>
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                  <td colspan="4"><div align="center">계</div></td>
                  <td colspan="4"><div align="center">원격</div></td>
                  <td colspan="4"><div align="center">방문</div></td>
                  <td colspan="4"><div align="center">신규설치</div></td>
                  <td colspan="4"><div align="center">신규설치공사</div></td>
                  <td colspan="4"><div align="center">이전설치</div></td>
                  <td colspan="4"><div align="center">이전설치공사</div></td>
                  <td colspan="4"><div align="center">랜공사</div></td>
                  <td colspan="4"><div align="center">이전랜공사</div></td>
                  <td colspan="4"><div align="center">회수</div></td>
                  <td colspan="4"><div align="center">예방</div></td>
                  <td colspan="4"><div align="center">기타</div></td>
                  <td rowspan="3"><div align="center">입.완</div></td>
                  <td rowspan="3"><div align="center">입고</div></td>
                </tr>
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                  <td colspan="2">총수량</td>
                  <td colspan="2">야특근</td>
                </tr>
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                  <td>처리건</td>
                  <td>작업량</td>
                </tr>
						<% 
                        dim month_sum(13)
                        dim month_qty_sum(13)
                        dim month_tot(13)
                        dim month_qty_tot(13)
                        dim overtime_sum(13)
                        dim overtime_qty_sum(13)
                        dim overtime_tot(13)
                        dim overtime_qty_tot(13)
                        for i = 0 to 13
                            month_sum(i) = 0
                            month_qty_sum(i) = 0
                            month_tot(i) = 0
                            month_qty_tot(i) = 0
                            overtime_sum(i) = 0
                            overtime_qty_sum(i) = 0
                            overtime_tot(i) = 0
                            overtime_qty_tot(i) = 0
                        next
                
						ce_cnt = 0
                        do until rs.eof 
							ce_cnt = ce_cnt + 1
				' 월간 유형별 처리
                            sql = "select ce_work.as_type, holiday_yn, count(*) as end_cnt, sum(person_amt) as sum_cnt from ce_work  WHERE (ce_work.work_id='2') and (ce_work.mg_ce_id='"+rs("mg_ce_id")+"') and (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') GROUP BY ce_work.as_type,holiday_yn"		
                            rs_as.Open Sql, Dbconn, 1
                            do until rs_as.eof
                                select case rs_as("as_type")
                                    case "원격처리"
                                        month_sum(1) = month_sum(1) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(1) = month_qty_sum(1) + cint(rs_as("sum_cnt"))	
                                    case "방문처리"
                                        month_sum(2) = month_sum(2) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(2) = month_qty_sum(2) + cint(rs_as("sum_cnt"))	
                                    case "신규설치"
                                        month_sum(3) = month_sum(3) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(3) = month_qty_sum(3) + cint(rs_as("sum_cnt"))	
                                    case "신규설치공사"
                                        month_sum(4) = month_sum(4) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(4) = month_qty_sum(4) + cint(rs_as("sum_cnt"))	
                                    case "이전설치"
                                        month_sum(5) = month_sum(5) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(5) = month_qty_sum(5) + cint(rs_as("sum_cnt"))	
                                    case "이전설치공사"
                                        month_sum(6) = month_sum(6) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(6) = month_qty_sum(6) + cint(rs_as("sum_cnt"))	
                                    case "랜공사"
                                        month_sum(7) = month_sum(7) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(7) = month_qty_sum(7) + cint(rs_as("sum_cnt"))	
                                    case "이전랜공사"
                                        month_sum(8) = month_sum(8) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(8) = month_qty_sum(8) + cint(rs_as("sum_cnt"))	
                                    case "장비회수"
                                        month_sum(9) = month_sum(9) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(9) = month_qty_sum(9) + cint(rs_as("sum_cnt"))	
                                    case "예방점검"
                                        month_sum(10) = month_sum(10) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(10) = month_qty_sum(10) + cint(rs_as("sum_cnt"))	
                                    case "기타"
                                        month_sum(11) = month_sum(11) + cint(rs_as("end_cnt"))	
                                        month_qty_sum(11) = month_qty_sum(11) + cint(rs_as("sum_cnt"))	
                                end select												
								if rs_as("holiday_yn") = "Y" then
									select case rs_as("as_type")
										case "원격처리"
											overtime_sum(1) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(1) = cint(rs_as("sum_cnt"))	
										case "방문처리"
											overtime_sum(2) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(2) = cint(rs_as("sum_cnt"))	
										case "신규설치"
											overtime_sum(3) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(3) = cint(rs_as("sum_cnt"))	
										case "신규설치공사"
											overtime_sum(4) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(4) = cint(rs_as("sum_cnt"))	
										case "이전설치"
											overtime_sum(5) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(5) = cint(rs_as("sum_cnt"))	
										case "이전설치공사"
											overtime_sum(6) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(6) = cint(rs_as("sum_cnt"))	
										case "랜공사"
											overtime_sum(7) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(7) = cint(rs_as("sum_cnt"))	
										case "이전랜공사"
											overtime_sum(8) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(8) = cint(rs_as("sum_cnt"))	
										case "장비회수"
											overtime_sum(9) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(9) = cint(rs_as("sum_cnt"))	
										case "예방점검"
											overtime_sum(10) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(10) = cint(rs_as("sum_cnt"))	
										case "기타"
											overtime_sum(11) = cint(rs_as("end_cnt"))	
											overtime_qty_sum(11) = cint(rs_as("sum_cnt"))	
									end select												
								end if
                                rs_as.movenext()
                            loop
                            rs_as.close()
                ' 입고후 처리 완료
                            sql = "select count(*) as end_cnt from as_acpt "
                            sql = sql + "WHERE (in_date <> '') and (as_process='완료') and (mg_ce_id='"+rs("mg_ce_id")+"') and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"') GROUP BY mg_ce_id"		
							Set rs_as = Dbconn.Execute (sql)
							if rs_as.eof or rs_as.bof then
								month_sum(12) = 0
							  else
                                month_sum(12) = cint(rs_as("end_cnt"))	
							end if
							rs_as.close()
                ' 입고
                            sql = "select count(*) as end_cnt from as_acpt "
                            sql = sql + "WHERE (as_process='입고') and (mg_ce_id='"+rs("mg_ce_id")+"') and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"') GROUP BY mg_ce_id"		
                            rs_as.Open Sql, Dbconn, 1
							Set rs_as = Dbconn.Execute (sql)
							if rs_as.eof or rs_as.bof then
								month_sum(13) = 0
							  else
                                month_sum(13) = cint(rs_as("end_cnt"))	
							end if
							rs_as.close()
                
                            for i = 1 to 13
                                month_sum(0) = month_sum(0) + month_sum(i)
                                month_qty_sum(0) = month_qty_sum(0) + month_qty_sum(i)
 '                               month_tot(0) = month_tot(0) + month_tot(i)			
  '                              month_qty_tot(0) = month_qty_tot(0) + month_qty_tot(i)			
                                overtime_sum(0) = overtime_sum(0) + overtime_sum(i)
                                overtime_qty_sum(0) = overtime_qty_sum(0) + overtime_qty_sum(i)
'                                overtime_tot(0) = overtime_tot(0) + overtime_tot(i)			
'                                overtime_qty_tot(0) = overtime_qty_tot(0) + overtime_qty_tot(i)			
                            next
                            for i = 1 to 13
                                month_tot(i) = month_tot(i) + month_sum(i)			
                                overtime_tot(i) = overtime_tot(i) + overtime_sum(i)			
                                month_qty_tot(i) = month_qty_tot(i) + month_qty_sum(i)			
                                overtime_qty_tot(i) = overtime_qty_tot(i) + overtime_qty_sum(i)			
                            next
                
                            if month_sum(0) <> 0 then
								if rs("team") = "" or isnull(rs("team")) then
									org_view = rs("org_name") 
								  else
								  	org_view = rs("team")
								end if
	%>
                <tr class="style12">
                  <td><div align="center"><%=org_view%></div></td>
                  <td bgcolor="#FFFFCC"><div align="center"><%=rs("user_name")%></div></td>
                  <td><div align="center"><%=rs("reside_place")%></div></td>
			<% if company_cnt = 0 then	%>
                  <%   else	%>
                  <td bgcolor="#FFD8B0"><strong><%=company_cnt%></strong></td>
                  <td bgcolor="#FFD8B0"><strong><%=company_over%></strong></td>
            <% end if	%>
                  <td bgcolor="#CCCCCC"><%=formatnumber(clng(month_sum(0)),0)%></td>
                  <td bgcolor="#CCCCCC"><%=formatnumber(clng(month_qty_sum(0)),0)%></td>
                  <td bgcolor="#CCCCCC"><%=overtime_sum(0)%></td>
                  <td bgcolor="#CCCCCC"><%=overtime_qty_sum(0)%></td>
                  <td><%=formatnumber(clng(month_sum(1)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(1)),0)%></td>
                  <td><%=overtime_sum(1)%></td>
                  <td><%=overtime_qty_sum(1)%></td>
                  <td><%=formatnumber(clng(month_sum(2)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(2)),0)%></td>
                  <td><%=overtime_sum(2)%></td>
                  <td><%=overtime_qty_sum(2)%></td>
                  <td><%=formatnumber(clng(month_sum(3)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(3)),0)%></td>
                  <td><%=overtime_sum(3)%></td>
                  <td><%=overtime_qty_sum(3)%></td>
                  <td><%=formatnumber(clng(month_sum(4)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(4)),0)%></td>
                  <td><%=overtime_sum(4)%></td>
                  <td><%=overtime_qty_sum(4)%></td>
                  <td><%=formatnumber(clng(month_sum(5)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(5)),0)%></td>
                  <td><%=overtime_sum(5)%></td>
                  <td><%=overtime_qty_sum(5)%></td>
                  <td><%=formatnumber(clng(month_sum(6)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(6)),0)%></td>
                  <td><%=overtime_sum(6)%></td>
                  <td><%=overtime_qty_sum(6)%></td>
                  <td><%=formatnumber(clng(month_sum(7)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(7)),0)%></td>
                  <td><%=overtime_sum(7)%></td>
                  <td><%=overtime_qty_sum(7)%></td>
                  <td><%=formatnumber(clng(month_sum(8)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(8)),0)%></td>
                  <td><%=overtime_sum(8)%></td>
                  <td><%=overtime_qty_sum(8)%></td>
                  <td><%=formatnumber(clng(month_sum(9)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(9)),0)%></td>
                  <td><%=overtime_sum(9)%></td>
                  <td><%=overtime_qty_sum(9)%></td>
                  <td><%=formatnumber(clng(month_sum(10)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(10)),0)%></td>
                  <td><%=overtime_sum(10)%></td>
                  <td><%=overtime_qty_sum(10)%></td>
                  <td><%=formatnumber(clng(month_sum(11)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_sum(11)),0)%></td>
                  <td><%=overtime_sum(11)%></td>
                  <td><%=overtime_qty_sum(11)%></td>
                  <td><%=formatnumber(clng(month_sum(12)),0)%></td>
                  <td><%=formatnumber(clng(month_sum(13)),0)%></td>
                <%
			end if
			
			for i = 0 to 13
				month_sum(i) = 0
                overtime_sum(i) = 0
				month_qty_sum(i) = 0
                overtime_qty_sum(i) = 0
			next

			rs.movenext()
		loop
		rs.close()
		month_tot(0) = month_tot(1) + month_tot(2) + month_tot(3) + month_tot(4) + month_tot(5) + month_tot(6) + month_tot(7) + month_tot(8) + month_tot(9) + month_tot(10) + month_tot(11) + month_tot(12) + month_tot(13)
        overtime_tot(0) = overtime_tot(1) + overtime_tot(2) + overtime_tot(3) + overtime_tot(4) + overtime_tot(5) + overtime_tot(6) + overtime_tot(7) + overtime_tot(8) + overtime_tot(9) + overtime_tot(10) + overtime_tot(11) + overtime_tot(12) + overtime_tot(13)
		month_qty_tot(0) = month_qty_tot(1) + month_qty_tot(2) + month_qty_tot(3) + month_qty_tot(4) + month_qty_tot(5) + month_qty_tot(6) + month_qty_tot(7) + month_qty_tot(8) + month_qty_tot(9) + month_qty_tot(10) + month_qty_tot(11) + month_qty_tot(12) + month_qty_tot(13)
        overtime_qty_tot(0) = overtime_qty_tot(1) + overtime_qty_tot(2) + overtime_qty_tot(3) + overtime_qty_tot(4) + overtime_qty_tot(5) + overtime_qty_tot(6) + overtime_qty_tot(7) + overtime_qty_tot(8) + overtime_qty_tot(9) + overtime_qty_tot(10) + overtime_qty_tot(11) + overtime_qty_tot(12) + overtime_qty_tot(13)
		%>
                <tr valign="middle" bgcolor="#CCCCCC" class="style12B">
                  <td>총계</td>
                  <td><%=ce_cnt%></td>
                  <td>&nbsp;</td>
                  <td><%=formatnumber(clng(month_tot(0)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(0)),0)%></td>
                  <td><%=overtime_tot(0)%></td>
                  <td><%=overtime_qty_tot(0)%></td>
                  <td><%=formatnumber(clng(month_tot(1)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(1)),0)%></td>
                  <td><%=overtime_tot(1)%></td>
                  <td><%=overtime_qty_tot(1)%></td>
                  <td><%=formatnumber(clng(month_tot(2)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(2)),0)%></td>
                  <td><%=overtime_tot(2)%></td>
                  <td><%=overtime_qty_tot(2)%></td>
                  <td><%=formatnumber(clng(month_tot(3)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(3)),0)%></td>
                  <td><%=overtime_tot(3)%></td>
                  <td><%=overtime_qty_tot(3)%></td>
                  <td><%=formatnumber(clng(month_tot(4)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(4)),0)%></td>
                  <td><%=overtime_tot(4)%></td>
                  <td><%=overtime_qty_tot(4)%></td>
                  <td><%=formatnumber(clng(month_tot(5)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(5)),0)%></td>
                  <td><%=overtime_tot(5)%></td>
                  <td><%=overtime_qty_tot(5)%></td>
                  <td><%=formatnumber(clng(month_tot(6)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(6)),0)%></td>
                  <td><%=overtime_tot(6)%></td>
                  <td><%=overtime_qty_tot(6)%></td>
                  <td><%=formatnumber(clng(month_tot(7)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(7)),0)%></td>
                  <td><%=overtime_tot(7)%></td>
                  <td><%=overtime_qty_tot(7)%></td>
                  <td><%=formatnumber(clng(month_tot(8)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(8)),0)%></td>
                  <td><%=overtime_tot(8)%></td>
                  <td><%=overtime_qty_tot(8)%></td>
                  <td><%=formatnumber(clng(month_tot(9)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(9)),0)%></td>
                  <td><%=overtime_tot(9)%></td>
                  <td><%=overtime_qty_tot(9)%></td>
                  <td><%=formatnumber(clng(month_tot(10)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(10)),0)%></td>
                  <td><%=overtime_tot(10)%></td>
                  <td><%=overtime_qty_tot(10)%></td>
                  <td><%=formatnumber(clng(month_tot(11)),0)%></td>
                  <td><%=formatnumber(clng(month_qty_tot(11)),0)%></td>
                  <td><%=overtime_tot(11)%></td>
                  <td><%=overtime_qty_tot(11)%></td>
                  <td><%=formatnumber(clng(month_tot(12)),0)%></td>
                  <td><%=formatnumber(clng(month_tot(13)),0)%></td>
                </tr>
              </table></td>
            </tr>
          </table>
        </td>
      </tr>
    </table></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
dbconn.Close()
Set dbconn = Nothing
%>
