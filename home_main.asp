<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/army_dbcon.asp" -->

<%
dim s_tab(12,3)
dim mon_tab(12)
dim group_tab(4,5)

for i = 0 to 12
	for j = 0 to 3
		s_tab(i,j) = 0
	next
next
for i = 0 to 4
	group_tab(i,1) = 0
	group_tab(i,2) = "."
	group_tab(i,3) = 0
	group_tab(i,4) = 0
	group_tab(i,5) = "."
next
group_tab(0,0) = "주전산기"
group_tab(1,0) = "1지역"
group_tab(2,0) = "2지역"
group_tab(3,0) = "3지역"
group_tab(4,0) = "합계"

user_name = request.cookies("army_user")("coo_user_name")
user_grade = request.cookies("army_user")("coo_user_grade")
mg_grade = request.cookies("army_user")("coo_mg_grade")
user_id = request.cookies("army_user")("coo_user_id")

Set dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs1 = Server.CreateObject("ADODB.Recordset")
Set Rs2 = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

be_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
	
'sql = "select max(sla_month) as max_month from sla_close where close_stat ='A'"
sql = "select max(sla_month) as max_month from sla_close"
set rs_max = Dbconn.execute(sql)
if isnull(rs_max("max_month")) or rs_max("max_month") = "" then
	sla_month = be_month
  else
	sla_month = rs_max("max_month")
end if	

mon_tab(12) = sla_month
sub_tit = mid(sla_month,1,4) + "년" + mid(sla_month,5,2) + "월 관리그룹별 SLA 평가현황"

cal_yy = cint(mid(sla_month,1,4))
cal_mm = cint(mid(sla_month,5,2))

for i = 1 to 11
	cal_mm = cal_mm - 1
	if cal_mm = 0 then
		cal_mm = 12
		cal_yy = cal_yy - 1
	end if
	if cal_mm < 10 then
		cal_month = cstr(cal_yy) + "0" + cstr(cal_mm)
	  else
		cal_month = cstr(cal_yy) + cstr(cal_mm)
	end if
	mon_tab(12-i) = cal_month
next

Sql = "SELECT * FROM sla_close where sla_month >= '" + mon_tab(1) + "' and sla_month <= '" + mon_tab(12) + "' order by sla_month, sla_group ASC"
Rs.Open Sql, Dbconn, 1

do until rs.eof
	for i = 1 to 12
		if mon_tab(i) = rs("sla_month") then
			j = int(rs("sla_group"))
			s_tab(i,j) = rs("sla_score")
			exit for
		end if
	next
	rs.movenext()
loop
rs.close()

sql = "select * from sla_close where sla_month = '" + sla_month + "'"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	i = int(rs("sla_group"))
	group_tab(i,1) = rs("sla_score")
	group_tab(i,2) = rs("sla_sum_grade")
	group_tab(i,3) = rs("reward_plus_cost")
	group_tab(i,4) = rs("reward_minus_cost")
	if rs("close_stat") = "E" then
		group_tab(i,5) = "입력마감"
	  elseif rs("close_stat") = "A" then
		group_tab(i,5) = "최종승인"
	  elseif rs("close_stat") = "C" then
		group_tab(i,5) = "마감취소"
	  else
		group_tab(i,5) = "기타"
	end if	  	
	group_tab(4,3) = group_tab(4,3) + rs("reward_plus_cost")		
	group_tab(4,4) = group_tab(4,4) + rs("reward_minus_cost")		
	rs.movenext()
loop
rs.close()


sql = "select * from sla_board  order by reg_date desc limit 0,5"
Rs1.Open Sql, Dbconn, 1

new_date = now() - 14
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>SLA 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<script type="text/javascript" src="/java/jquery.min.js"></script>
		<script type="text/javascript" src="/java/highcharts.js"></script>
		<script type="text/javascript" src="/java/modules/exporting.js"></script>
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript" src="/java/graph_bar_4.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
   			 <div id="container">
                <form action="" method="post" name="frm">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td colspan="2" height="35" style="color:#02880a; font-size:15px; font-weight: bold;">&nbsp;<%=sub_tit%></td>
                    <td width="2%"></td>
                    <td width="25%" height="35" style="color:#02880a; font-size:15px; font-weight: bold;">&nbsp;공지사항 / 자료실</td>
                    <td width="24%" align="right"><a href="board_list.asp?board_gubun=1"><img src="image/list_view.gif" alt="" width="44" height="15"></a>&nbsp;&nbsp;&nbsp;</td>
                  </tr>
                  <tr>
                    <td colspan="2" rowspan="3" valign="top">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="tableList">
						<colgroup>
							<col width="*%" >
							<col width="15%" >
							<col width="15%" >
							<col width="18%" >
							<col width="18%" >
							<col width="18%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">관리그룹</th>
								<th scope="col">평가점수</th>
								<th scope="col">평가등급</th>
								<th scope="col">보상금</th>
								<th scope="col">위약금</th>
								<th scope="col">차액</th>
							</tr>
						</thead>
					<% 
					for i = 0 to 3	
						aver_score = (group_tab(0,1)+group_tab(1,1)+group_tab(2,1)+group_tab(3,1))/4
						aver_view = "보통"
						if aver_score >= 95.00 then
							aver_view = "우수"
						end if
						if aver_score < 80.00 then
							aver_view = "미흡"
						end if						
						if aver_score = 0 then
							aver_view = "."
						end if
					%>
                      <tr>
                        <td><%=group_tab(i,0)%></td>
                        <td bgcolor="#EEFFFF">
						<a href="#" onClick="pop_Window('sla_month_group_view.asp?sla_month=<%=sla_month%>&sla_group=<%=i%>','month_group_popup','width=950,height=500')"><%=formatnumber(group_tab(i,1),2)%></a>
                        </td>
                        <td><%=group_tab(i,2)%></td>
                        <td><%=formatnumber(group_tab(i,3),0)%></td>
                        <td><%=formatnumber(group_tab(i,4),0)%></td>
                        <td><%=formatnumber(group_tab(i,3)+group_tab(i,4),0)%></td>
                      </tr>
					<% next	%>
                      <tr>
                        <td><%=group_tab(4,0)%></td>
                        <td bgcolor="#EEFFFF"><%=formatnumber(aver_score,2)%></td>
                        <td><%=aver_view%></td>
                        <td><%=formatnumber(group_tab(4,3),0)%></td>
                        <td><%=formatnumber(group_tab(4,4),0)%></td>
                        <td><%=formatnumber(group_tab(4,3)+group_tab(4,4),0)%></td>
                      </tr>
                    </table></td>
                    <td width="2%"></td>
                    <td colspan="2" rowspan="3" valign="top">
                    <table cellpadding="0" cellspacing="0" class="tableList">
                      <colgroup>
                        <col width="7%" >
                        <col width="10%" >
                        <col width="*" >
                        <col width="12%" >
                      </colgroup>
                      <thead>
                        <tr>
                          <th class="first" scope="col">순번</th>
                          <th scope="col">구분</th>
                          <th scope="col">제목</th>
                          <th scope="col">작성일</th>
                        </tr>
                      </thead>
                      <tbody>
                        <%
						i = 0
						do until rs1.eof
							i = i + 1
							if rs1("board_gubun") = "1" then
								board_gubun = "공지사항"
							  else
							  	board_gubun = "자료실"
							end if							
						  	board_title = rs1("board_title")
							if len(rs1("board_title")) > 30 then
								board_title = mid(rs1("board_title"),1,30) + " ..."
							end if
						%>
                        <tr>
                          <td class="first"><%=i%></td>
                          <td><%=board_gubun%></td>
                          <td class="left"><a href="board_view.asp?board_gubun=<%=rs1("board_gubun")%>&board_seq=<%=rs1("board_seq")%>&be_pg=<%="C"%>"><%=board_title%></a>&nbsp;
                            <%	if rs1("reg_date") > new_date then 	%>
                            <img src="image/new.gif" width="24" height="11" border="0">
                          <%	end if	%></td>
                          <td><%=mid(rs1("reg_date"),1,10)%></td>
                        </tr>
                        <%
							rs1.movenext()
						loop
						%>
                      </tbody>
                    </table>
					<% if i = 0 then	%>
                       <h3 class="teof"> 내역이 없습니다 !!! </h3>
                    <% end if %>
                    </td>
                  </tr>
                  <tr>
                    <td width="2%" height="30">&nbsp;</td>
                    </tr>
                  <tr>
                    <td width="2%">&nbsp;</td>
                    </tr>
                  <tr>
                    <td width="25%">&nbsp;</td>
                    <td width="24%">&nbsp;</td>
                    <td width="2%">&nbsp;</td>
                    <td width="25%">&nbsp;</td>
                    <td width="24%">&nbsp;</td>
                  </tr>
                </table>

		<table width="100%" border="0" cellpadding="0" cellspacing="0">
		  <tr>
		    <td width="25%">
				<div id="graph_view1" style="width: 305px; height: 200px; margin: 0 auto"></div>
            </td>
		    <td width="25%">
				<div id="graph_view2" style="width: 305px; height: 200px; margin: 0 auto"></div>
            </td>
		    <td width="25%">
				<div id="graph_view3" style="width: 305px; height: 200px; margin: 0 auto"></div>
            </td>
		    <td width="25%">
				<div id="graph_view4" style="width: 305px; height: 200px; margin: 0 auto"></div>
            </td>
	      </tr>
		  </table>

		<input name="sla_month" type="hidden" value="<%=sla_month%>">
	<% 
		for i = 1 to 12
			for j = 0 to 3
	%>
		<input name="s_tab<%=i%><%=j%>" type="hidden" value="<%=s_tab(i,j)%>">
	<%
			next
		next
	%>
	</div>
</div>	     				
	</form>
	</body>
</html>

