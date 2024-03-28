<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sido_tab
dim acpt_tab(17,2)
dim inst_tab(17,2)
dim ran_tab(17,2)
sido_tab = array("계","서울","경기","부산","대구","인천","광주","대전","울산","강원","경남","경북","세종","충남","충북","전남","전북","제주")

large_paper_no=Request("large_paper_no")
company=Request("company")
as_type=Request("as_type")
acpt_cnt=Request("acpt_cnt")

for i = 0 to 17
	for j = 1 to 2
		acpt_tab(i,j) = 0
		inst_tab(i,j) = 0
		ran_tab(i,j) = 0
	next
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

' 완료건
sql = "select sido,as_process,count(*) as pro_cnt,sum(dev_inst_cnt) as inst_cnt,sum(ran_cnt) as ran_cnt from as_acpt WHERE large_paper_no ='"&large_paper_no&"' GROUP BY sido, as_process"
Rs.Open Sql, Dbconn, 1

do until rs.eof
	select case rs("sido")
		case "서울"
			i = 1
		case "경기"
			i = 2
		case "부산"
			i = 3
		case "대구"
			i = 4
		case "인천"
			i = 5
		case "광주"
			i = 6
		case "대전"
			i = 7
		case "울산"
			i = 8
		case "강원"
			i = 9
		case "경남"
			i = 10
		case "경북"
			i = 11
		case "세종"
			i = 12
		case "충남"
			i = 13
		case "충북"
			i = 14
		case "전남"
			i = 15
		case "전북"
			i = 16
		case "제주"
			i = 17
	end select	

	if rs("as_process") = "완료" or rs("as_process") = "취소" then
		acpt_tab(i,1) = acpt_tab(i,1) + cint(rs("pro_cnt"))
		inst_tab(i,1) = inst_tab(i,1) + cint(rs("inst_cnt"))
		ran_tab(i,1) = ran_tab(i,1) + cint(rs("ran_cnt"))
		acpt_tab(0,1) = acpt_tab(0,1) + cint(rs("pro_cnt"))
		inst_tab(0,1) = inst_tab(0,1) + cint(rs("inst_cnt"))
		ran_tab(0,1) = ran_tab(0,1) + cint(rs("ran_cnt"))
	  else
		acpt_tab(i,2) = acpt_tab(i,2) + cint(rs("pro_cnt"))
		inst_tab(i,2) = inst_tab(i,2) + cint(rs("inst_cnt"))
		ran_tab(i,2) = ran_tab(i,2) + cint(rs("ran_cnt"))
		acpt_tab(0,2) = acpt_tab(0,2) + cint(rs("pro_cnt"))
		inst_tab(0,2) = inst_tab(0,2) + cint(rs("inst_cnt"))
		ran_tab(0,2) = ran_tab(0,2) + cint(rs("ran_cnt"))
	end if
	rs.movenext()
loop
rs.close()

title_line = "지역별 대량건 처리 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="*" >
							<col width="35%" >
							<col width="15%" >
							<col width="35%" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first" scope="col">회사</th>
								<td><%=company%></td>
								<th>문서번호</th>
								<td><%=large_paper_no%></td>
							</tr>
							<tr>
								<th class="first" scope="col">처리유형</th>
								<td><%=as_type%></td>
								<th>총건수</th>
								<td><%=formatnumber(acpt_cnt,0)%></td>
							</tr>
						</tbody>
					</table>
					<h3 class="stit">* 시도별 내역</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">시도</th>
								<th scope="col" colspan="3" style=" border-bottom:1px solid #e3e3e3;">접수건수</th>
								<th scope="col" colspan="2" style=" border-bottom:1px solid #e3e3e3;">설치수량</th>
								<th scope="col" colspan="2" style=" border-bottom:1px solid #e3e3e3;">랜공사수량</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">완료건수</th>
								<th scope="col">미처리건수</th>
								<th scope="col">진척율</th>
								<th scope="col">완료수량</th>
								<th scope="col">미처리수량</th>
								<th scope="col">완료수량</th>
								<th scope="col">미처리수량</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                              <th>계</th>
                              <th class="right"><%=formatnumber(acpt_tab(0,1),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(acpt_tab(0,2),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(acpt_tab(0,1)/(acpt_tab(0,1)+acpt_tab(0,2))*100,2)%>%&nbsp;</th>
                              <th class="right"><%=formatnumber(inst_tab(0,1),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(inst_tab(0,2),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(ran_tab(0,1),0)%>&nbsp;</th>
                              <th class="right"><%=formatnumber(ran_tab(0,2),0)%>&nbsp;</th>
							</tr>
						<% 	
                    	for i = 1 to 17
                		%>
							<tr>
                              <td><%=sido_tab(i)%></td>
                              <td class="right"><%=formatnumber(acpt_tab(i,1),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(acpt_tab(i,2),0)%>&nbsp;</td>
                              <td class="right">
							<% if acpt_tab(i,1) = 0 and acpt_tab(i,2) = 0 then	%>
                               0.00%
                            <%   else	%>							
							  <%=formatnumber(acpt_tab(i,1)/(acpt_tab(i,1)+acpt_tab(i,2))*100,2)%>%&nbsp;
							<% end if %>
                              </td>
                              <td class="right"><%=formatnumber(inst_tab(i,1),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(inst_tab(i,2),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(ran_tab(i,1),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(ran_tab(i,2),0)%>&nbsp;</td>
							</tr>
                		<% 	
						next 
						%>
						</tbody>
					</table>
				</div>
		</div>				
	</body>
</html>

