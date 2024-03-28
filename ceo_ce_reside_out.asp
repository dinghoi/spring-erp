<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim end_cnt(200,10,2)
dim ce_tab(200,3)

from_date=Request.form("from_date")
to_date=Request.form("to_date")
team = "전체"
company_sum = 0

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
End If

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

for i = 0 to 200
	for j = 0 to 10
		end_cnt(i,j,1) = 0
		end_cnt(i,j,2) = 0
	next
next

sql = "select ce_work.mg_ce_id,memb.team,memb.org_name,memb.reside,memb.reside_place,memb.user_name from ce_work inner join memb on ce_work.mg_ce_id=memb.user_id where (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') and (memb.reside = '1') GROUP BY ce_work.mg_ce_id,memb.team,memb.org_name,memb.reside,memb.reside_place,memb.user_name Order By memb.team, memb.user_name Asc"
Rs.Open Sql, Dbconn, 1

i = 0
do until rs.eof
	i = i + 1
	if rs("team") = "" or isnull(rs("team")) then
		org_view = rs("org_name") 
	  else
	  	org_view = rs("team")
	end if
	ce_tab(i,1) = org_view
	ce_tab(i,2) = rs("user_name")
	ce_tab(i,3) = rs("reside_place")
	
    sql = "select ce_work.company,ce_work.reside_company,ce_work.mg_ce_id,ce_work.as_type,ce_work.holiday_yn,count(*) as end_cnt from ce_work WHERE (ce_work.company <> ce_work.reside_company) and (ce_work.reside_company<>'') and (ce_work.as_type<>'원격처리') and (ce_work.work_id='2') and (ce_work.mg_ce_id='"+rs("mg_ce_id")+"') and (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') GROUP BY ce_work.company,ce_work.reside_company,ce_work.as_type,ce_work.holiday_yn"		
    rs_as.Open Sql, Dbconn, 1
	do until rs_as.eof
		sql_trade = "select support_company from trade where trade_id ='매출' and trade_name = '"&rs_as("company")&"'"
		Set rs_trade = Dbconn.Execute (sql_trade)
		if rs_trade.eof or rs_trade.bof then
			company1 = rs_as("company")
		  else
			if rs_trade("support_company") = "없음" then
				company1 = rs_as("company")
			  else												
				company1 = rs_trade("support_company")
			end if
		end if
		rs_trade.close()
		
		sql_trade = "select support_company from trade where trade_id ='매출' and trade_name = '"&rs_as("reside_company")&"'"
		Set rs_trade = Dbconn.Execute (sql_trade)
		if rs_trade.eof or rs_trade.bof then
			company2 = rs_as("reside_company")
		  else
			if rs_trade("support_company") = "없음" then
				company2 = rs_as("reside_company")
			  else												
				company2 = rs_trade("support_company")
			end if
		end if
		rs_trade.close()									
		
        select case rs_as("as_type")
        	case "방문처리"
            	j = 1
        	case "신규설치"
            	j = 2
        	case "신규설치공사"
            	j = 3
        	case "이전설치"
            	j = 4
        	case "이전설치공사"
            	j = 5
        	case "랜공사"
            	j = 6
        	case "이전랜공사"
            	j = 7
        	case "장비회수"
            	j = 8
        	case "예방점검"
            	j = 9
        	case "기타"
            	j = 10
        end select												

		if company1 <> company2 then
			end_cnt(i,j,1) = end_cnt(i,j,1) + cint(rs_as("end_cnt"))
			end_cnt(i,0,1) = end_cnt(i,0,1) + cint(rs_as("end_cnt"))
			end_cnt(0,j,1) = end_cnt(0,j,1) + cint(rs_as("end_cnt"))
			end_cnt(0,0,1) = end_cnt(0,0,1) + cint(rs_as("end_cnt"))
		end if
		if rs_as("holiday_yn") = "Y" then
			if company1 <> company2 then
				end_cnt(i,j,2) = end_cnt(i,j,2) + cint(rs_as("end_cnt"))
				end_cnt(i,0,2) = end_cnt(i,0,2) + cint(rs_as("end_cnt"))
				end_cnt(0,j,2) = end_cnt(0,j,2) + cint(rs_as("end_cnt"))
				end_cnt(0,0,2) = end_cnt(0,0,2) + cint(rs_as("end_cnt"))
			end if
		end if
		rs_as.movenext()
	loop
	rs_as.close()

	rs.movenext()
loop
title_line = "상주자 외각 처리 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>임원 정보 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
  	<script src="/java/jquery-1.9.1.js"></script>
  	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>
	<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
		</script>
	<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.from_date.value > document.frm.to_date.value) {
					alert ("시작일이 종료일보다 클수가 없습니다");
					return false;
				}	
				return true;
			}
		</script>

</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/ceo_header.asp" -->
			<!--#include virtual = "/include/ceo_as_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=ce_reside_out.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                              <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="6%" >
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">소속</th>
								<th scope="col" rowspan="2">CE명</th>
								<th scope="col" rowspan="2">상주처</th>
								<th colspan="11" style=" border-bottom:1px solid #e3e3e3;" scope="col">
                                유형별 처리 현황 ( 전체수량/휴일근무수량 )
                                </th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">소계</th>
								<th scope="col">방문</th>
								<th scope="col">신규설치</th>
								<th scope="col">신규설치<br>공사</th>
								<th scope="col">이전설치</th>
								<th scope="col">이전설치<br>공사</th>
								<th scope="col">랜공사</th>
								<th scope="col">이전랜<br>공사</th>
								<th scope="col">회수</th>
								<th scope="col">예방</th>
								<th scope="col">기타</th>
							</tr>
						</thead>
					  <tbody>
					<% 
					ce_cnt = 0
					for  i = 1 to 200
						if end_cnt(i,0,1) > 0 then
							ce_cnt = ce_cnt + 1
                   	%>
						<tr>
                              <td><%=ce_tab(i,1)%></td>
                              <td><%=ce_tab(i,2)%></td>
                        <td><%=ce_tab(i,3)%></td>
							<%
                            for j = 0 to 10                        
                            %>
                              <td class="right"><%=formatnumber(end_cnt(i,j,1),0)%>/<%=end_cnt(i,j,2)%></td>
							<%
                            next                     
                            %>
							</tr>
					<%
						end if
					next
					%>
						<tr>
                          <th>총계</th>
                              <th><%=ce_cnt%></th>
                              <th>&nbsp;</th>
							<%
                            for j = 0 to 10                        
                            %>
                          <th class="right"><%=formatnumber(end_cnt(0,j,1),0)%>/<%=end_cnt(0,j,2)%></th>
							<%
                            next                     
                            %>
                          </tr>
 						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

