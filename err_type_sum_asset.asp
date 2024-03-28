<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim err_type
dim err_code(6,25)
dim err_cnt(6,25)
dim err_name
dim company_tab(150)

for i = 0 to 6
	for j = 0 to 25
		err_cnt(i,j) = 0
		err_code(i,j) = ""
	next
next

err_type = array("S/W장애","H/W장애","모니터","프린터류","통신장비","서버/워크","아답터")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_com = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect
'ck_sw=Request("ck_sw")

c_name = user_name

if c_name = "전체" then
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	company = request.form("company")
  else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	company = c_name
end if  

If to_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	curr_dd = cstr(datepart("d",to_date))
	from_date = mid(to_date,1,8) + "01"
	company = user_name
End If
curr_dd = cstr(datepart("d",to_date))
if from_date > to_date then
	from_date = mid(to_date,1,8) + "01"
end if

if c_name = "전체" then
	k = 0
	company_tab(0) = "전체"
	if	c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
		Sql="select * from etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' and group_name = '"+user_name+"' order by etc_name asc"
		  else
		Sql="select * from etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' order by etc_name asc"
	end if
	Rs_etc.Open Sql, Dbconn, 1
	while not rs_etc.eof
		k = k + 1
		company_tab(k) = rs_etc("etc_name")
		rs_etc.movenext()
	Wend
rs_etc.close()						
end if				

'데스크탑 S/W 장애 (완료, 방문처리 및 원격처리)

grade_sql = ""
if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
	com_sql = "company = '" + company_tab(1) + "'"	
	for kk = 1 to k
		com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
	next
	grade_sql = " and (" + com_sql + ")"
end if
kkk = k

for k = 1 to 31 step 6
	if company = "전체" then
		sql = "select substring(err_pc_sw,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_pc_sw,"& k& ",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
			sql = sql + grade_sql
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_pc_sw,"& k& ",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY substring(err_pc_sw,"& k &",4)"
	  else
		sql = "select company,substring(err_pc_sw,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or c_grade = "8" or c_grade = "5" then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_pc_sw,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_pc_sw,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY company, substring(err_pc_sw,"& k &",4)"
	end if
	end_cnt = 0	
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		end_cnt = end_cnt + 1
		i = int(rs("err_code")) - 101
		if i > 25 then
			i = 25
		end if
		err_code(0,i) = rs("err_code")
		err_cnt(0,i) = err_cnt(0,i) + cint(rs("err_cnt"))
		rs.movenext()
	loop
	rs.close()
	if end_cnt = 0 then
		k = 99
	end if

next

'H/W 장애 (완료, A/S 방문 및 원격처리)
for k = 1 to 19 step 6
	if company = "전체" then
		sql = "select substring(err_pc_hw,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_pc_hw,"& k& ",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
			sql = sql + grade_sql
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_pc_hw,"& k& ",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY substring(err_pc_hw,"& k &",4)"
	  else
		sql = "select company,substring(err_pc_hw,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or c_grade = "8" or c_grade = "5" then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_pc_hw,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_pc_hw,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY company, substring(err_pc_hw,"& k &",4)"
	end if
	
	end_cnt = 0	
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		end_cnt = end_cnt + 1
		i = int(rs("err_code")) - 201
		if i > 25 then
			i = 25
		end if
		err_code(1,i) = rs("err_code")
		err_cnt(1,i) = err_cnt(1,i) + cint(rs("err_cnt"))
		rs.movenext()
	loop
	rs.close()
	if end_cnt = 0 then
		k = 99
	end if
next

'모니터 장애 (완료, A/S 방문 및 원격처리)
for k = 1 to 13 step 6
	if company = "전체" then
		sql = "select substring(err_monitor,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_monitor,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
			sql = sql + grade_sql
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_monitor,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY substring(err_monitor,"& k &",4)"
	  else
		sql = "select company,substring(err_monitor,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or c_grade = "8" or c_grade = "5" then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_monitor,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_monitor,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY company, substring(err_monitor,"& k &",4)"
	end if
	
	end_cnt = 0	
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		end_cnt = end_cnt + 1
		i = int(rs("err_code")) - 301
		if i > 25 then
			i = 25
		end if
		err_code(2,i) = rs("err_code")
		err_cnt(2,i) = err_cnt(2,i) + cint(rs("err_cnt"))
		rs.movenext()
	loop
	rs.close()
	if end_cnt = 0 then
		k = 99
	end if
next

'프린터 장애 (완료, A/S 방문 및 원격처리)
for k = 1 to 13 step 6
	if company = "전체" then
		sql = "select substring(err_printer,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_printer,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
			sql = sql + grade_sql
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_printer,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY substring(err_printer,"& k &",4)"
	  else
		sql = "select company,substring(err_printer,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or c_grade = "8" or c_grade = "5" then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_printer,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_printer,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY company, substring(err_printer,"& k &",4)"
	end if
	
	end_cnt = 0	
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		end_cnt = end_cnt + 1
		i = int(rs("err_code")) - 401
		if i > 25 then
			i = 25
		end if
		err_code(3,i) = rs("err_code")
		err_cnt(3,i) = err_cnt(3,i) + cint(rs("err_cnt"))
		rs.movenext()
	loop
	rs.close()
	if end_cnt = 0 then
		k = 99
	end if
next

'통신 장애 (완료, A/S 방문 및 원격처리)
for k = 1 to 13 step 6
	if company = "전체" then
		sql = "select substring(err_network,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_network,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
			sql = sql + grade_sql
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_network,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY substring(err_network,"& k &",4)"
	  else
		sql = "select company,substring(err_network,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or c_grade = "8" or c_grade = "5" then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_network,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_network,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY company, substring(err_network,"& k &",4)"
	end if
	
	end_cnt = 0	
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		end_cnt = end_cnt + 1
		i = int(rs("err_code")) - 501
		if i > 25 then
			i = 25
		end if
		err_code(4,i) = rs("err_code")
		err_cnt(4,i) = err_cnt(4,i) + cint(rs("err_cnt"))
		rs.movenext()
	loop
	rs.close()
	if end_cnt = 0 then
		k = 99
	end if
next

'서버,워크스테이션 (완료, A/S 방문 및 원격처리)
for k = 1 to 13 step 6
	if company = "전체" then
		sql = "select substring(err_server,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_server,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
			sql = sql + grade_sql
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_server,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY substring(err_server,"& k &",4)"
	  else
		sql = "select company,substring(err_server,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or c_grade = "8" or c_grade = "5" then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_server,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_server,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY company, substring(err_server,"& k &",4)"
	end if
	end_cnt = 0	
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		end_cnt = end_cnt + 1
		i = int(rs("err_code")) - 601
		if i > 25 then
			i = 25
		end if
		err_code(5,i) = rs("err_code")
		err_cnt(5,i) = err_cnt(5,i) + cint(rs("err_cnt"))
		rs.movenext()
	loop
	rs.close()
	if end_cnt = 0 then
		k = 99
	end if
next

'아답터 (완료, A/S 방문 및 원격처리)
for k = 1 to 13 step 6
	if company = "전체" then
		sql = "select substring(err_adapter,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or ( c_grade = "5" and c_reside = "1" ) then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_adapter,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
			sql = sql + grade_sql
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_adapter,"& k &",4)<>'') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY substring(err_adapter,"& k &",4)"
	  else
		sql = "select company,substring(err_adapter,"& k &",4) AS err_code, COUNT(*) AS err_cnt from as_acpt" 
		if c_grade = "7" or c_grade = "8" or c_grade = "5" then
			sql = sql + " WHERE (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_adapter,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		  else
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (as_process = '완료') and (as_type = '방문처리' or as_type = '원격처리') and (substring(err_adapter,"& k &",4)<>'') and (company = '"+company+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		end if
		sql = sql + " GROUP BY company, substring(err_adapter,"& k &",4)"
	end if
	end_cnt = 0	
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		end_cnt = end_cnt + 1
		i = int(rs("err_code")) - 701
		if i > 25 then
			i = 25
		end if
		err_code(6,i) = rs("err_code")
		err_cnt(6,i) = err_cnt(6,i) + cint(rs("err_cnt"))
		rs.movenext()
	loop
	rs.close()
	if end_cnt = 0 then
		k = 99
	end if
next

err_tot = 0
for k = 0 to 6
	for kk = 0 to 25
		if err_code(k,kk) <> "" then
			err_tot = err_tot + err_cnt(k,kk)
		end if
	next
next

title_line = "장애 유형별 현황"
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
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
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
				if (formcheck(document.frm) && chkfrm()) {
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
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/asset_header.asp" -->
			<!--#include virtual = "/include/asset_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=err_type_sum_asset.asp" method="post" name="frm">
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
								<strong>회사</strong>
                                <input name="company" type="hidden" id="company" value="<%=company%>">
								<%=company%>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="20%" >
							<col width="*" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">장애장비</th>
								<th scope="col">소계</th>
								<th scope="col">백분율(%)</th>
								<th scope="col">장애유형</th>
								<th scope="col">그래프</th>
								<th scope="col">건수</th>
								<th scope="col">백분율</th>
							</tr>
						</thead>
						<tbody>
              			<% for k = 0 to 6 %>
							<tr>
                              <td><%=err_type(k)%></td>
                        <%
							err_sub = 0
	
							for kk = 0 to 25
								if err_code(k,kk) <> "" then
									err_sub = err_sub + err_cnt(k,kk)
								end if
							next				
							if err_tot = 0 then
								err_sub_per = 0
							  else
								err_sub_per = err_sub/err_tot * 100
							end if
						%>
                              <td><%=formatnumber(err_sub,0)%></td>
                              <td><%=formatnumber(err_sub_per,2)%>%</td>
                              <td colspan="4">
								<table cellpadding="0" cellspacing="0" width="100%">
						<%
							if err_sub = 0 then
								err_cnt(k,0) = 0
								err_per = 0
						%>
									<tr>
		                      			<td width="240">처리내용 없음</td>
		                      			<td width="*">&nbsp;</td>
		                      			<td width="119">&nbsp;</td>
		                      			<td width="119">&nbsp;</td>
                                    </tr>
                        <%
		 					  else

								for kk = 0 to 25
								  if err_code(k,kk) <> "" then
					'					exit for
					'				end if
					
									if k < 7  or k = 12 then
										sql_etc = "select * from etc_code where etc_code = '" + err_code(k,kk) + "'"
										set rs_etc=dbconn.execute(sql_etc)
										if rs_etc.eof then
											err_name = "개발자문의"
										  else
											err_name = rs_etc("etc_name")
										end if
									  else
										err_name = err_code(k,kk)
									end if
									err_per = formatnumber((err_cnt(k,kk)/err_tot * 100),2)
						%>
									<tr>
		                      			<td width="240"><%=err_name%></td>
		                      			<td width="*" class="left"><img src="image/graph02.gif" width="<%=err_per%>%" height="13" align="center"></td>
		                      			<td width="119"><%=err_cnt(k,kk)%></td>
		                      			<td width="119"><%=err_per%>%</td>
                                    </tr>
                        <%
								  end if
								next
							end if
						%>
 								</table>
                              </td>
							</tr>
						<%  next	%>
                        </tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

