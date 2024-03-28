<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	dim title_name
	dim company_tab(50)
	view_c=Request("view_c")
	dong = request("dong")

	title_name = array("접수번호","접수일자","접수자","직급","사용자","전화번호","핸드폰","회사","조직명","주소","CE명","장애내역","요청일","요청시간","처리방법","진행","고객요청","제조사","장애장비","모델명","입고사유","입고상태")

	savefilename = user_id + ".xls"
 	Response.Buffer = True
  	Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
  	Response.CacheControl = "public"
  	Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_in = Server.CreateObject("ADODB.Recordset")
	Set rs_trade = Server.CreateObject("ADODB.Recordset")
	Dbconn.open DbConnect

	if c_grade = "7" then
		k = 0
	'	Sql="select * from trade where use_sw = 'Y' and mg_group = '"+mg_group+"' and group_name = '"+user_name+"' order by trade_name asc"
		Sql="select * from trade where use_sw = 'Y' and group_name = '"+user_name+"' order by trade_name asc"
		rs_trade.Open Sql, Dbconn, 1
		do until rs_trade.eof
			k = k + 1
			company_tab(k) = rs_trade("trade_name")
			rs_trade.movenext()
		loop
		rs_trade.close()
	end if

	if view_sort = "" then
		view_sort = "DESC"
	end if

	view_sql = " "
	if view_c = "as" then
		view_sql = " and (as_type = '방문처리' or as_type = '원격처리') "
	end if
	if view_c = "inst" then
		view_sql = " and (as_type <> '방문처리' and as_type <> '원격처리') "
	end if
	order_Sql = " ORDER BY acpt_date " + view_sort

	if view_c = "dong" then
		view_sql = " and (dong like '%" + dong + "%' )"
		order_Sql = " ORDER BY sido, gugun, dong " + view_sort
	end if

	if view_c = "large" then
		view_sql = " and (large_paper_no <> '') "
		order_Sql = " ORDER BY large_paper_no, sido, gugun, dong " + view_sort
	end if

	'where_sql = " WHERE (mg_group = '" + mg_group + "') and "
	base_sql = " where (as_process = '접수' or as_process = '입고' or as_process = '연기' or as_process = '대체입고') "
	condi_sql = " and (mg_ce_id = '" + user_id + "') "
	if c_grade = "0" or ( c_grade = "1" and team = "수도권지사" ) then
		condi_Sql = " "
	end if
	if ( c_grade = "1" and team <> "수도권지사" ) then
		condi_Sql = " and (team = '"+team+"' or mg_ce_id = '"+user_id+"') "
	end if
	if c_grade = "2" then
		'condi_Sql = " and (reside_place = '"+reside_place+"' or mg_ce_id = '"+user_id+"') "
		condi_Sql = " AND (company = '"+reside_company+"' OR mg_ce_id = '"+user_id+"') " ' 위의 사항에서 수정  (2018-11-16) 화면(ax_list_ce.asp) 와 다름으로 인햬..
	end if
	if c_grade = "3"  and team <> "수도권지사" then
		condi_Sql = " and (team = '"+team+"' or mg_ce_id = '"+user_id+"') "
	end if
	if c_grade = "3"  and team = "수도권지사" then
		condi_Sql = "and (mg_ce_id = '"+user_id+"') "
	end if

	if c_grade = "7" then
		com_sql = "company = '" + company_tab(1) + "'"
		for kk = 2 to k
			com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
		next
		where_sql = "WHERE "
		condi_Sql = " and (" + com_sql + ") "
	end if

	if c_grade = "8" then
		where_sql = "WHERE "
		condi_Sql = " and (company = '" + user_name + "') "
	end if

	sql = "select acpt_no,acpt_date,acpt_man,acpt_grade,acpt_user,concat(tel_ddd,'-',tel_no1,'-',tel_no2),concat(hp_ddd,'-',hp_no1,'-',hp_no2),company,dept,concat(sido,' ',gugun,' ',dong,' ',addr),mg_ce,as_memo,request_date,request_time,as_type,as_process,visit_request_yn,maker,as_device,model_no,into_reason from as_acpt "
	sql = sql + base_sql + view_sql + condi_sql + order_sql
	Rs.Open Sql, Dbconn, 1
	if rs.eof or rs.bof then
		response.write"<script language=javascript>"
		response.write"alert('다운 할 자료가 없습니다 ....');"
		response.write"history.go(-1);"
		response.write"</script>"
	end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title></title>
</head>
<body>
<table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
	<tr><%=chr(13)&chr(10)%>
<%
	i = 0
	for i = 0 to 21
'	for each whatever in rs.fields
'		if i < 21 then
%>
			<td><b><%=title_name(i)%></b></TD><%=chr(13)&chr(10)%>
<%
	next
'		end if
'		i = i + 1
'	next
%>
	</tr><%=chr(13)&chr(10)%>
<%
alldata=rs.getrows

numcols=ubound(alldata,1) + 1
numrows=ubound(alldata,2)

FOR j= 0 TO numrows
	in_process = ""
	if alldata(15,j) = "입고" then
		sql = "select into_date,in_process,in_place from as_into where acpt_no="&alldata(0,j)&" and in_seq="&"(select max(in_seq) from as_into where acpt_no="&alldata(0,j)&")"
		Set Rs_in=dbconn.execute(sql)
		if	Rs_in.eof then
				in_process = "없음"
			else
				in_process = rs_in("in_process")
		end if
	end if
	if alldata(16,j) = "Y" then
		alldata(16,j) = "방문요청"
	  else
		alldata(16,j) = ""
	end if

%>
	<tr><%=chr(13)&chr(10)%>
<%  FOR i=0 to numcols
	if i = 21 then
    	thisfield = in_process
	  else
		thisfield=alldata(i,j)
	end if
      if isnull(thisfield) then
         thisfield=""
      end if
      if trim(thisfield)="" then
         thisfield=""
      end if
%>
<%	if i = 1 or i = 11 then %>
		<td valign=top><%=thisfield%> </td><%=chr(13)&chr(10)%>
<%		else	%>
		<td style="mso-number-format:'\@'" valign=top><%=thisfield%> </td><%=chr(13)&chr(10)%>
<%	end if 		%>
<%  NEXT	%>
	</tr><%=chr(13)&chr(10)%>
<%NEXT%>
</table>

</body>
</html>
