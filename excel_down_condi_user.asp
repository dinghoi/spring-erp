<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

	dim title_name
	dim company_tab(50)

	title_name = array("접수번호","접수일자","접수자","사용자","전화번호","핸드폰","회사","조직명","주소","CE명","장애내역","요청일","요청시간","처리일","처리시간","진행","처리방법","입고/지연사유","입고일자","대체여부","메이커","장애장비","자산코드","모델명","S/N번호","처리내용","설치수량","PC S/W","PC H/W","모니터","프린터/스케너","통신장비","서버/워크","아답터","기타")
	from_date = request("from_date")
	to_date = request("to_date")
	company = request("company")
	date_sw = request("date_sw")
	process_sw = request("process_sw")
	field_check = request("field_check")
	field_view = request("field_view")
	savefilename = from_date + to_date + ".xls"

 	Response.Buffer = True
  	Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
  	Response.CacheControl = "public"
  	Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

	if reside = "9" then
		k = 0
		Sql="select * from trade where use_sw = 'Y' and group_name = '"+user_name+"' order by trade_name asc"
		rs_trade.Open Sql, Dbconn, 1
		do until rs_trade.eof
			k = k + 1
			company_tab(k) = rs_trade("trade_name")
			rs_trade.movenext()
		loop
		rs_trade.close()						
	end if
	
	if reside = "9" and company = "전체" then
		com_sql = "company = '" + company_tab(1) + "'"	
		for kk = 2 to k
			com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
		next
		condi_sql = " or " + com_sql + ") "
	  else
		condi_sql = " or company = '" + reside_company + "' or company = '" + company + "') "
	end if

	'//2017-06-07 아이티퓨처(사번:900002) 로그인시 웅진관련 기업 검색하게 수정
	If  user_id = "900002" Then
		If Trim(company&"")="" Then 
		condi_sql = " or company in ('웅진식품','웅진씽크빅','코웨이') " & condi_sql
		End IF
	End IF

	base_sql = "select acpt_no,acpt_date,acpt_man,acpt_user,concat(tel_ddd,'-',tel_no1,'-',tel_no2),concat(hp_ddd,'-',hp_no1,'-',hp_no2),company,dept,concat(sido,' ',gugun,' ',dong,' ',addr),mg_ce,as_memo,request_date,request_time,visit_date,visit_time,"
	base_sql = base_sql + "as_process,as_type,visit_request_yn,into_reason,in_date,in_replace,maker,as_device,asets_no,model_no,serial_no,as_history,dev_inst_cnt,err_pc_sw,err_pc_hw,err_monitor,err_printer,err_network,err_server,err_adapter,err_etc from as_acpt "
	
	if date_sw = "acpt" then
		date_sql = "where (CAST(acpt_date as date) >= '" + from_date  + "' and CAST(acpt_date as date) <= '" + to_date  + "') and (acpt_man = '" + user_name + "'" + condi_sql
	  else
		date_sql = "where (visit_date >= '" + from_date  + "' and visit_date <= '" + to_date  + "') and (acpt_man = '" + user_name + "'" + condi_sql
	end if
	
	if process_sw = "Y" then
		process_sql = " and ( as_process = '완료' or as_process = '대체' or as_process = '취소' ) "
	  else
		process_sql = " and ( as_process = '접수' or as_process = '연기' or as_process = '입고' or as_process = '대체입고' ) "
	end if
	
	if field_check <> "total" then
		if field_check = "asets_no" then
			field_sql = " and ( " + field_check + " = '" + field_view + "' ) "
		  else			
			field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
		end if
	  else
		field_sql = " "
	end if
	order_sql = " ORDER BY acpt_date DESC"
	
	sql = base_sql + date_sql + process_sql + field_sql + order_sql
	Rs.Open Sql, Dbconn, 1
	if rs.eof then
		response.write"<script language=javascript>"
		response.write"alert('다운 할 자료가 없습니다 ....');"
		response.write"history.go(-1);"
		response.write"</script>"	
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
<body>
<table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
	<tr><%=chr(13)&chr(10)%>
<%	
	i = 0
	for each whatever in rs.fields	
		if i < 35 then
%>
			<td><b><%=title_name(i)%></b></TD><%=chr(13)&chr(10)%>
<%		
		end if
		i = i + 1
	next	%>
	</tr><%=chr(13)&chr(10)%>
<%
alldata=rs.getrows

numcols=ubound(alldata,1)
numrows=ubound(alldata,2)

jj = 0
FOR j= 0 TO numrows 
	jj = jj + 1
	if	jj > 1500 then
		jj = 0
		response.flush
	end if
%>
	<tr><%=chr(13)&chr(10)%>
<%  FOR i=0 to numcols
	if i > 34 then
		exit for
	end if
    thisfield=alldata(i,j)
'	if i = 8 then
'    	thisfield=Replace(alldata(i,j),chr(13)&chr(10),"<pre>")
'	end if	  
      if isnull(thisfield) then
         thisfield=""
      end if
      if trim(thisfield)="" then
         thisfield=""
      end if
	err_memo = ""
	if i > 26  and i < 35 then
		if thisfield <> "" then
			for k = 1 to 100 step 6
				chkfield = mid(thisfield,k,4)
				if chkfield = "" or chkfield= null then
					exit for		
				end if
				sql_etc = "select * from etc_code where etc_code = '" + chkfield +"'"
				Set Rs_etc=dbconn.execute(Sql_etc)
				if rs_etc.eof or rs_etc.bof then
					etc_name = ""
				  else
				  	etc_name = rs_etc("etc_name")
					if err_memo = "" then
						err_memo = etc_name
					  else
						err_memo = err_memo + "," +etc_name
					end if
				end if				
				rs_etc.close()
			next			
		end if		
		thisfield = err_memo
	end if
%>
<%	if i = 1 then %>
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
