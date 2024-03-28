<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/asmg_dbcon.asp" -->
<!--#include virtual="/include/asmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

	dim title_name
	title_name = array("접수일자","접수자","사용자","전화번호","핸드폰","회사","조직명","주소","CE명","장애내역","요청일","요청시간","처리일","처리시간","진행","처리방법","입고/지연사유","입고일자","대체여부","메이커","장애장비","자산코드","모델명","처리내용","설치수량","PC S/W","PC H/W","모니터","프린터/스케너","통신장비","서버/워크","아답터","기타")
	from_date = request("from_date")
	to_date = request("to_date")
	company = request("company")
	savefilename = company + "_" + from_date + "_" + to_date + ".xls"

 	Response.Buffer = True
  	Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
  	Response.CacheControl = "public"
  	Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect

	if	company = "전체" then
'		sql = "select * from as_acpt "
		sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),replace(hp_ddd,' ','')+'-'+replace(hp_no1,' ','')+'-'+replace(hp_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,mg_ce,as_memo,request_date,request_time,visit_date,visit_time,"
		sql = sql + "as_process,as_type,into_reason,in_date,in_replace,maker,as_device,asets_no,model_no,as_history,dev_inst_cnt,err_pc_sw,err_pc_hw,err_monitor,err_printer,err_network,err_server,err_adapter,err_etc from as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
	  else
'	  	sql = "select * from as_acpt "
		sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),replace(hp_ddd,' ','')+'-'+replace(hp_no1,' ','')+'-'+replace(hp_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,mg_ce,as_memo,request_date,request_time,visit_date,visit_time,"
		sql = sql + "as_process,as_type,into_reason,in_date,in_replace,maker,as_device,asets_no,model_no,as_history,dev_inst_cnt,err_pc_sw,err_pc_hw,err_monitor,err_printer,err_network,err_server,err_adapter,err_etc,dev_inst_cnt from as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and company ='"+company+"' and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
	end if

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
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title></title>
</head>
<body>
<table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
	<tr><td></td><%=chr(13)&chr(10)%></tr>
	<tr><%=chr(13)&chr(10)%>
<%	
	i = 0
	for each whatever in rs.fields	
		if i < 33 then
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
	if i > 32 then
		exit for
	end if
    thisfield=alldata(i,j)
      if isnull(thisfield) then
         thisfield=""
      end if
      if trim(thisfield)="" then
         thisfield=""
      end if
	err_memo = ""
	if i > 24  and i < 33 then
		if thisfield <> "" then
			for k = 1 to 100 step 6
				chkfield = mid(thisfield,k,4)
				if chkfield = "" or chkfield= null then
					exit for		
				end if
				sql_etc = "select * from etc_code where etc_code = '" + chkfield +"'"
				Set Rs_etc=dbconn.execute(Sql_etc)
				if err_memo = "" then
					err_memo = rs_etc("etc_name")
				  else
					err_memo = err_memo + "," +rs_etc("etc_name")
				end if
				rs_etc.close()
			next			
		end if		
		thisfield = err_memo
	end if
%>
<%	if i = 0 or i = 9 then %>
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
