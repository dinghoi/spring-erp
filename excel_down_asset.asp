<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

	dim title_name
	dim company_tab(50)

	title_name = array("접수번호","접수일자","접수자","사용자","전화번호","핸드폰","회사","조직명","주소","CE명","장애내역","요청일","요청시간","처리일","처리시간","진행","처리방법","입고/지연사유","입고일자","대체여부","메이커","장애장비","자산코드","모델명","처리내용","설치수량","PC S/W","PC H/W","모니터","프린터/스케너","통신장비","서버/워크","아답터","기타")
	from_date = request("from_date")
	to_date = request("to_date")
'	company = request("company")
	date_sw = request("date_sw")
	process_sw = request("process_sw")
	field_check = request("field_check")
	field_view = request("field_view")
	savefilename = from_date + to_date + ".xls"

 	Response.Buffer = True
  	Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
  	Response.CacheControl = "public"
  	Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect

	base_sql = "select acpt_no,acpt_man,as_type,acpt_date,as_process,acpt_user,as_memo,company,dept,tel_ddd,tel_no1,tel_no2,sido,gugun,request_date,visit_date,mg_ce,asets_no from as_acpt "
	
	if date_sw = "acpt" then
		date_sql = "where (CAST(acpt_date as date) >= '" + from_date  + "' and CAST(acpt_date as date) <= '" + to_date  + "') and (mg_group ='" + mg_group + "') and company = '" + user_name + "'"
	  else
		date_sql = "where (visit_date >= '" + from_date  + "' and visit_date <= '" + to_date  + "') and (mg_group ='" + mg_group + "') and company = '" + user_name + "'"
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
	
	'sql = base_sql + date_sql + process_sql + field_sql + order_sql
	
	com_sql = " "
	
	sql = base_sql + date_sql + com_sql + process_sql + field_sql + order_sql
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
		if i < 34 then
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
	if i > 33 then
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
	if i > 25  and i < 34 then
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
