<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/mysql_schema_db.asp" -->
<%
' 최근수정 2010-06-28

	write_date = "2010-01-04 09:00:21"
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect
	sql = "select write_date as bbbb from itft2005.as_acpt where write_date = '"+write_date+"'"
'	set rs=dbconn.execute(sql)
'	if rs.eof or rs.bof then
'		response.write("데이터 못 찾음")
'	  else
'		response.write(rs("bbbb"))
'	end if
'	write_date = rs("bbbb")
'	rs.close()
	w_date = formatdatetime(write_date,2)
	w_time = formatdatetime(write_date,4)
	w_sec = right(write_date,3)
	ww_date = w_date + " " + w_time + w_sec
	
	sql = "select * from itft2005.as_acpt where date_format(write_date,'%Y-%m-%d %h:%i:%s') = '"+ww_date+"'"
'	set rs=dbconn.execute(sql)
'	if rs.eof or rs.bof then
'		response.write("데이터 못 찾음")
'	  else
'		response.write(rs("acpt_no"))
'	end if
%>
