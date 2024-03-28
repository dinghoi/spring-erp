<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	u_type = request.form("u_type")
	oil_unit_month = request.form("oil_unit_month")
	oil_unit_middle11 = int(request.form("oil_unit_middle11"))
	oil_unit_last11 = int(request.form("oil_unit_last11"))
	if oil_unit_last11 = 0 then
		oil_unit_average11 = oil_unit_middle11
	  else
		oil_unit_average11 = (oil_unit_middle11 + oil_unit_last11) / 2
	end if	  	
	oil_unit_middle12 = int(request.form("oil_unit_middle12"))
	oil_unit_last12 = int(request.form("oil_unit_last12"))
	if oil_unit_last12 = 0 then
		oil_unit_average12 = oil_unit_middle12
	  else
		oil_unit_average12 = (oil_unit_middle12 + oil_unit_last12) / 2
	end if	  	
	oil_unit_middle13 = int(request.form("oil_unit_middle13"))
	oil_unit_last13 = int(request.form("oil_unit_last13"))
	if oil_unit_last13 = 0 then
		oil_unit_average13 = oil_unit_middle13
	  else
		oil_unit_average13 = (oil_unit_middle13 + oil_unit_last13) / 2
	end if	  	
	oil_unit_middle21 = int(request.form("oil_unit_middle21"))
	oil_unit_last21 = int(request.form("oil_unit_last21"))
	if oil_unit_last21 = 0 then
		oil_unit_average21 = oil_unit_middle21
	  else
		oil_unit_average21 = (oil_unit_middle21 + oil_unit_last21) / 2
	end if	  	
	oil_unit_middle22 = int(request.form("oil_unit_middle22"))
	oil_unit_last22 = int(request.form("oil_unit_last22"))
	if oil_unit_last22 = 0 then
		oil_unit_average22 = oil_unit_middle22
	  else
		oil_unit_average22 = (oil_unit_middle22 + oil_unit_last22) / 2
	end if	  	
	oil_unit_middle23 = int(request.form("oil_unit_middle23"))
	oil_unit_last23 = int(request.form("oil_unit_last23"))
	if oil_unit_last23 = 0 then
		oil_unit_average23 = oil_unit_middle23
	  else
		oil_unit_average23 = (oil_unit_middle23 + oil_unit_last23) / 2
	end if	  	

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	if	u_type = "U" then
		sql = "delete from oil_unit where oil_unit_month ='"&oil_unit_month&"'"
		dbconn.execute(sql)
	end if

	sql = "insert into oil_unit (oil_unit_month,oil_unit_id,oil_kind,oil_unit_middle,oil_unit_last,oil_unit_average,reg_id,reg_user,reg_date)"& _
	" values ('"&oil_unit_month&"','1','휘발유',"&oil_unit_middle11&","&oil_unit_last11&","&oil_unit_average11&",'"&user_id&"','"&user_name& _
	"',now())"
	dbconn.execute(sql)
	sql = "insert into oil_unit (oil_unit_month,oil_unit_id,oil_kind,oil_unit_middle,oil_unit_last,oil_unit_average,reg_id,reg_user,reg_date)"& _
	" values ('"&oil_unit_month&"','1','디젤',"&oil_unit_middle12&","&oil_unit_last12&","&oil_unit_average12&",'"&user_id&"','"&user_name& _
	"',now())"
	dbconn.execute(sql)
	sql = "insert into oil_unit (oil_unit_month,oil_unit_id,oil_kind,oil_unit_middle,oil_unit_last,oil_unit_average,reg_id,reg_user,reg_date)"& _
	" values ('"&oil_unit_month&"','1','가스',"&oil_unit_middle13&","&oil_unit_last13&","&oil_unit_average13&",'"&user_id&"','"&user_name& _
	"',now())"
	dbconn.execute(sql)
	sql = "insert into oil_unit (oil_unit_month,oil_unit_id,oil_kind,oil_unit_middle,oil_unit_last,oil_unit_average,reg_id,reg_user,reg_date)"& _
	" values ('"&oil_unit_month&"','2','휘발유',"&oil_unit_middle21&","&oil_unit_last21&","&oil_unit_average21&",'"&user_id&"','"&user_name& _
	"',now())"
	dbconn.execute(sql)
	sql = "insert into oil_unit (oil_unit_month,oil_unit_id,oil_kind,oil_unit_middle,oil_unit_last,oil_unit_average,reg_id,reg_user,reg_date)"& _
	" values ('"&oil_unit_month&"','2','디젤',"&oil_unit_middle22&","&oil_unit_last22&","&oil_unit_average22&",'"&user_id&"','"&user_name& _
	"',now())"
	dbconn.execute(sql)
	sql = "insert into oil_unit (oil_unit_month,oil_unit_id,oil_kind,oil_unit_middle,oil_unit_last,oil_unit_average,reg_id,reg_user,reg_date)"& _
	" values ('"&oil_unit_month&"','2','가스',"&oil_unit_middle23&","&oil_unit_last23&","&oil_unit_average23&",'"&user_id&"','"&user_name& _
	"',now())"
	dbconn.execute(sql)
		
	response.write"<script language=javascript>"
	response.write"alert('입력 완료 되었습니다....');"		
	response.Redirect "oil_unit_mg.asp"
	response.write"</script>"	
	Response.End
	dbconn.Close()
	Set dbconn = Nothing
%>
