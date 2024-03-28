<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type      = request.form("u_type")
	card_no     = request.form("card_no")
	old_emp_no  = request.form("old_emp_no")
	emp_no      = request.form("emp_no")
    emp_name    = request.form("emp_name")
    emp_grade   = request.form("emp_grade")     
    org_name    = request.form("org_name") 
	change_date = request.form("change_date")
	start_date  = request.form("start_date")
	mod_memo    = request.form("mod_memo")
	end_date    = dateadd("d",-1,datevalue(change_date))

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql="select * from memb where user_id = '"&old_emp_no&"'"
	set rs = dbconn.execute(sql)

	sql="select * from memb where user_id = '"&emp_no&"'"
	set rs_name = dbconn.execute(sql)

	sql="select max(history_seq) as max_seq from card_owner_history where card_no='" + card_no + "'"
	set rs_max = dbconn.execute(sql)
		
	if	isnull(rs_max("max_seq"))  then
		history_seq = "01"
	  else
		max_seq = "0" + cstr((int(rs_max("max_seq")) + 1))
		history_seq = right(max_seq,2)
	end if

    if not Rs.BOF then
        sql = "INSERT INTO card_owner_history ( card_no                    "&chr(13)&_
              "                               , history_seq                "&chr(13)&_
              "                               , emp_no                     "&chr(13)&_
              "                               , emp_name                   "&chr(13)&_
              "                               , emp_job                    "&chr(13)&_
              "                               , emp_company                "&chr(13)&_
              "                               , bonbu                      "&chr(13)&_
              "                               , saupbu                     "&chr(13)&_
              "                               , team                       "&chr(13)&_
              "                               , org_name                   "&chr(13)&_
              "                               , reside_place               "&chr(13)&_
              "                               , reside_company             "&chr(13)&_
              "                               , start_date                 "&chr(13)&_
              "                               , end_date                   "&chr(13)&_
              "                               , mod_memo                   "&chr(13)&_
              "                               , reg_id                     "&chr(13)&_
              "                               , reg_name                   "&chr(13)&_
              "                               , reg_date)                  "&chr(13)&_
              "                       VALUES  ( '"&card_no&"'              "&chr(13)&_
              "                               , '"&history_seq&"'          "&chr(13)&_
              "                               , '"&old_emp_no&"'           "&chr(13)&_
              "                               , '"&rs("user_name")&"'      "&chr(13)&_
              "                               , '"&rs("user_grade")&"'     "&chr(13)&_
              "                               , '"&rs("emp_company")&"'    "&chr(13)&_
              "                               , '"&rs("bonbu")&"'          "&chr(13)&_
              "                               , '"&rs("saupbu")&"'         "&chr(13)&_
              "                               , '"&rs("team")&"'           "&chr(13)&_
              "                               , '"&rs("org_name")&"'       "&chr(13)&_
              "                               , '"&rs("reside_place")&"'   "&chr(13)&_
              "                               , '"&rs("reside_company")&"' "&chr(13)&_
              "                               , '"&start_date&"'           "&chr(13)&_
              "                               , '"&end_date&"'             "&chr(13)&_
              "                               , '"&mod_memo&"'             "&chr(13)&_
              "                               , '"&user_id&"'              "&chr(13)&_
              "                               , '"&user_name&"'            "&chr(13)&_
              "                               , now()                      "&chr(13)&_
              "                               )                            "&chr(13)
    else    
        sql = "INSERT INTO card_owner_history ( card_no                    "&chr(13)&_
              "                               , history_seq                "&chr(13)&_
              "                               , emp_no                     "&chr(13)&_
              "                               , emp_name                   "&chr(13)&_
              "                               , emp_job                    "&chr(13)&_
              "                               , emp_company                "&chr(13)&_
              "                               , bonbu                      "&chr(13)&_
              "                               , saupbu                     "&chr(13)&_
              "                               , team                       "&chr(13)&_
              "                               , org_name                   "&chr(13)&_
              "                               , reside_place               "&chr(13)&_
              "                               , reside_company             "&chr(13)&_
              "                               , start_date                 "&chr(13)&_
              "                               , end_date                   "&chr(13)&_
              "                               , mod_memo                   "&chr(13)&_
              "                               , reg_id                     "&chr(13)&_
              "                               , reg_name                   "&chr(13)&_
              "                               , reg_date)                  "&chr(13)&_
              "                       VALUES  ( '"&card_no&"'              "&chr(13)&_
              "                               , '"&history_seq&"'          "&chr(13)&_
              "                               , '"&old_emp_no&"'           "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , ''                         "&chr(13)&_
              "                               , '"&start_date&"'           "&chr(13)&_
              "                               , '"&end_date&"'             "&chr(13)&_
              "                               , '"&mod_memo&"'             "&chr(13)&_
              "                               , '"&user_id&"'              "&chr(13)&_
              "                               , '"&user_name&"'            "&chr(13)&_
              "                               , now()                      "&chr(13)&_
              "                               )                            "&chr(13)        
    end if

    'Response.Write  "<pre>" & Sql &"</pre>"
    dbconn.execute(sql)

    sql = "UPDATE card_owner                             "&chr(13)&_
          "   SET emp_no     ='"&emp_no&"'               "&chr(13)&_
          "     , emp_name   ='"&rs_name("user_name")&"' "&chr(13)&_
          "     , start_date ='"&change_date&"'          "&chr(13)&_
          "     , use_yn     ='Y'                        "&chr(13)&_
          "     , mod_id     ='"&user_id&"'              "&chr(13)&_
          "     , mod_name   ='"&user_name&"'             "&chr(13)&_
          "     , mod_date   = now()                     "&chr(13)&_
          " WHERE card_no = '"&card_no&"'                "&chr(13)
    'Response.Write  "<pre>" & Sql &"</pre>"      
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "변경되었습니다...."
	end if

	Response.write"<script language=javascript>"
	Response.write"alert('"&end_msg&"');"
	Response.write"parent.opener.location.reload();"
	Response.write"window.close();"		
	Response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
