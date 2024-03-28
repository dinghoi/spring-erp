<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include file="xmlrpc.asp"-->
<!--#include file="class.EmmaSMS.asp"-->
<%
'	on Error resume next

	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

    emp_user = request.cookies("nkpmg_user")("coo_user_name")

	u_type = abc("u_type")
	
	ask_empno = abc("ask_empno")
    ask_seq = abc("ask_seq")
    ask_date = abc("ask_date")
    ask_emp_name = abc("ask_emp_name")
    ask_company = abc("ask_company")
    ask_org = abc("ask_org")
    ask_org_name = abc("ask_org_name")
	ask_id = abc("ask_id")
    ask_type = abc("ask_type")
    ask_sawo_place = abc("ask_sawo_place")
    ask_sawo_comm = abc("ask_sawo_comm")
	
	v_att_file= abc("v_att_file")
	
	Set filenm = abc("att_file")(1)
	
	path = Server.MapPath ("/emp_sawo")
	filename = filenm.safeFileName
	
	fileType = mid(filename,inStrRev(filename,".")+1)

	save_path = path & "\" & filename
	
	if filenm.length > 1024*1024*8  then 
    	response.write "<script language=javascript>"
      	response.write "alert('파일 용량 2M를 넘으면 안됩니다.');"
		response.write "history.go(-1);"
      	response.write "</script>"
      	response.end
	End If		

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs_mem = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

' 경조회회원마스터 update를 위한...
'    Sql="select * from emp_sawo_mem where sawo_empno = '"&ask_empno&"'"
'    Rs_mem.Open Sql, Dbconn, 1
 
'    sawo_give_cnt = 0
'	sawo_give_pay = 0
'    if not Rs_mem.eof then
'       sawo_give_cnt = Rs_mem("sawo_give_count")
'       sawo_give_pay = Rs_mem("sawo_give_pay")
'    end if
'       Rs_mem.Close()

'    sawo_give_cnt = sawo_give_cnt + 1
'    sawo_give_pay = sawo_give_pay + give_pay

	dbconn.BeginTrans


	if	u_type = "U" then
	    if filenm <> "" then 
	       filenm.save save_path	
		sql = "update emp_sawo_ask set ask_id='"&ask_id&"',ask_type='"&ask_type&"',ask_sawo_place='"&ask_sawo_place&"',ask_sawo_comm='"&ask_sawo_comm&"',ask_att_file ='"+filename+"',ask_mod_date=now(),ask_mod_user='"&emp_user&"' where ask_empno ='"&ask_empno&"' and ask_seq = '"&ask_seq&"' and ask_date = '"&ask_date&"'"
		
		else
		
		sql = "update emp_sawo_ask set ask_id='"&ask_id&"',ask_type='"&ask_type&"',ask_sawo_place='"&ask_sawo_place&"',ask_sawo_comm='"&ask_sawo_comm&"',ask_mod_date=now(),ask_mod_user='"&emp_user&"' where ask_empno ='"&ask_empno&"' and ask_seq = '"&ask_seq&"' and ask_date = '"&ask_date&"'"
		end if
		dbconn.execute(sql)	  
	 else
	    if filenm <> "" then 
           filenm.save save_path	 
		sql = "insert into emp_sawo_ask(ask_empno,ask_seq,ask_date,ask_emp_name,ask_company,ask_org,ask_org_name,ask_id,ask_type,ask_process,ask_sawo_place,ask_sawo_comm,ask_att_file,ask_reg_date,ask_reg_user) values "
		sql = sql +	" ('"&ask_empno&"','"&ask_seq&"','"&ask_date&"','"&ask_emp_name&"','"&ask_company&"','"&ask_org&"','"&ask_org_name&"','"&ask_id&"','"&ask_type&"','0','"&ask_sawo_place&"','"&ask_sawo_comm&"','"&filename&"',now(),'"&emp_user&"')"
		
		else
		
		sql = "insert into emp_sawo_ask(ask_empno,ask_seq,ask_date,ask_emp_name,ask_company,ask_org,ask_org_name,ask_id,ask_type,ask_process,ask_sawo_place,ask_sawo_comm,ask_reg_date,ask_reg_user) values "
		sql = sql +	" ('"&ask_empno&"','"&ask_seq&"','"&ask_date&"','"&ask_emp_name&"','"&ask_company&"','"&ask_org&"','"&ask_org_name&"','"&ask_id&"','"&ask_type&"','0','"&ask_sawo_place&"','"&ask_sawo_comm&"',now(),'"&emp_user&"')"
        end if
		dbconn.execute(sql)
		
		'response.write sql
		
	end if
	
'	    sql = "update emp_sawo_mem set sawo_give_count='"&sawo_give_cnt&"',sawo_give_pay='"&sawo_give_pay&"' where sawo_empno ='"&give_empno&"'"
		'response.write sql
'		dbconn.execute(sql)	  
	

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
