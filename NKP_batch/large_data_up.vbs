' 대량 데이터 batch upload
'On Error resume next

Dim DbConnect
DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=kwon_admin(*)14;"
Set Dbconn=CreateObject("ADODB.Connection")
Set Rs = CreateObject("ADODB.Recordset")
Set Rs_etc = CreateObject("ADODB.Recordset")
Set rs_trade = CreateObject("ADODB.Recordset")
Set rs_hol = CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

Dbconn.BeginTrans

sql = "select * from large_acpt where upload_ok = 'N'" 
Rs.Open Sql, Dbconn, 1 

i = 0
do until rs.eof
	i = i + 1
	sql = "insert into as_acpt(acpt_date,acpt_man,acpt_grade,acpt_user,user_grade,tel_ddd,tel_no1,tel_no2,hp_ddd,hp_no1,hp_no2,company,dept"& _					
	",sido,gugun,dong,addr,mg_ce_id,mg_ce,mg_group,as_memo,request_date,request_time,as_process,as_type,maker,as_device,model_no,serial_no"& _
	",asets_no,reside,reside_place,reside_company,team,large_paper_no,sms,dev_inst_cnt,ran_cnt,work_man_cnt,alba_cnt,start_date,end_date"& _
	",reg_id) values (now(),'"&rs("acpt_man")&"','"&rs("acpt_grade")&"','"&rs("acpt_user")&"','"&rs("user_grade")&"','"&rs("tel_ddd")& _
	"','"&rs("tel_no1")&"','"&rs("tel_no2")&"','"&rs("hp_ddd")&"','"&rs("hp_no1")&"','"&rs("hp_no2")&"','"&rs("company")&"','"&rs("dept")& _
	"','"&rs("sido")&"','"&rs("gugun")&"','"&rs("dong")&"','"&rs("addr")&"','"&rs("mg_ce_id")&"','"&rs("mg_ce")&"','"&rs("mg_group")& _
	"','"&rs("as_memo")&"','"&rs("request_date")&"','"&rs("request_time")&"','"&rs("as_process")&"','"&rs("as_type")&"','"&rs("maker")& _
	"','"&rs("as_device")&"','"&rs("model_no")&"','"&rs("serial_no")&"','"&rs("asets_no")&"','"&rs("reside")&"','"&rs("reside_place")& _
	"','"&rs("reside_company")&"','"&rs("team")&"','"&rs("paper_no")&"','"&rs("sms")&"',"&rs("dev_inst_cnt")&","&rs("ran_cnt")& _
	","&rs("work_man_cnt")&","&rs("alba_cnt")&",'"&rs("request_date")&"','"&rs("end_date")&"','"&rs("reg_id")&"')"	
	dbconn.execute(sql)
	sql = "update large_acpt set upload_ok='Y' where acpt_no="&int(rs("acpt_no"))
	dbconn.execute(sql)	  

	rs.movenext
loop

if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
end if

msgbox cstR(i) + "건 완료되었습니다 (대량건 업로드) 시간 : " + cstr(now())

set rs = nothing

dbconn.Close()
Set dbconn = Nothing
