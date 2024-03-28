' 대량 데이터 batch upload
'On Error resume next

Dim DbConnect
dim work_date
DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=Wlsgustn6!;"

Set Dbconn=CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
Set rs_in = CreateObject("ADODB.Recordset")

work_date = datevalue(mid(dateadd("d",-1,now()),1,10))
work_date = "2014-08-14"
Dbconn.open dbconnect

Dbconn.BeginTrans

sql = "select as_acpt.*,memb.*, as_acpt.reg_id as reg_id, memb.team as team, memb.reside_place as reside_place, memb.reside as reside,"& _
	  " memb.reside_company from as_acpt inner join memb on as_acpt.reg_id = memb.user_id where (CAST(acpt_date as date) >= '"&work_date& _
	  "' and CAST(acpt_date as date) <= '"&work_date&"')"

Rs.Open Sql, Dbconn, 1 

i = 0
do until rs.eof
	i = i + 1
	sql="insert into ce_work (acpt_no,mg_ce_id,work_id,work_date,as_type,company,emp_company,bonbu,saupbu,team,org_name,reside_place,reside"& _
	",reside_company,work_man_cnt"&",dev_inst_cnt,ran_cnt,alba_cnt,person_amt,reg_id,reg_date) values ('"&rs("acpt_no")&"','"&rs("reg_id")& _
	"','1','"&work_date&"','"&rs("as_type")&"','"&rs("company")&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")& _
	"','"&rs("org_name")&"','"&rs("reside_place")&"','"&rs("reside")&"','"&rs("reside_company")&"',1,0,0,0,1,'"&rs("reg_id")&"',now())"
	dbconn.execute(sql)
	rs.movenext()
loop
rs.close()

if err.number <> 0 then
	Dbconn.RollbackTrans 
	msgbox "error !!!"
else    
	Dbconn.CommitTrans 
	msg = work_date + " " + cstr(i) + "건 접수자 처리되었습니다 ) 시간 : " + cstr(now()) 
	msgbox msg
end if

set rs = nothing

dbconn.Close()
Set dbconn = Nothing
