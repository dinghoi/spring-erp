' 대량 데이터 batch upload
'On Error resume next

Dim DbConnect
DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=kwon_admin(*)14;"

Set Dbconn=CreateObject("ADODB.Connection")
Set Rs = CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

Dbconn.BeginTrans

curr_date = now()
curr_dd = cstr(mid(curr_date,9,2))
end_sw = "N"

	be_date = dateadd("m",-1,curr_date)
	be_month = cstr(mid(be_date,1,4)) + cstr(mid(be_date,6,2))

	from_date = mid(be_month,1,4) + "-" + mid(be_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))

	sql = "insert into emp_org_mst_month select '"&be_month&"' as org_month,emp_org_mst.* from emp_org_mst"
	dbconn.execute(sql)
	end_sw = "Y"

if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
	if end_sw = "Y" then
		msgbox be_month + "월 조직마감되었습니다. 시간 : " + cstr(now())
	  else
		msgbox "마감 처리건수가 없습니다. 시간 : " + cstr(now())
	end if
end if

set rs = nothing

dbconn.Close()
Set dbconn = Nothing
