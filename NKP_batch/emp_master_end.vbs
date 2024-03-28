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

	sql = "insert into emp_master_month select '"&be_month&"' as emp_month,emp_master.* from emp_master"
	dbconn.execute(sql)
	end_sw = "Y"

	'sql = "select * from emp_master_month where emp_month = '"&be_month&"' order by emp_no" 
	'Rs.Open Sql, Dbconn, 1 
	'do until rs.eof
		' 일반비용 
		'sql = "update general_cost set emp_company='"&rs("emp_company")&"',bonbu='"&rs("emp_bonbu")&"',saupbu='"&rs("emp_saupbu")&"',team='"&rs("emp_team")&"',org_name='"&rs("emp_org_name")&"',reside_place='"&rs("emp_reside_place")&"' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (slip_gubun = '비용') and (tax_bill_yn = 'N' or isnull(tax_bill_yn)) and (emp_no='"&rs("emp_no")&"')"
		'dbconn.execute(sql)	  

		' 교통비
		'sql = "update transit_cost set emp_company='"&rs("emp_company")&"',bonbu='"&rs("emp_bonbu")&"',saupbu='"&rs("emp_saupbu")&"',team='"&rs("emp_team")&"',org_name='"&rs("emp_org_name")&"',reside_place='"&rs("emp_reside_place")&"' where (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (mg_ce_id='"&rs("emp_no")&"')"
		'dbconn.execute(sql)	  

		' 야특근
		'sql = "update overtime set emp_company='"&rs("emp_company")&"',bonbu='"&rs("emp_bonbu")&"',saupbu='"&rs("emp_saupbu")&"',team='"&rs("emp_team")&"',org_name='"&rs("emp_org_name")&"',reside_place='"&rs("emp_reside_place")&"' where (work_date >='"&from_date&"' and work_date <='"&to_date&"') and (mg_ce_id='"&rs("emp_no")&"')"
		'dbconn.execute(sql)	  

		' 카드전표
		'sql = "update card_slip set emp_company='"&rs("emp_company")&"',bonbu='"&rs("emp_bonbu")&"',saupbu='"&rs("emp_saupbu")&"',team='"&rs("emp_team")&"',org_name='"&rs("emp_org_name")&"',reside_place='"&rs("emp_reside_place")&"' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_no='"&rs("emp_no")&"')"
		'dbconn.execute(sql)	  
		
		'rs.movenext()

	'loop

if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
	if end_sw = "Y" then
		msgbox be_month + "월 마감되었습니다. 시간 : " + cstr(now())
	  else
		msgbox "마감 처리건수가 없습니다. 시간 : " + cstr(now())
	end if
end if

set rs = nothing

dbconn.Close()
Set dbconn = Nothing
