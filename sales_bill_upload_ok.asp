<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next
	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	sales_month = abc("sales_month")
	objFile = abc("objFile")

	from_date = mid(sales_month,1,4) + "-" + mid(sales_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))

	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	Set DbConn = Server.CreateObject("ADODB.Connection")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

	dbconn.BeginTrans

	cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
	rs.Open "select * from [1:10000]",cn,"0"

	rowcount=-1
	xgr = rs.getrows
	rowcount = ubound(xgr,2)
	fldcount = rs.fields.count

	tot_cnt = rowcount + 1
	read_cnt = 0
	write_cnt = 0

	if rowcount > -1 then
		for i=0 to rowcount
			if xgr(1,i) = "" or isnull(xgr(1,i)) then
				exit for
			end if
			if xgr(0,i) => from_date and xgr(0,i) <= to_date then
				sales_date = xgr(0,i)
				approve_no = xgr(1,i)
				sales_company = xgr(3,i)
				trade_no = xgr(4,i)
				trade_owner = xgr(6,i)
				price = cdbl(xgr(7,i))
				cost = cdbl(xgr(8,i))
				cost_vat = cdbl(xgr(9,i))
				imsi_sales_memo = xgr(10,i)
				sales_memo = replace(imsi_sales_memo,"'","")
				emp_name = xgr(11,i)
				saupbu = xgr(12,i)
				if isnull(xgr(13,i)) then
					slip_no = ""
				  else
					slip_no = replace(xgr(13,i),","," ")
				end if
				collect_due_date = xgr(14,i)
				if collect_due_date = "" or isnull(collect_due_date) then
					collect_due_date = ""
				  else
				  	collect_due_date = "20" + replace(collect_due_date,".","-")
				end if

				trade_no = Replace(trade_no,"-","")
				sql = "select * from trade where trade_no = '"&trade_no&"'"
				set rs_trade=dbconn.execute(sql)
				if rs_trade.eof or rs_trade.bof then
					trade_name = xgr(5,i)
					group_name = ""
				  else
					trade_name = rs_trade("trade_name")
					group_name = rs_trade("group_name")
				end if
				rs_trade.close()

				'sql = "select * from emp_master where emp_name = '"&emp_name&"'"		
				If saupbu = "기타사업부" Then 
					SQL = "SELECT emp_no FROM emp_master WHERE emp_name = '"&emp_name&"' "
				Else 
					SQL = "SELECT emp_no FROM emp_master AS emmt "
					SQL = SQL & "INNER JOIN emp_org_mst AS eomt ON eomt.org_code = emmt.emp_org_code "
					SQL = SQL & "WHERE emmt.emp_name = '"&emp_name&"' AND eomt.org_bonbu = '"&saupbu&"' "
				End If
				
				Set rs_emp = dbconn.execute(sql)

				If rs_emp.EOF Or rs_emp.BOF Then
					emp_no = "error"
				Else
					emp_no = rs_emp("emp_no")
				End If 
				rs_emp.close()

				read_cnt = read_cnt + 1

				sql = "select * from saupbu_sales where approve_no = '"&approve_no&"'"
				set rs=dbconn.execute(sql)
				if rs.eof or rs.bof then
					if isdate(collect_due_date) then
						sql="insert into saupbu_sales (sales_date,approve_no,slip_no,sales_company,saupbu,company,group_name,trade_no,sales_amt,cost_amt,vat_amt,collect_due_date,emp_no,emp_name,sales_memo,reg_id,reg_name,reg_date) values ('"&sales_date&"','"&approve_no&"','"&slip_no&"','"&sales_company&"','"&saupbu&"','"&trade_name&"','"&group_name&"','"&trade_no&"',"&price&","&cost&","&cost_vat&",'"&collect_due_date&"','"&emp_no&"','"&emp_name&"','"&sales_memo&"','"&user_id&"','"&user_name&"',now())"
					  else
						sql="insert into saupbu_sales (sales_date,approve_no,slip_no,sales_company,saupbu,company,group_name,trade_no,sales_amt,cost_amt,vat_amt,emp_no,emp_name,sales_memo,reg_id,reg_name,reg_date) values ('"&sales_date&"','"&approve_no&"','"&slip_no&"','"&sales_company&"','"&saupbu&"','"&trade_name&"','"&group_name&"','"&trade_no&"',"&price&","&cost&","&cost_vat&",'"&emp_no&"','"&emp_name&"','"&sales_memo&"','"&user_id&"','"&user_name&"',now())"
                    end if
'Response.write sql&"<br>"
					dbconn.execute(sql)
					write_cnt = write_cnt + 1
				end if
			end if
		next
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = "변경중 Error가 발생하였습니다...."
	else
		dbconn.CommitTrans
	end if

	sales_saupbu = "전체"
	field_check = "total"
	field_view = ""
	ck_sw = "y"
	url = "/sales_bill_mg.asp?sales_month="&sales_month&"&sales_saupbu="&sales_saupbu&"&field_check="&field_check&"&field_view="&field_view&"&ck_sw="&ck_sw
	err_msg = "총 " + cstr(read_cnt) + "건 읽고 " + cstr(write_cnt) + " 건 처리되었습니다..."
	response.write"<script language=javascript>"
	response.write"alert('"&err_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	rs.close
	cn.close
	rs_etc.close
	set rs = nothing
	set cn = nothing
	set rs_etc = nothing
%>