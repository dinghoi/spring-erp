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

	bill_id = abc("bill_id")
	bill_month = abc("bill_month")
	objFile = abc("objFile")

	from_date = mid(bill_month,1,4) + "-" + mid(bill_month,5,2) + "-01"
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
				bill_date = xgr(0,i)
				approve_no = xgr(1,i)
				if bill_id = "1" then
					trade_no = xgr(4,i)
					trade_name = replace(xgr(6,i)," ","")
					trade_owner = xgr(7,i)
					owner_trade_no = xgr(8,i)
				  else
					trade_no = xgr(8,i)
					trade_name = replace(xgr(10,i)," ","")
					trade_owner = xgr(11,i)
					owner_trade_no = xgr(4,i)
				end if
				price = int(xgr(12,i))
				cost = int(xgr(13,i))
				cost_vat = int(xgr(14,i))
				bill_collect = xgr(19,i)
				send_email = xgr(20,i)
				receive_email = xgr(21,i)
				imsi_bill_memo = xgr(24,i)
				tax_bill_memo = replace(imsi_bill_memo,"'","")

				tradename = Replace(trade_name,"'","&quot;")
				trade_name = replace(tradename,"（주）","(주)")
				tradename = trade_name
				trade_name = replace(tradename,"㈜","(주)")
				owner_trade_no = Replace(owner_trade_no,"-","")
				trade_no = Replace(trade_no,"-","")

				sql = "select * from trade where trade_no = '"&owner_trade_no&"'"
				set rs_trade=dbconn.execute(sql)
				if rs_trade.eof or rs_trade.bof then
					owner_company = owner_trade_no + "_Error"
				  else
					owner_company = rs_trade("trade_name")
				end if
				rs_trade.close()

				read_cnt = read_cnt + 1

				sql = "select * from tax_bill where approve_no = '"&approve_no&"'"
				set rs=dbconn.execute(sql)
				if rs.eof or rs.bof then
					sql="insert into tax_bill (approve_no,bill_id,bill_date,owner_trade_no,owner_company,trade_no,trade_name,trade_owner,price,cost,cost_vat,bill_collect,send_email,receive_email,tax_bill_memo,reg_id,reg_name,reg_date) values ('"&approve_no&"','"&bill_id&"','"&bill_date&"','"&owner_trade_no&"','"&owner_company&"','"&trade_no&"','"&trade_name&"','"&trade_owner&"',"&price&","&cost&","&cost_vat&",'"&bill_collect&"','"&send_email&"','"&receive_email&"','"&tax_bill_memo&"','"&user_id&"','"&user_name&"',now())"
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

	err_msg = "총 " + cstr(read_cnt) + "건 읽고 " + cstr(write_cnt) + " 건 처리되었습니다..."
	response.write"<script language=javascript>"
	response.write"alert('"&err_msg&"');"
	response.write"location.replace('tax_esero_mg.asp');"
	response.write"</script>"
	Response.End

	rs.close
	cn.close
	rs_etc.close
	set rs = nothing
	set cn = nothing
	set rs_etc = nothing
%>