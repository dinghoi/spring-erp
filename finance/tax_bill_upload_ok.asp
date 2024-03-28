<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'On Error Resume Next
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn=Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder=New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim uploadForm, bill_id, bill_month, objFile, from_date, end_date, to_date
Dim cn, rs, rowcount, xgr, fldcount, tot_cnt, read_cnt, write_cnt
Dim rs_trade, rsTax, i, bill_date, approve_no, trade_no, trade_name
Dim trade_owner, owner_trade_no, price, cost, cost_vat, bill_collect
Dim send_email, receive_email, imsi_bill_memo, tax_bill_memo
Dim owner_company, err_msg, url

'업로드 컴포넌트 객체 생성
Set uploadForm = Server.CreateObject("ABCUpload4.XForm")
uploadForm.AbsolutePath = True
uploadForm.Overwrite = true
uploadForm.MaxUploadSize = 1024*1024*50

bill_id = uploadForm("bill_id")
bill_month = uploadForm("bill_month")
objFile = uploadForm("objFile")

from_date = Mid(bill_month,1,4)&"-"&Mid(bill_month,5,2)&"-01"
end_date = DateValue(from_date)
end_date = DateAdd("m",1,from_date)
to_date = CStr(DateAdd("d",-1,end_date))

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
rs.Open "select * from [6:10000]",cn,"0"

rowcount=-1
xgr = rs.getrows
rowcount = UBound(xgr,2)
fldcount = rs.fields.count

tot_cnt = rowcount + 1
read_cnt = 0
write_cnt = 0

'DB 트랜잭션 시작
DBConn.BeginTrans

If rowcount > -1 Then
	For i = 0 To rowcount
		If f_toString(xgr(1,i), "") = "" Then
			Exit For
		End If

		bill_date = xgr(0,i)'작성일자
		approve_no = xgr(1,i)'승인번호

		If bill_id = "1" Then
			trade_no = xgr(4,i)'사업자등록번호
			trade_name = Replace(xgr(6,i)," ","")'상호(공급자)
			trade_owner = xgr(7,i)'대표자명(공급자)
			owner_trade_no = xgr(9,i)'공급받는자사업자등록번호
		Else
			owner_trade_no = xgr(4,i)'사업자등록번호
			trade_no = xgr(9,i)'공급받는자사업자등록번호
			trade_name = Replace(xgr(11,i)," ","")'상호(공급받는자)
			trade_owner = xgr(12,i)'대표자명(공급받는자)
		End If

		price = Int(xgr(14,i))'합계급액
		cost = Int(xgr(15,i))'공급가액
		cost_vat = Int(xgr(16,i))'세액
		bill_collect = xgr(19,i)'영수/청구 구분
		send_email = xgr(22,i)'공급자이메일
		receive_email = xgr(23,i)'공급받는자이메일1

		imsi_bill_memo = xgr(26,i)'품목명

		tax_bill_memo = Replace(imsi_bill_memo,"'","")

		'tradename = Replace(trade_name,"'","&quot;")
		'trade_name = Replace(tradename,"（주）","(주)")
		'tradename = trade_name
		'trade_name = replace(tradename,"㈜","(주)")

		trade_name = Replace(Replace(Replace(trade_name, "'", "&quot;"), "（주）","(주)"), "㈜","(주)")

		owner_trade_no = Replace(owner_trade_no,"-","")
		trade_no = Replace(trade_no,"-","")

		'sql = "select * from trade where trade_no = '"&owner_trade_no&"'"
		objBuilder.Append "SELECT trade_name FROM trade WHERE trade_no='"&owner_trade_no&"';"

		Set rs_trade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_trade.EOF Or rs_trade.BOF Then
			owner_company = owner_trade_no&"_Error"
		Else
			owner_company = rs_trade("trade_name")
		End If
		rs_trade.Close()

		read_cnt = read_cnt + 1

		'sql = "select * from tax_bill where approve_no = '"&approve_no&"'"
		objBuilder.Append "SELECT approve_no FROM tax_bill WHERE approve_no='"&approve_no&"';"

		Set rsTax = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsTax.EOF Or rsTax.BOF Then
			'sql="insert into tax_bill (approve_no,bill_id,bill_date,owner_trade_no,owner_company,trade_no,trade_name,trade_owner,price,cost,cost_vat,bill_collect,send_email,receive_email,tax_bill_memo,reg_id,reg_name,reg_date) values ('"&approve_no&"','"&bill_id&"','"&bill_date&"','"&owner_trade_no&"','"&owner_company&"','"&trade_no&"','"&trade_name&"','"&trade_owner&"',"&price&","&cost&","&cost_vat&",'"&bill_collect&"','"&send_email&"','"&receive_email&"','"&tax_bill_memo&"','"&user_id&"','"&user_name&"',now())"

			objBuilder.Append "INSERT INTO tax_bill(approve_no,bill_id,bill_date,owner_trade_no,owner_company,"
			objBuilder.Append "trade_no,trade_name,trade_owner,price,cost,"
			objBuilder.Append "cost_vat,bill_collect,send_email,receive_email,tax_bill_memo,"
			objBuilder.Append "reg_id,reg_name,reg_date"
			objBuilder.Append ")VALUES("
			objBuilder.Append "'"&approve_no&"','"&bill_id&"','"&bill_date&"','"&owner_trade_no&"','"&owner_company&"',"
			objBuilder.Append "'"&trade_no&"','"&trade_name&"','"&trade_owner&"',"&price&","&cost&","
			objBuilder.Append cost_vat&",'"&bill_collect&"','"&send_email&"','"&receive_email&"','"&tax_bill_memo&"',"
			objBuilder.Append "'"&user_id&"','"&user_name&"',NOW());"
		Else
			objBuilder.Append "UPDATE tax_bill SET "
			objBuilder.Append "	bill_id='"&bill_id&"',bill_date='"&bill_date&"',owner_trade_no='"&owner_trade_no&"',owner_company='"&owner_company&"', "
			objBuilder.Append "	trade_no='"&trade_no&"',trade_name='"&trade_name&"',trade_owner='"&trade_owner&"',price="&price&",cost="&cost&", "
			objBuilder.Append "	cost_vat="&cost_vat&",bill_collect="&bill_collect&"',send_email='"&send_email&"',receive_email='"&receive_email&"', "
			objBuilder.Append "	tax_bill_memo='"&tax_bill_memo&"', mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=NOW() "
			objBuilder.Append "WHERE approve_no='"&approve_no&"';"
		End If

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'처리 개수
		write_cnt = write_cnt + 1
	Next
End If
Set rs_trade = Nothing
Set rsTax = Nothing

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "변경 중 Error가 발생하였습니다."
	url = "/finance/tax_bill_upload.asp"
Else
	DBConn.CommitTrans
	err_msg = "총 "&CStr(read_cnt)&"건 읽고 "&CStr(write_cnt)&" 건 처리되었습니다."
	url = "/finance/tax_esero_mg.asp?bill_id="&bill_id&"&bill_month="&bill_month&"&cost_reg_yn=T&end_yn=T"
End If

rs.Close() : Set rs = Nothing
cn.Close() : Set cn = Nothing
DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&err_msg&"');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>