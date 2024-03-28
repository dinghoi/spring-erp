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

slip_month = abc("slip_month")
objFile = abc("objFile")

Set DbConn = Server.CreateObject("ADODB.Connection")
DbConn.Open dbconnect

set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")

dbconn.BeginTrans

cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
rs.Open "select * from [1:10000]",cn,"0"

'날짜값을 입력받아 원하는 포멧으로 변경하는 함수
'입력값 : now()
'출력값 : 20080101000000
Function ConvertDateFormat(ByVal strDate)
	Dim tmpDate1, tmpDate2
	Dim returnDate
	tmpDate1 = Split(strDate, " ")
	tmpDate2 = Split(tmpDate1(2), ":")

	'오후라면 12시간을 더해준다
	If tmpDate1(1) = "오후" Then
		'오후 12시는 정오를 가르키기 때문에 제외
		If CDbl(tmpDate2(0)) < 12 Then
			tmpDate2(0) = CDbl(tmpDate2(0)) + 12
		End If
	End If

	returnDate = tmpDate1(0)& " " & CheckFormat(tmpDate2(0),2) & ":" & CheckFormat(tmpDate2(1),2) & ":" & CheckFormat(tmpDate2(2),2)
	ConvertDateFormat = returnDate
End Function

'자릿수를 맞추기 위한 함수
Function CheckFormat(ByVal num, ByVal splitpos)
	Dim tmpNum : tmpNum = 10000000
	tmpNum = tmpNum + CDbl(num)
	CheckFormat = Right(tmpNum, splitpos)
End Function

rowcount=-1
xgr = rs.getrows
rowcount = ubound(xgr,2)
fldcount = rs.fields.count

tot_cnt = rowcount + 1
read_cnt = 0
write_cnt = 0
slip_gubun = "비용"
pay_method = "현금"
vat_yn = "N"
sign_no = ""
cancel_yn = "N"
end_yn = "N"
reg_id   = user_id
reg_user = user_name
reg_date = ConvertDateFormat(Now()) ' yyyy-mm-dd HH:mm:ss

if rowcount > -1 then
	for i=0 to rowcount
		if xgr(1,i) = "" or isnull(xgr(1,i)) then
			exit for
		end If

		slip_date = xgr(0,i)
		org_name = xgr(1,i)
		emp_name = xgr(2,i)
		emp_company = xgr(3,i)
		account_name = xgr(4,i)
		account_item = xgr(5,i)
		price = toString(xgr(6,i), 0)
		company = xgr(7,i)
		customer = xgr(8,i)
		pay_yn = xgr(9,i)
		slip_memo = xgr(10,i)
		pl_yn = xgr(11,i)

		sql = "select max(slip_seq) as max_seq from general_cost where slip_date='" & slip_date & "'"
		set rsSeq = dbconn.execute(sql)

		if	isnull(rsSeq("max_seq"))  then
			slip_seq = "001"
		else
			max_seq = "00" & cstr((int(rsSeq("max_seq")) + 1))
			slip_seq = right(max_seq,3)
		end If
		rsSeq.Close()

		SQL = "SELECT org_bonbu, org_saupbu, org_team, org_reside_place "
		SQL = SQL & "FROM emp_org_mst "
		SQL = SQL & "WHERE org_name = '"&org_name&"' AND org_company = '"&emp_company&"' "
		Set rsOrg = DBConn.Execute(SQL)

		bonbu = rsOrg("org_bonbu")
		saupbu = rsOrg("org_saupbu")
		team = rsOrg("org_team")
		reside_place = rsOrg("org_reside_place")

		rsOrg.Close()

		if vat_yn = "Y" then
			cost     = price / 1.1
			cost_vat = cost * 0.1
			cost_vat = round(cost_vat,0)
			cost     = price - cost_vat
		else
			cost_vat = 0
			cost     = price
		end If

		sql = "insert into general_cost (slip_date              "&chr(13)&_
              "                         ,slip_seq               "&chr(13)&_
              "                         ,slip_gubun             "&chr(13)&_
              "                         ,emp_company            "&chr(13)&_
              "                         ,bonbu                  "&chr(13)&_
              "                         ,saupbu                 "&chr(13)&_
              "                         ,team                   "&chr(13)&_
              "                         ,org_name               "&chr(13)&_
              "                         ,reside_place           "&chr(13)&_
              "                         ,company                "&chr(13)&_
              "                         ,account                "&chr(13)&_
              "                         ,account_item           "&chr(13)&_
              "                         ,pay_method             "&chr(13)&_
              "                         ,price                  "&chr(13)&_
              "                         ,cost                   "&chr(13)&_
              "                         ,vat_yn                 "&chr(13)&_
              "                         ,cost_vat               "&chr(13)&_
              "                         ,customer               "&chr(13)&_
              "                         ,sign_no                "&chr(13)&_
              "                         ,emp_name               "&chr(13)&_
              "                         ,emp_no                 "&chr(13)&_
              "                         ,emp_grade              "&chr(13)&_
              "                         ,pay_yn                 "&chr(13)&_
              "                         ,slip_memo              "&chr(13)&_
              "                         ,tax_bill_yn            "&chr(13)&_
              "                         ,cost_reg               "&chr(13)&_
              "                         ,cancel_yn              "&chr(13)&_
              "                         ,end_yn                 "&chr(13)&_
              "                         ,reg_id                 "&chr(13)&_
              "                         ,reg_user               "&chr(13)&_
              "                         ,reg_date               "&chr(13)&_
              "                         ,pl_yn                  "&chr(13)&_
              "                         )                       "&chr(13)&_
              "                  values ('"&slip_date&"'        "&chr(13)&_
              "                         ,'"&slip_seq&"'         "&chr(13)&_
              "                         ,'"&slip_gubun&"'       "&chr(13)&_
              "                         ,'"&emp_company&"'      "&chr(13)&_
              "                         ,'"&bonbu&"'            "&chr(13)&_
              "                         ,'"&saupbu&"'           "&chr(13)&_
              "                         ,'"&team&"'             "&chr(13)&_
              "                         ,'"&org_name&"'         "&chr(13)&_
              "                         ,'"&reside_place&"'     "&chr(13)&_
              "                         ,'"&company&"'          "&chr(13)&_
              "                         ,'"&account_name&"'     "&chr(13)&_
              "                         ,'"&account_item&"'     "&chr(13)&_
              "                         ,'"&pay_method&"'       "&chr(13)&_
              "                         ,"&price&"              "&chr(13)&_
              "                         ,"&cost&"               "&chr(13)&_
              "                         ,'"&vat_yn&"'           "&chr(13)&_
              "                         ,"&cost_vat&"           "&chr(13)&_
              "                         ,'"&customer&"'         "&chr(13)&_
              "                         ,'"&sign_no&"'          "&chr(13)&_
              "                         ,'"&emp_name&"'         "&chr(13)&_

              "                         ,'"&user_id&"'				"&chr(13)&_
              "                         ,'"&user_grade&"'				"&chr(13)&_

              "                         ,'"&pay_yn&"'           "&chr(13)&_
              "                         ,'"&slip_memo&"'        "&chr(13)&_
              "                         ,'N'                    "&chr(13)&_
              "                         ,'0'                    "&chr(13)&_
              "                         ,'"&cancel_yn&"'        "&chr(13)&_
              "                         ,'"&end_yn&"'           "&chr(13)&_
              "                         ,'"&reg_id&"'           "&chr(13)&_
              "                         ,'"&reg_user&"'         "&chr(13)&_
              "                         ,'"&reg_date&"'         "&chr(13)&_
              "                         ,'"&pl_yn&"'            "&chr(13)&_
              "                         );                       "&chr(13)

			dbconn.execute(sql)
			write_cnt = write_cnt + 1

			'Response.write sql&"<br/>"
	Next
	'Response.end
end if

if Err.number <> 0 then
	dbconn.RollbackTrans
	end_msg = "변경중 Error가 발생하였습니다."
else
	dbconn.CommitTrans
end if

sales_saupbu = "전체"

url = "/general_cost_mg.asp?slip_month="&slip_month
err_msg = "총 " &  cstr(write_cnt) & " 건 처리되었습니다."

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&err_msg&"');"
Response.write "	location.replace('"&url&"');"
Response.write "</script>"
Response.End

rs.Close() : set rs = nothing
cn.Close() : set cn = nothing
DBConn.Close() : Set DBConn = Nothing
%>