<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
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



'	on Error resume next

	u_type       = request.form("u_type")
	slip_seq     = request.form("slip_seq")
	slip_date    = request.form("slip_date")
	old_date     = request.form("old_date")
	emp_company  = request.form("emp_company")
	bonbu        = request.form("bonbu")
	saupbu       = request.form("saupbu")
	team         = request.form("team")
	org_name     = request.form("org_name")
	reside_place = request.form("reside_place")
	emp_name     = request.form("emp_name")
	emp_no       = request.form("emp_no")
	emp_grade    = request.form("emp_grade")
	accountitem  = request.form("account")
	i            = instr(1,accountitem,"-")
	account      = mid(accountitem,1,i-1)
	account_item = mid(accountitem,i+1)
	pay_method   = request.form("pay_method")
	price        = int(request.form("price"))
'	vat_yn       = request.form("vat_yn")
	vat_yn       = "N"
	customer     = request.form("customer")
	company      = request.form("company")
'	emp_no       = request.form("emp_no")
'	pay_yn       = "N"
	pay_yn       = request.form("pay_yn")
	sign_no      = request.form("sign_no")
	slip_memo    = request.form("slip_memo")
	end_yn       = request.form("end_yn")
	cancel_yn    = request.form("cancel_yn")
	if vat_yn = "Y" then
		cost     = price / 1.1
		cost_vat = cost * 0.1
		cost_vat = round(cost_vat,0)
		cost     = price - cost_vat
	else
	  	cost_vat = 0
		cost     = price
	end if
	mod_id   = request.form("mod_id")
	mod_user = request.form("mod_user")
	mod_date = request.form("mod_date")

	if mod_id <> "" then
		mod_yymmdd = datevalue(mod_date)
		mod_hhmm = formatdatetime(mod_date,4)
		mod_date = cstr(mod_yymmdd) + " " + cstr(mod_hhmm)
	end if
	pl_yn = request.form("pl_yn")

	slip_gubun = "비용"

	set dbconn = server.CreateObject("adodb.connection")
    Set rs = Server.CreateObject("ADODB.Recordset")
    Set obj_rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

    reg_id   = user_id
    reg_user = user_name
    reg_date = ConvertDateFormat(Now()) ' yyyy-mm-dd HH:mm:ss

	if	u_type = "U" then
        sql = "select reg_id                                              "&chr(13)&_
              "     , reg_user                                            "&chr(13)&_
              "     , date_format(reg_date, '%Y-%m-%d %H:%i:%s') reg_date "&chr(13)&_
              "  from general_cost                                        "&chr(13)&_
              " where slip_date ='"&old_date&"'                           "&chr(13)&_
              "   and slip_seq='"&slip_seq&"'                             "&chr(13)
        set obj_rs = dbconn.execute(sql)
        if Not obj_rs.EOF then
            reg_id   = obj_rs("reg_id")
            reg_user = obj_rs("reg_user")
            reg_date = obj_rs("reg_date")
        End if

        sql = "delete from general_cost where slip_date ='"&old_date&"' and slip_seq='"&slip_seq&"'"
		dbconn.execute(sql)
    end if

	sql="select max(slip_seq) as max_seq from general_cost where slip_date='" + slip_date + "'"
	set rs=dbconn.execute(sql)
		
	if	isnull(rs("max_seq"))  then
		slip_seq = "001"
	else
		max_seq = "00" + cstr((int(rs("max_seq")) + 1))
		slip_seq = right(max_seq,3)
	end if

    if	u_type = "U" then
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
              "                         ,mod_id                 "&chr(13)&_
              "                         ,mod_user               "&chr(13)&_
              "                         ,mod_date               "&chr(13)&_
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
              "                         ,'"&account&"'          "&chr(13)&_
              "                         ,'"&account_item&"'     "&chr(13)&_
              "                         ,'"&pay_method&"'       "&chr(13)&_
              "                         ,"&price&"              "&chr(13)&_
              "                         ,"&cost&"               "&chr(13)&_
              "                         ,'"&vat_yn&"'           "&chr(13)&_
              "                         ,"&cost_vat&"           "&chr(13)&_
              "                         ,'"&customer&"'         "&chr(13)&_
              "                         ,'"&sign_no&"'          "&chr(13)&_
              "                         ,'"&emp_name&"'         "&chr(13)&_
              "                         ,'"&emp_no&"'           "&chr(13)&_
              "                         ,'"&emp_grade&"'        "&chr(13)&_
              "                         ,'"&pay_yn&"'           "&chr(13)&_
              "                         ,'"&slip_memo&"'        "&chr(13)&_
              "                         ,'N'                    "&chr(13)&_
              "                         ,'0'                    "&chr(13)&_
              "                         ,'"&cancel_yn&"'        "&chr(13)&_
              "                         ,'"&end_yn&"'           "&chr(13)&_
              "                         ,'"&reg_id&"'           "&chr(13)&_
              "                         ,'"&reg_user&"'         "&chr(13)&_
              "                         ,'"&reg_date&"'         "&chr(13)&_
              "                         ,'"&user_id&"'           "&chr(13)&_
              "                         ,'"&user_name&"'         "&chr(13)&_
              "                         ,now()                  "&chr(13)&_
              "                         ,'"&pl_yn&"'            "&chr(13)&_
              "                         )                       "&chr(13)
        dbconn.execute(sql)
    else
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
              "                         ,'"&account&"'          "&chr(13)&_
              "                         ,'"&account_item&"'     "&chr(13)&_
              "                         ,'"&pay_method&"'       "&chr(13)&_
              "                         ,"&price&"              "&chr(13)&_
              "                         ,"&cost&"               "&chr(13)&_
              "                         ,'"&vat_yn&"'           "&chr(13)&_
              "                         ,"&cost_vat&"           "&chr(13)&_
              "                         ,'"&customer&"'         "&chr(13)&_
              "                         ,'"&sign_no&"'          "&chr(13)&_
              "                         ,'"&emp_name&"'         "&chr(13)&_
              "                         ,'"&emp_no&"'           "&chr(13)&_
              "                         ,'"&emp_grade&"'        "&chr(13)&_
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
              "                         )                       "&chr(13)
        dbconn.execute(sql)
    end if    

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	Response.write"<script language=javascript>"
	Response.write"alert('"&end_msg&"');"
	Response.write"opener.document.frm.submit();"
	Response.write"window.close();"		
	Response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
