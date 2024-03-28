<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
    On Error Resume Next

	u_type        = request.form("u_type")
	del_sw        = request.form("del_sw")
	card_type     = request.form("card_type")
	card_no       = request.form("card_no")
	owner_company = request.form("owner_company")
'	card_no1      = request.form("card_no1")
'	card_no2      = request.form("card_no2")
'	card_no3      = request.form("card_no3")
'	card_no4      = request.form("card_no4")
'	card_no       = card_no1 + "-" + card_no2 + "-" + card_no3 + "-" + card_no4
	emp_no        = request.form("emp_no")
'	emp_name      = request.form("emp_name")
	card_issue    = request.form("card_issue")
	card_limit    = request.form("card_limit")
	valid_thru    = request.form("valid_thru")
	create_date   = request.form("create_date")
	start_date    = request.form("start_date")
	card_memo     = request.form("card_memo")
	car_vat_sw    = request.form("car_vat_sw")
    use_yn        = request.form("use_yn")
    pl_yn         = request.form("pl_yn")
	mod_id        = request.form("mod_id")
	mod_name      = request.form("mod_name")
	mod_date      = request.form("mod_date")

	if mod_id <> "" then
		mod_yymmdd = datevalue(mod_date)
		mod_hhmm = formatdatetime(mod_date,4)
		mod_date = cstr(mod_yymmdd) + " " + cstr(mod_hhmm)
	end if

	dbconn.BeginTrans

	Sql="SELECT * FROM memb WHERE user_id = '"&emp_no&"'"
	Set rs_emp = DbConn.Execute(Sql)
	emp_name = rs_emp("user_name")

	if	u_type = "U" or del_sw = "Y" then
		sql = "DELETE FROM card_owner WHERE card_no ='"&card_no&"' "
		dbconn.execute(sql)
	end if

	if	del_sw <> "Y" then
        sql = "INSERT INTO card_owner ( card_no              "&chr(13)&_
            "                         , card_type            "&chr(13)&_
            "                         , owner_company        "&chr(13)&_
            "                         , emp_no               "&chr(13)&_
            "                         , emp_name             "&chr(13)&_
            "                         , card_issue           "&chr(13)&_
            "                         , card_limit           "&chr(13)&_
            "                         , valid_thru           "&chr(13)&_
            "                         , create_date          "&chr(13)&_
            "                         , start_date           "&chr(13)&_
            "                         , card_memo            "&chr(13)&_
            "                         , car_vat_sw           "&chr(13)&_
            "                         , use_yn               "&chr(13)&_
            "                         , pl_yn                "&chr(13)&_
            "                         , reg_id               "&chr(13)&_
            "                         , reg_name             "&chr(13)&_
            "                         , reg_date             "&chr(13)&_
            "                         )                      "&chr(13)&_
            "                  VALUES ( '"&card_no&"'        "&chr(13)&_
            "                         , '"&card_type&"'      "&chr(13)&_
            "                         , '"&owner_company&"'  "&chr(13)&_
            "                         , '"&emp_no&"'         "&chr(13)&_
            "                         , '"&emp_name&"'       "&chr(13)&_
            "                         , '"&card_issue&"'     "&chr(13)&_
            "                         , '"&card_limit&"'     "&chr(13)&_
            "                         , '"&valid_thru&"'     "&chr(13)&_
            "                         , '"&create_date&"'    "&chr(13)&_
            "                         , '"&start_date&"'     "&chr(13)&_
            "                         , '"&card_memo&"'      "&chr(13)&_
            "                         , '"&car_vat_sw&"'     "&chr(13)&_
            "                         , '"&use_yn&"'         "&chr(13)&_
            "                         , '"&pl_yn&"'          "&chr(13)&_
            "                         , '"&user_id&"'        "&chr(13)&_
            "                         , '"&user_name&"'      "&chr(13)&_
            "                         , now()                "&chr(13)&_
            "                         )                      "&chr(13)
        dbconn.execute(sql)
    end if
    
    If Err.number <> 0 Then     '오류 발생 시 이 부분 실행
        dbconn.RollbackTrans 
        end_msg = "처리중 Error가 발생하였습니다...."
        
        Response.Write "" & Err.Source & "<br>"
	    Response.Write "오류 번호 : " & Err.number & "<br>"
	    Response.Write "내용 : " & Err.Description & "<br>"
	Else
        dbconn.CommitTrans
        end_msg = "처리되었습니다...."
	End If


	Response.write"<script language=javascript>"
	Response.write"alert('"&end_msg&"');"
	Response.write"parent.opener.location.reload();"
	Response.write"window.close();"		
	Response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
