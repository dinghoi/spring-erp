<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
    u_type        = request.form("u_type")
    au_code       = request.form("au_code")
    au_name       = request.form("au_name")
    cost_center   = request.form("cost_center")
    as_unitprice1 = int(request.form("as_unitprice1"))
    as_unitprice2 = int(request.form("as_unitprice2"))

	set dbconn = server.CreateObject("adodb.connection")

    Dbconn.open dbconnect


	if u_type = "U" then
        sql = "UPDATE as_unitprice                            " & chr(13) & _
              "   SET au_name       = '" & au_name & "'       " & chr(13) & _    
              "     , cost_center   = '" & cost_center & "'   " & chr(13) & _        
              "     , as_unitprice1 = '" & as_unitprice1 & "' " & chr(13) & _          
              "     , as_unitprice2 = '" & as_unitprice2 & "' " & chr(13) & _          
              " WHERE au_code = '" & au_code & "'             "
    elseif u_type = "D" then
        ' 삭제처리
        sql = "UPDATE as_unitprice                 " & chr(13) & _
              "   SET delete_yn = 'Y'              " & chr(13) & _          
              " WHERE au_code = '" & au_code & "'  "
    else
        sql = "DELETE FROM as_unitprice                 " & chr(13) & _
              "      WHERE au_code = '" & au_code & "'  "
        dbconn.execute(sql)          

        sql = "INSERT INTO as_unitprice ( au_code                   " & chr(13) & _       
              "                         , au_name                   " & chr(13) & _    
              "                         , cost_center               " & chr(13) & _        
              "                         , as_unitprice1             " & chr(13) & _          
              "                         , as_unitprice2             " & chr(13) & _          
              "                         )                           " & chr(13) & _          
              "                  VALUES ( '" & au_code & "'         " & chr(13) & _       
              "                         , '" & au_name & "'         " & chr(13) & _    
              "                         , '" & cost_center & "'     " & chr(13) & _        
              "                         , '" & as_unitprice1 & "'   " & chr(13) & _          
              "                         , '" & as_unitprice2 & "'   " & chr(13) & _          
              "                         )                           "        
    end if	
    'Response.Write  "<pre>"&sql &"</pre>"
    dbconn.execute(sql)

    Response.Write "<script language=javascript>"
    
    if u_type = "D" then
        massage = escape("삭제 완료 되었습니다....")
    else
        massage = escape("등록 완료 되었습니다....")
    end if
    
    if u_type = "U" then
        Response.Redirect "as_unitprice_mg.asp?au_code="& au_code &"&u_type="& u_type &"&message="& massage
    else
        Response.Redirect "as_unitprice_mg.asp?message="& massage
    end if
	Response.Write "</script>"
	'Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>
