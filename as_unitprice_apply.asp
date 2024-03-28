<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
    apply_year  = request.form("apply_year")
    apply_month = request.form("apply_month")

    au_month = apply_year & apply_month ' YYYYMM
    au_month_date = apply_year & "-" & apply_month & "-01" ' YYYY-MM-01

    set dbconn = server.CreateObject("adodb.connection")
    Set rs = Server.CreateObject("ADODB.Recordset")

    Dbconn.open dbconnect

    Sql = "SELECT count(*) AS month_cnt           " & chr(13) & _
          "  FROM as_unitprice_month              " & chr(13) & _
          " WHERE au_month >= '" & au_month  & "' "
    'Response.Write  "<pre>" & Sql &"</pre>"
    Set rs = DbConn.Execute(Sql)

    if not rs.eof then
        month_cnt = cint(rs("month_cnt"))
        rs.Close()

        ' 월 표준단가가 있는 경우
        if  (month_cnt > 0) then
            Sql = "DELETE FROM as_unitprice_month               " & chr(13) & _
                  "      WHERE au_month >= '" & au_month  & "'  "
            dbconn.execute(Sql)      
            'Response.Write  "<pre>" & Sql &"</pre>"
        end if
    
        curr_date = mid(cstr(now()),1,10) ' 현재일자
        mon_cnt = datediff("m", au_month_date, curr_date) ' 현재 달까지의 달수
        'Response.Write  "<pre>" & mon_cnt &"</pre>"

        startDate = CDate(au_month_date) 

        Sql = "  SELECT au_code          " & chr(13) & _
              "       , au_name          " & chr(13) & _
              "       , cost_center      " & chr(13) & _
              "       , as_unitprice1    " & chr(13) & _
              "       , as_unitprice2    " & chr(13) & _
              "    FROM as_unitprice     " & chr(13) & _
              "   WHERE delete_yn = 'N'  "
        rs.Open Sql, Dbconn, 1
        'Response.Write  "<pre>" & Sql &"</pre>"

        do until rs.eof 

            for i = 0 to mon_cnt
                iDate = DateAdd("m",i,startDate)                
                sDate = mid(cstr(iDate),1,4) & mid(cstr(iDate),6,2) ' YYYYMM
                'Response.Write  "<pre>" & sDate &"</pre>"

                Sql = " INSERT INTO as_unitprice_month ( au_month                       " & chr(13) &_
                      "                                , au_code                        " & chr(13) &_
                      "                                , au_name                        " & chr(13) &_
                      "                                , cost_center                    " & chr(13) &_
                      "                                , as_unitprice1                  " & chr(13) &_
                      "                                , as_unitprice2                  " & chr(13) &_
                      "                                )                                " & chr(13) &_
                      "                         VALUES ( '" & sDate & "'                " & chr(13) &_
                      "                                , '" & rs("au_code") & "'        " & chr(13) &_
                      "                                , '" & rs("au_name") & "'        " & chr(13) &_
                      "                                , '" & rs("cost_center") & "'    " & chr(13) &_
                      "                                , '" & rs("as_unitprice1") & "'  " & chr(13) &_
                      "                                , '" & rs("as_unitprice2") & "'  " & chr(13) &_
                      "                                )                                "
                'Response.Write  "<pre>" & Sql &"</pre>"
                dbconn.execute(Sql)                
            next

            rs.movenext()  
        loop 
        rs.Close()

    end if
    

    Response.Write "<script language=javascript>"

    massage = escape("월별 표준단가가 지정되었습니다....")
    
    'Response.Write "alert('월별 표준단가가 지정되었습니다....');"		
    'Response.write "history.go(-1);"
    Response.Redirect "as_unitprice_mg.asp?message="& massage
    
	Response.Write "</script>"
	'Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>
