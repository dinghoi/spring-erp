<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	on Error resume next

	'��ó: http://start0.tistory.com/109 
	'------------------------------------------------
	' ��뿹 : format(123, "00000") ==> 000123
	'------------------------------------------------
	function format(ByVal szString, ByVal Expression)
		if len(szString) < len(Expression) then
			format = left(expression, len(szString)) & szString
		else
			format = szString
		end if
	end function

    u_type	               = request.form("u_type") ' U �������

    fDate	               = request.form("fDate")
    lDate	               = request.form("lDate")
	acpt_no 		       = request.form("acpt_no")
	work_item 	           = request.form("work_item")   ' �۾�����
	work_date1 	           = request.form("work_date1")
	work_date2 	           = request.form("work_date2")
	company 		       = request.form("company")
	dept 				   = request.form("dept")
	from_hh 		       = format(request.form("from_hh"),"00")
	from_mm 		       = format(request.form("from_mm"),"00")	
	from_time 	           = from_hh + from_mm
	to_hh 			       = format(request.form("to_hh"),"00")
	to_mm 			       = format(request.form("to_mm"),"00")	
	to_time 		       = to_hh + to_mm
	delta_time             = request.form("delta_time")	
	delta_minute           = request.form("delta_minute")	
	rest_time              = request.form("rest_time")	
	rest_minute            = request.form("rest_minute")	
	work_gubun 	           = request.form("work_gubun") ' �߱��׸�
	work_memo 	           = work_item
	sign_no 		       = request.form("sign_no")	
    you_yn 			       = request.form("you_yn")	
    cancel_yn		       = request.form("cancel_yn")	
	alter_timeoff_date     = request.form("alter_timeoff_date")	
	alter_timeoff_hh 	   = request.form("alter_timeoff_hh")
	alter_timeoff_mm 	   = request.form("alter_timeoff_mm")	
	alter_timeoff_minute_w = request.form("alter_timeoff_minute_w")	
	alter_timeoff_minute_d = request.form("alter_timeoff_minute_d")	
	alter_timeoff_time     = format(cstr(alter_timeoff_hh),"00") + format(cstr(alter_timeoff_mm),"00")

    if  acpt_no = "0" then  ' ������ ��Ư��
        mg_ce_id= request.form("mg_ce_id")	
    end if
	
'	cost_detail  `= work_gubun
'	if work_gubun = "���Ͼ߱�" or work_gubun = "Ư�ٹ���" or work_gubun = "Ư������" or work_gubun = "Ư�پ߱�" then
'		cost_detail = "�߱�"
'	end if
'	if work_gubun = "�����Ͼ߱�" or work_gubun = "��Ư��12�������" or work_gubun = "��Ư��13����̻�" or work_gubun = "��Ư�پ߱�" or work_gubun = "��Ư��ö��" then
'		cost_detail = "���߱�"
'	end if
	
	Set Dbconn = Server.CreateObject("ADODB.Connection")
	Set Rs    = Server.CreateObject("ADODB.Recordset")
	Set rs_as = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

    if  acpt_no <> "0" then  ' ������ �ƴ� �Ϲ� AS���� ��Ư��
        '//2017-10-31 ���CE�� ��Ư�ٺ� ��� ����
        sql = "select * from as_acpt where acpt_no =" & int(acpt_no) & " and mg_ce_id='" & user_id & "' "
        'Response.write sql & "<br>"
        rs_as.Open Sql, Dbconn, 1
        if Err.number <> 0 then		
            Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql & "]<br>"
        end if 
    
        If rs_as.bof Or rs_as.eof Then
            response.write "<script language=javascript>"
            response.write "alert('���CE�� ��Ư�ٺ� ����� �� �ֽ��ϴ�.');"
            response.write "parent.opener.location.reload();"
            response.write "self.close() ;"
            response.write "</script>"
            Response.End
    
            dbconn.Close()
            Set dbconn = Nothing
        End If
        rs_as.close()
    end if 	
	
	sql1 = "SELECT * FROM overtime_code WHERE work_gubun = '"&work_gubun&"'"
	'Response.write sql1 & "<br>"
    set rs_etc=dbconn.execute(sql1)
    if Err.number <> 0 then		
        Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql1 & "]<br>"
    end if 
	
	cost_detail  = rs_etc("cost_detail")
    overtime_amt = rs_etc("overtime_amt")
    
    '''''''''''''''''''''''''''''''''''''''''
    isFound = False
  
    if  acpt_no <> "0" then  ' ������ �ƴ� �Ϲ� AS���� ��Ư��

        ' �۾� �ο��� ��ŭ ����.
        sql2 = "SELECT * FROM ce_work where work_id = '2'  and acpt_no =" & int(acpt_no)
        'Response.write sql2 & "<br>"
        Rs.Open sql2, Dbconn, 1
        if Err.number <> 0 then		
            Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql2 & "]<br>"
        end if 

        do until rs.eof
        
            sql4 = " SELECT * FROM overtime WHERE work_date = '"&work_date1&"' AND mg_ce_id = '"&rs("mg_ce_id")&"'"
            'Response.write sql4 & "<br>"
            set rs_etc=dbconn.execute(sql4)     
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
            end if 
            
            if not rs_etc.eof then
                isFound = True
            end if 
            
            rs.movenext()
        loop  
        rs.close()
    
    else ' ���� ��Ư��
    
        if u_type <> "U" then
            sql4 = " SELECT * FROM overtime WHERE work_date = '"&work_date1&"' AND mg_ce_id = '" & mg_ce_id & "'"
            'Response.write sql4 & "<br>"
            set rs_etc=dbconn.execute(sql4)     
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
            end if 
            
            if not rs_etc.eof then
                isFound = True
            end if 
        end if 

    end if

	
	If isFound = True Then
		response.write "<script language=javascript>"
		response.write "alert('�̹� ���� ��("& work_date1 &")�� ��Ư���� ��� �Ǿ��ֽ��ϴ�.\n(�������� ��Ư���� 1�Ǹ������մϴ�.) ');"
		response.write "history.back();"
		response.write "</script>"
		Response.End
  
		dbconn.Close()
		Set dbconn = Nothing
	End If
	
    '''''''''''''''''''''''''''''''''''''''''

    dbconn.BeginTrans
	
    '''''''''''''''''''''''''''''''''''''''''

    if Len(alter_timeoff_date) = 0 then
        alter_timeoff_date_f = "NULL"
        alter_timeoff_time = "0000" ' ��ü�ް��������� ������ �ð��� 0000
    else
        alter_timeoff_date_f = "'"&alter_timeoff_date&"'"  
    end if

    if  acpt_no <> "0" then  ' ������ �ƴ� �Ϲ� AS���� ��Ư��
        ' �۾� �ο��� ��ŭ ����.
        sql2 = "SELECT * FROM ce_work where work_id = '2' and acpt_no =" & int(acpt_no)
        'Response.write sql2 & "]<br>"
        rs.Open sql2, Dbconn, 1
        if Err.number <> 0 then		
            Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql2 & "]<br>"
        end if 

        do until rs.eof
        
            sql4 = " DELETE FROM overtime WHERE  work_date = '"&work_date1&"' AND  mg_ce_id = '"&rs("mg_ce_id")&"'"
            'Response.write sql4 & "<br>"
            dbconn.execute(sql4)     
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
            end if 
		
            sql4 =  "INSERT INTO overtime                                                          "&chr(13)&_
                    "          ( work_date                                                         "&chr(13)&_
                    "          , end_date                                                          "&chr(13)&_
                    "          , mg_ce_id                                                          "&chr(13)&_
                    "          , user_name                                                         "&chr(13)&_
                    "          , user_grade                                                        "&chr(13)&_
                    "          , acpt_no                                                           "&chr(13)&_
                    "          , emp_company                                                       "&chr(13)&_
                    "          , bonbu                                                             "&chr(13)&_
                    "          , saupbu                                                            "&chr(13)&_
                    "          , team                                                              "&chr(13)&_
                    "          , org_name                                                          "&chr(13)&_
                    "          , reside_place                                                      "&chr(13)&_
                    "          , company                                                           "&chr(13)&_
                    "          , dept                                                              "&chr(13)&_
                    "          , work_item                                                         "&chr(13)&_
                    "          , from_time                                                         "&chr(13)&_
                    "          , to_time                                                           "&chr(13)&_
                    "          , delta_time                                                        "&chr(13)&_
                    "          , delta_minute                                                      "&chr(13)&_
                    "          , rest_time                                                         "&chr(13)&_
                    "          , rest_minute                                                       "&chr(13)&_
                    "          , work_gubun                                                        "&chr(13)&_
                    "          , cost_detail                                                       "&chr(13)&_
                    "          , person_amt                                                        "&chr(13)&_
                    "          , overtime_amt                                                      "&chr(13)&_
                    "          , alter_timeoff_date                                                "&chr(13)&_
                    "          , alter_timeoff_time                                                "&chr(13)&_
                    "          , alter_timeoff_minute_w                                            "&chr(13)&_           
                    "          , alter_timeoff_minute_d                                            "&chr(13)&_           
                    "          , work_memo                                                         "&chr(13)&_
                    "          , you_yn                                                            "&chr(13)&_
                    "          , cancel_yn                                                         "&chr(13)&_
                    "          , end_yn                                                            "&chr(13)&_
                    "          , reg_id                                                            "&chr(13)&_
                    "          , reg_user                                                          "&chr(13)&_
                    "          , reg_date                                                          "&chr(13)&_
                    "          )                                                                   "&chr(13)&_
                    " SELECT '"&work_date1&"' AS work_date                                         "&chr(13)&_
                    "      , '"&work_date2&"' AS end_date                                          "&chr(13)&_
                    "      , mg_ce_id                                                              "&chr(13)&_
                    "      , (SELECT user_name FROM memb WHERE user_id = mg_ce_id)  AS user_name   "&chr(13)&_
                    "      , (SELECT user_grade FROM memb WHERE user_id = mg_ce_id) AS user_grade  "&chr(13)&_
                    "      , acpt_no                                                               "&chr(13)&_
                    "      , emp_company                                                           "&chr(13)&_
                    "      , bonbu                                                                 "&chr(13)&_
                    "      , saupbu                                                                "&chr(13)&_
                    "      , team                                                                  "&chr(13)&_
                    "      , org_name                                                              "&chr(13)&_
                    "      , reside_place                                                          "&chr(13)&_
                    "      , company                                                               "&chr(13)&_
                    "      , '"&dept&"'                   AS dept                                  "&chr(13)&_
                    "      , '"&work_item&"'              AS work_item                             "&chr(13)&_ 
                    "      , '"&from_time&"'              AS from_time                             "&chr(13)&_
                    "      , '"&to_time&"'                AS to_time                               "&chr(13)&_
                    "      , '"&delta_time&"'             AS delta_time                            "&chr(13)&_
                    "      , '"&delta_minute&"'           AS delta_minute                          "&chr(13)&_
                    "      , '"&rest_time&"'              AS rest_time                             "&chr(13)&_
                    "      , '"&rest_minute&"'            AS rest_minute                           "&chr(13)&_
                    "      , '"&work_gubun&"'             AS work_gubun                            "&chr(13)&_
                    "      , '"&cost_detail&"'            AS cost_detail                           "&chr(13)&_
                    "      , person_amt                                                            "&chr(13)&_
                    "      , "&overtime_amt&"             AS overtime_amt                          "&chr(13)&_
                    "      , "&alter_timeoff_date_f&"     AS alter_timeoff_date                    "&chr(13)&_
                    "      , '"&alter_timeoff_time&"'     AS alter_timeoff_time                    "&chr(13)&_
                    "      , '"&alter_timeoff_minute_w&"' AS alter_timeoff_minute_w                "&chr(13)&_
                    "      , '"&alter_timeoff_minute_d&"' AS alter_timeoff_minute_d                "&chr(13)&_
                    "      , '"&work_memo&"'              AS work_memo                             "&chr(13)&_		
                    "      , '"&you_yn&"'                 AS you_yn                                "&chr(13)&_
                    "      , 'N'                          AS cancel_yn                             "&chr(13)&_
                    "      , 'N'                          AS end_yn                                "&chr(13)&_
                    "      , '"&user_id&"'                AS reg_id                                "&chr(13)&_
                    "      , '"&user_name&"'              AS reg_user                              "&chr(13)&_
                    "      , now()                        AS reg_date                              "&chr(13)&_
                    "   FROM ce_work                                                               "&chr(13)&_
                    "  WHERE work_id  = '2'                                                        "&chr(13)&_
                    "    and acpt_no  = "& int(acpt_no) &"                                         "&chr(13)&_
                    "    and mg_ce_id = '"& rs("mg_ce_id") &"'                                     "&chr(13)		
            'Response.write "<pre>"&sql4 & "</pre><br>"
            dbconn.execute(sql4)
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
            end if 
            
            ' 52�ð� �ʰ��� ���� ��ü�ް� �Ѻ��� �ش����� �����Ϳ� �ϰ������Ѵ�.
            sql3 = " UPDATE overtime                                              "&chr(13)&_
                   "    SET alter_timeoff_minute_w = '"&alter_timeoff_minute_w&"' "&chr(13)&_
                   "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'         "&chr(13)&_
                   "    AND mg_ce_id = '"& rs("mg_ce_id") &"'                     "
            'Response.write "<pre>"&sql3 & "</pre><br>"
            dbconn.execute(sql3)	   
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql3 & "]<br>"
            end if 
            
            rs.movenext()
        loop   
        rs.close()                                    		

        ' �۾��ο����� ������� acpt_no�� �ش��ϴ� as_acpt�� ��Ư���� �����Ѵ�.
        sql5 = "UPDATE as_acpt SET overtime ='Y' WHERE acpt_no = " & int(acpt_no)
        'Response.write sql5 & "<br>"
        dbconn.execute(sql5)
        if Err.number <> 0 then		
            Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql5 & "]<br>"
        end if 

    else  ' ���� ��Ư��

        sql4 = " SELECT user_name                   "&chr(13)&_   
               "      , user_grade                  "&chr(13)&_   
               "      , emp_company                 "&chr(13)&_   
               "      , bonbu                       "&chr(13)&_   
               "      , saupbu                      "&chr(13)&_   
               "      , team                        "&chr(13)&_   
               "      , org_name                    "&chr(13)&_   
               "      , reside_place                "&chr(13)&_   
               "      , reside_company              "&chr(13)&_   
               "   FROM memb                        "&chr(13)&_   
               "  WHERE user_id = " & mg_ce_id & "  "&chr(13)
        'Response.write sql4 & "<br>"
        rs.Open sql4, Dbconn, 1
        if Err.number <> 0 then		
            Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
        end if 

        if not (rs.eof or rs.bof) then
            mmb_user_name       = rs("user_name")      
            mmb_user_grade      = rs("user_grade")      
            mmb_emp_company     = rs("emp_company")      
            mmb_bonbu           = rs("bonbu")      
            mmb_saupbu          = rs("saupbu")      
            mmb_team            = rs("team")      
            mmb_org_name        = rs("org_name")      
            mmb_reside_place    = rs("reside_place")      
            mmb_reside_company  = rs("reside_company")      
        end if
        rs.close()

        if u_type <> "U" then    

            Sql4 = "SELECT count(*) FROM overtime WHERE  work_date = '"&work_date1&"' AND  mg_ce_id = '" & mg_ce_id & "'"
            rs.Open sql4, Dbconn, 1
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
            end if 
        
            if cint(rs(0)) > 0 then
                sql4 = " DELETE FROM overtime WHERE  work_date = '"&work_date1&"' AND  mg_ce_id = '" & mg_ce_id & "'"
                'Response.write sql4 & "<br>"
                dbconn.execute(sql4)     
                if Err.number <> 0 then		
                    Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
                end if 
            end if
            rs.close()

            sql4 =  "INSERT INTO overtime                          "&chr(13)&_
                    "          ( work_date                         "&chr(13)&_
                    "          , end_date                          "&chr(13)&_
                    "          , mg_ce_id                          "&chr(13)&_
                    "          , user_name                         "&chr(13)&_
                    "          , user_grade                        "&chr(13)&_
                    "          , acpt_no                           "&chr(13)&_
                    "          , emp_company                       "&chr(13)&_
                    "          , bonbu                             "&chr(13)&_
                    "          , saupbu                            "&chr(13)&_
                    "          , team                              "&chr(13)&_
                    "          , org_name                          "&chr(13)&_
                    "          , reside_place                      "&chr(13)&_
                    "          , company                           "&chr(13)&_
                    "          , dept                              "&chr(13)&_
                    "          , work_item                         "&chr(13)&_
                    "          , from_time                         "&chr(13)&_
                    "          , to_time                           "&chr(13)&_
                    "          , delta_time                        "&chr(13)&_
                    "          , delta_minute                      "&chr(13)&_
                    "          , rest_time                         "&chr(13)&_
                    "          , rest_minute                       "&chr(13)&_
                    "          , work_gubun                        "&chr(13)&_
                    "          , cost_detail                       "&chr(13)&_
                    "          , person_amt                        "&chr(13)&_
                    "          , overtime_amt                      "&chr(13)&_
                    "          , alter_timeoff_date                "&chr(13)&_
                    "          , alter_timeoff_time                "&chr(13)&_
                    "          , alter_timeoff_minute_w            "&chr(13)&_
                    "          , alter_timeoff_minute_d            "&chr(13)&_
                    "          , work_memo                         "&chr(13)&_
                    "          , sign_no                           "&chr(13)&_
                    "          , you_yn                            "&chr(13)&_
                    "          , cancel_yn                         "&chr(13)&_
                    "          , end_yn                            "&chr(13)&_
                    "          , reg_id                            "&chr(13)&_
                    "          , reg_user                          "&chr(13)&_
                    "          , reg_date                          "&chr(13)&_
                    "          )                                   "&chr(13)&_
                    "   VALUES ( '" & work_date1 & "'              "&chr(13)&_
                    "          , '" & work_date2 & "'              "&chr(13)&_
                    "          , '" & mg_ce_id & "'                "&chr(13)&_
                    "          , '" & mmb_user_name & "'           "&chr(13)&_
                    "          , '" & mmb_user_grade & "'          "&chr(13)&_
                    "          , " & acpt_no & "                   "&chr(13)&_
                    "          , '" & mmb_emp_company & "'         "&chr(13)&_
                    "          , '" & mmb_bonbu & "'               "&chr(13)&_
                    "          , '" & mmb_saupbu & "'              "&chr(13)&_
                    "          , '" & mmb_team & "'                "&chr(13)&_
                    "          , '" & mmb_org_name & "'            "&chr(13)&_
                    "          , '" & mmb_reside_place & "'        "&chr(13)&_
                    "          , '" & mmb_reside_company & "'      "&chr(13)&_
                    "          , '" & dept & "'                    "&chr(13)&_
                    "          , '" & work_item & "'               "&chr(13)&_
                    "          , '" & from_time & "'               "&chr(13)&_
                    "          , '" & to_time & "'                 "&chr(13)&_
                    "          , '" & delta_time & "'              "&chr(13)&_
                    "          , '" & delta_minute & "'            "&chr(13)&_
                    "          , '" & rest_time & "'               "&chr(13)&_
                    "          , '" & rest_minute & "'             "&chr(13)&_
                    "          , '" & work_gubun & "'              "&chr(13)&_
                    "          , '" & cost_detail & "'             "&chr(13)&_
                    "          , 0                                 "&chr(13)&_
                    "          , "&overtime_amt&"                  "&chr(13)&_
                    "          , "&alter_timeoff_date_f&"          "&chr(13)&_
                    "          , '" & alter_timeoff_time & "'      "&chr(13)&_
                    "          , '" & alter_timeoff_minute_w & "'  "&chr(13)&_
                    "          , '" & alter_timeoff_minute_d & "'  "&chr(13)&_
                    "          , '" & work_memo & "'               "&chr(13)&_
                    "          , '" & sign_no & "'                 "&chr(13)&_
                    "          , '" & you_yn & "'                  "&chr(13)&_
                    "          , 'N'                               "&chr(13)&_
                    "          , 'N'                               "&chr(13)&_
                    "          , '" & user_id & "'                 "&chr(13)&_
                    "          , '" & user_name & "'               "&chr(13)&_
                    "          , now()                             "&chr(13)&_
                    "          )                                   "&chr(13)
            'Response.write "<pre>"&sql4 & "</pre><br>"
            dbconn.execute(sql4)
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
            end if 

        else ' ���� �������
            sql4 = "UPDATE overtime                                                    "&chr(13)&_
                   "   SET end_date               = '" & work_date2 & "'               "&chr(13)&_         
                   "     , mg_ce_id               = '" & mg_ce_id & "'                 "&chr(13)&_         
                   "     , user_name              = '" & mmb_user_name & "'            "&chr(13)&_         
                   "     , user_grade             = '" & mmb_user_grade & "'           "&chr(13)&_         
                   "     , acpt_no                = " & acpt_no & "                    "&chr(13)&_         
                   "     , emp_company            = '" & mmb_emp_company & "'          "&chr(13)&_         
                   "     , bonbu                  = '" & mmb_bonbu & "'                "&chr(13)&_         
                   "     , saupbu                 = '" & mmb_saupbu & "'               "&chr(13)&_         
                   "     , team                   = '" & mmb_team & "'                 "&chr(13)&_         
                   "     , org_name               = '" & mmb_org_name & "'             "&chr(13)&_         
                   "     , reside_place           = '" & mmb_reside_place & "'         "&chr(13)&_         
                   "     , company                = '" & mmb_reside_company & "'       "&chr(13)&_         
                   "     , dept                   = '" & dept & "'                     "&chr(13)&_         
                   "     , work_item              = '" & work_item & "'                "&chr(13)&_         
                   "     , from_time              = '" & from_time & "'                "&chr(13)&_         
                   "     , to_time                = '" & to_time & "'                  "&chr(13)&_         
                   "     , delta_time             = '" & delta_time & "'               "&chr(13)&_         
                   "     , delta_minute           = '" & delta_minute & "'             "&chr(13)&_         
                   "     , rest_time              = '" & rest_time & "'                "&chr(13)&_         
                   "     , rest_minute            = '" & rest_minute & "'              "&chr(13)&_         
                   "     , work_gubun             = '" & work_gubun & "'               "&chr(13)&_         
                   "     , cost_detail            = '" & cost_detail & "'              "&chr(13)&_         
                   "     , person_amt             = 0                                  "&chr(13)&_         
                   "     , overtime_amt           = "&overtime_amt&"                   "&chr(13)&_         
                   "     , alter_timeoff_date     = "&alter_timeoff_date_f&"           "&chr(13)&_         
                   "     , alter_timeoff_time     = '" & alter_timeoff_time & "'       "&chr(13)&_         
                   "     , alter_timeoff_minute_w = '" & alter_timeoff_minute_w & "'   "&chr(13)&_         
                   "     , alter_timeoff_minute_d = '" & alter_timeoff_minute_d & "'   "&chr(13)&_         
                   "     , work_memo              = '" & work_memo & "'                "&chr(13)&_         
                   "     , sign_no                = '" & sign_no & "'                  "&chr(13)&_         
                   "     , you_yn                 = '" & you_yn & "'                   "&chr(13)&_         
                   "     , cancel_yn              = '" & cancel_yn & "'                "&chr(13)&_         
                   "     , mod_date               = now()                              "&chr(13)&_ 
                   "     , mod_user               = '" & user_name & "'                "&chr(13)&_ 
                   "     , mod_id                 = '" & user_id & "'                  "&chr(13)&_ 
                   " WHERE work_date = '"&work_date1&"'                                "&chr(13)&_ 
                   "   AND mg_ce_id  = '" & mg_ce_id & "'                              "&chr(13)
            Response.write "<pre>"&sql4 & "</pre><br>"
            dbconn.execute(sql4)
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
            end if        
        end if 

        Sql4 = " SELECT count(*) FROM overtime                         "&chr(13)&_
               "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'  "&chr(13)&_
               "    AND mg_ce_id = '"& mg_ce_id &"'                    "&chr(13)&_
               "    AND alter_timeoff_minute_w <> 0                    "&chr(13)
        'Response.write "<pre>"&sql4 & "</pre><br>"       
        Set rs = Dbconn.Execute (sql4)
        if Err.number <> 0 then		
            Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
        end if 
    
        if cint(rs(0)) > 0 then
            ' 52�ð� �ʰ��� ���� ��ü�ް� �Ѻ��� �ش����� �����Ϳ� �ϰ������Ѵ�.
            sql3 = " UPDATE overtime                                            "&chr(13)&_
                   "    SET alter_timeoff_minute_w = "&alter_timeoff_minute_w&" "&chr(13)&_
                   "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'       "&chr(13)&_
                   "    AND mg_ce_id = '"& mg_ce_id &"'                         "&chr(13)
            'Response.write "<pre>"&sql3 & "</pre><br>"
            dbconn.execute(sql3)	   
            if Err.number <> 0 then		
                Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql3 & "]<br>"
            end if 
        end if
    end if

    if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "��� �� Error�� �߻��Ͽ����ϴ�....(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")"
	else    
		dbconn.CommitTrans
		end_msg = "��ϵǾ����ϴ�...."
    end if
    Response.write end_msg

	dbconn.Close()
	Set dbconn = Nothing
	
Response.write "<script language=javascript>"
Response.write "alert('"&end_msg&"');"
Response.write "parent.opener.location.reload();"
Response.write "self.close() ;"
Response.write "</script>"
Response.end

	
'  SELECT mg_ce_id     ���̵�
'       , allow_yn     ����
'       , allow_sayou  �̽��λ���
'       , concat(work_date,' ',from_time)         ������_�ú�                                     
'       , concat(end_date,' ',to_time)             ������_�ú�                                     
'       , concat(delta_time,'(',delta_minute,')')  �۾�_�ú�
'       , concat(rest_time,'(',rest_minute,')')    �ް�_�ú�
'       , concat(alter_timeoff_date,' ',alter_timeoff_time)   ��ü���Ͻ���
'       , concat( LPAD(Floor(alter_timeoff_minute_w/60),2,'0'), LPAD(Mod(alter_timeoff_minute_w,60),2,'0'),'(',alter_timeoff_minute_w,')') ��ü����_52�ʰ�_�ú�
'       , concat( LPAD(Floor(alter_timeoff_minute_d/60),2,'0'), LPAD(Mod(alter_timeoff_minute_d,60),2,'0'),'(',alter_timeoff_minute_d,')') ��ü����_228�ʰ�_�Ѻ�
'    FROM overtime A                                     
'   WHERE work_date BETWEEN '2018-08-22' AND '2018-08-28'         
'     AND  mg_ce_id = '101638'         
'ORDER BY work_date  
';
'
'SELECT concat( LPAD(Floor(sum_delta_minute/60),2,'0'), LPAD(Mod(sum_delta_minute,60),2,'0'),'(',sum_delta_minute,')') 52�ʰ�_�ú�
' FROM ( 
'        SELECT Sum(delta_minute) - (12*60) sum_delta_minute
'          FROM overtime A                                     
'         WHERE work_date BETWEEN '2018-08-22' AND '2018-08-28'         
'           AND mg_ce_id = '101638'    
'           AND allow_yn = 'Y'     
'      ) a
';

%>
