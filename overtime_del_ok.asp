<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

    work_date1 = request.form("work_date1")
    fDate	   = request.form("fDate")
    lDate	   = request.form("lDate")
	u_type     = request.form("u_type")
	work_item  = request.form("work_item")
	work_date  = request.form("work_date")
	mg_ce_id   = request.form("mg_ce_id")
	acpt_no    = request.form("acpt_no")

	set dbconn = server.CreateObject("adodb.connection")
	Set weeksRs = Server.CreateObject("ADODB.Recordset")
	
	dbconn.open dbconnect

	dbconn.BeginTrans
	
	if acpt_no = 0 then
		sql = "delete from overtime where work_date ='"&work_date1&"' and mg_ce_id='"&mg_ce_id&"'"
	else
		sql = "delete from overtime where acpt_no ="&int(acpt_no)
	end if
	'Response.write "<pre>"&sql&"</pre><br>"
	dbconn.execute(sql)

    if acpt_no <> 0 then
        sql = "Update as_acpt set overtime ='N' where acpt_no ="&int(acpt_no)
        'Response.write "<pre>"&sql&"</pre><br>"
        dbconn.execute(sql)
    end if

    if (fDate <> "") and (lDate <> "") then

        ' 해당 기준 주 내에서 작업시간 합을 구한다. (승인건만..)
        weeksSql = " SELECT ifnull(sum(delta_minute),0) total_minute          "&chr(13)&_
                "      , ifnull(Floor(sum(delta_minute)/60),0) floor_time  "&chr(13)&_
                "      , ifnull(Mod(sum(delta_minute),60),0)   mod_minute  "&chr(13)&_
                "   FROM overtime A                                        "&chr(13)&_ 
                "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'     "&chr(13)&_
                "    AND mg_ce_id = '"& user_id &"'                        "&chr(13)&_
                "    AND allow_yn = 'Y'                                    "&chr(13)
        'Response.write "<pre>"&weeksSql&"</pre><br>"
        weeksRs.Open weeksSql, Dbconn, 1

        if (weeksRs.eof or weeksRs.bof) then
            total_minute_aY = 0 
            work_time_aY    = 0
            work_minute_aY  = 0
        else
            total_minute_aY = Cint( weeksrs("total_minute") ) ' 총작업시간을 총분으로 .. (승인건만..)
            work_time_aY    = Cint( weeksRs("floor_time") )   ' 총작업시간을 시로 ..  (승인건만..)
            work_minute_aY  = Cint( weeksRs("mod_minute") )   ' 총작업시간을 시로 나눈몫인 분으로 ..  (승인건만..)
        end if

        weeksRs.close

        if  total_minute_aY > (12*60) then ' 12 시간 초과하면 초과분만 계산 출력
            alterTimeOff1   = total_minute_aY - 720
            alterTimeOff1_t = Fix(alterTimeOff1 / 60)
            alterTimeOff1_m = alterTimeOff1 mod 60
        else
            alterTimeOff1 = 0
            alterTimeOff1_t = 0
            alterTimeOff1_m = 0
        end if
        
        ' 52시간 초과에 대한 대체휴가 총분은 해당주의 데이터에 일괄적용한다.
        sql = " UPDATE overtime                                      "&chr(13)&_
            "    SET alter_timeoff_minute_w = '"&alterTimeOff1&"'  "&chr(13)&_
            "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"' "&chr(13)&_
            "    AND mg_ce_id = '"& mg_ce_id &"'                   "&chr(13)
        'Response.write "<pre>"&sql & "</pre><br>"
        dbconn.execute(sql)     

    end if 	

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
	end if

	Response.write"<script language=javascript>"
	Response.write"alert('"&end_msg&"');"
	Response.write"parent.opener.location.reload();"
	Response.write"self.close() ;"
	Response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing	
%>
