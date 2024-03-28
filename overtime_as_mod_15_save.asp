<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
on Error resume next

	'출처: http://start0.tistory.com/109 
	'------------------------------------------------
	' 사용예 : format(123, "00000") ==> 000123
	'------------------------------------------------
	function format(ByVal szString, ByVal Expression)
		if len(szString) < len(Expression) then
			format = left(expression, len(szString)) & szString
		else
			format = szString
		end if
	end function

  fDate                  = request.form("fDate")
  lDate                  = request.form("lDate")
  work_date1             = request.form("work_date1")
  work_date2             = request.form("work_date2")
  from_hh                = format(request.form("from_hh"),"00")
  from_mm                = format(request.form("from_mm"),"00") 
  from_time              = from_hh + from_mm
  to_hh                  = format(request.form("to_hh"),"00") 
  to_mm                  = format(request.form("to_mm"),"00") 
  to_time                = to_hh + to_mm
  work_gubun             = request.form("work_gubun")
  mg_ce_id               = request.form("mg_ce_id")
  work_memo              = request.form("work_memo")
  you_yn                 = request.form("you_yn") 
  cancel_yn              = request.form("cancel_yn") 
  delta_time             = request.form("delta_time")    
  delta_minute           = request.form("delta_minute")    
	rest_time              = request.form("rest_time")	
	rest_minute            = request.form("rest_minute")	  
  alter_timeoff_date     = request.form("alter_timeoff_date") 
  alter_timeoff_hh       = request.form("alter_timeoff_hh")
  alter_timeoff_mm       = request.form("alter_timeoff_mm") 
  alter_timeoff_time     = format(cstr(alter_timeoff_hh),"00") + format(cstr(alter_timeoff_mm),"00")
  alter_timeoff_minute_w = request.form("alter_timeoff_minute_w") 
  alter_timeoff_minute_d = request.form("alter_timeoff_minute_d") 

  set dbconn = server.CreateObject("adodb.connection")  
  Set Rs1    = Server.CreateObject("ADODB.Recordset")
  
  dbconn.open dbconnect

  dbconn.BeginTrans

  sql = "select * from overtime_code where work_gubun = '"&work_gubun&"'"
  'Response.write "<pre>"&sql & "</pre><br>"
  set rs_etc = dbconn.execute(sql)
  
  if not (rs_etc.bof or rs_etc.eof) then
    cost_detail  = rs_etc("cost_detail")
    overtime_amt = rs_etc("overtime_amt")
  end if
  
  rs_etc.close()
  
  if Len(alter_timeoff_date) = 0 then
    alter_timeoff_date_f = "NULL"
    alter_timeoff_time = "0000" ' 대체휴가시작일이 없으면 시간은 0000
  else
    alter_timeoff_date_f = "'"&alter_timeoff_date&"'"  
  end if

  ' 작업시작일자는 변경할 수 없습니다.    
  sql = "UPDATE overtime                                              "&chr(13)&_
        "   SET end_date               = '"&work_date2&"'             "&chr(13)&_
        "     , from_time              = '"&from_time&"'              "&chr(13)&_
        "     , to_time                = '"&to_time&"'                "&chr(13)&_
        "     , delta_time             = '"&delta_time&"'             "&chr(13)&_ 
        "     , delta_minute           = '"&delta_minute&"'           "&chr(13)&_ 
        "     , rest_time              = '"&rest_time&"'              "&chr(13)&_ 
        "     , rest_minute            = '"&rest_minute&"'            "&chr(13)&_ 
        "     , work_gubun             = '"&work_gubun&"'             "&chr(13)&_ 
        "     , cost_detail            = '"&cost_detail&"'            "&chr(13)&_ 
        "     , overtime_amt           = '"&overtime_amt&"'           "&chr(13)&_
        "     , work_memo              = '"&work_memo&"'              "&chr(13)&_
        "     , you_yn                 = '"&you_yn&"'                 "&chr(13)&_
        "     , alter_timeoff_date     = "&alter_timeoff_date_f&"     "&chr(13)&_
        "     , alter_timeoff_time     = '"&alter_timeoff_time&"'     "&chr(13)&_
        "     , alter_timeoff_minute_w = '"&alter_timeoff_minute_w&"' "&chr(13)&_
        "     , alter_timeoff_minute_d = '"&alter_timeoff_minute_d&"' "&chr(13)&_
        "     , cancel_yn              = '"&cancel_yn&"'              "&chr(13)&_
        "     , mod_id                 = '"&user_id&"'                "&chr(13)&_ 
        "     , mod_user               = '"&user_name&"'              "&chr(13)&_ 
        "     , mod_date               = now()                        "&chr(13)&_ 
        " WHERE work_date = '"& work_date1 &"'                        "&chr(13)&_ 
        "   AND mg_ce_id  = '"& mg_ce_id &"'                          "&chr(13)
  'Response.write "<pre>"&sql & "</pre><br>"
  dbconn.execute(sql)    
  
  ' 52시간 초과에 대한 대체휴가 총분은 해당주의 데이터에 일괄적용한다.
  sql = " UPDATE overtime                                              "&chr(13)&_
        "    SET alter_timeoff_minute_w = '"&alter_timeoff_minute_w&"' "&chr(13)&_
        "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'         "&chr(13)&_
        "    AND mg_ce_id = '"& mg_ce_id &"'                           "&chr(13)
  'Response.write "<pre>"&sql & "</pre><br>"
  dbconn.execute(sql)     
  
'  SELECT mg_ce_id     아이디
'       , allow_yn     승인
'       , allow_sayou  미승인사유
'       , concat(work_date,' ',from_time)         시작일_시분                                     
'       , concat(end_date,' ',to_time)             종료일_시분                                     
'       , concat(delta_time,'(',delta_minute,')')  작업_시분
'       , concat(rest_time,'(',rest_minute,')')    휴게_시분
'       , concat(alter_timeoff_date,' ',alter_timeoff_time)   대체휴일시작
'       , concat( LPAD(Floor(alter_timeoff_minute_w/60),2,'0'), LPAD(Mod(alter_timeoff_minute_w,60),2,'0'),'(',alter_timeoff_minute_w,')') 대체휴일_52초과_시분
'       , concat( LPAD(Floor(alter_timeoff_minute_d/60),2,'0'), LPAD(Mod(alter_timeoff_minute_d,60),2,'0'),'(',alter_timeoff_minute_d,')') 대체휴일_228초과_총분
'    FROM overtime A                                     
'   WHERE work_date BETWEEN '2018-08-22' AND '2018-08-28'         
'     AND  mg_ce_id = '101638'         
'ORDER BY work_date  
';
'
'SELECT concat( LPAD(Floor(sum_delta_minute/60),2,'0'), LPAD(Mod(sum_delta_minute,60),2,'0'),'(',sum_delta_minute,')') 52초과_시분
' FROM ( 
'        SELECT Sum(delta_minute) - (12*60) sum_delta_minute
'          FROM overtime A                                     
'         WHERE work_date BETWEEN '2018-08-22' AND '2018-08-28'         
'           AND mg_ce_id = '101638'    
'           AND allow_yn = 'Y'     
'      ) a
';

  
  if Err.number <> 0 then
    dbconn.RollbackTrans 
    end_msg = "처리중 Error가 발생하였습니다...."
  else    
    dbconn.CommitTrans
    end_msg = "처리 되었습니다...."
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
