<%@LANGUAGE="VBSCRIPT"%>
<%
Response.expires=-1
Response.ContentType = "application/json"
Response.Charset = "euc-kr"
%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
On Error Resume Next

Dim result : result = "fail"
Dim sql
Dim orgColumn, sqlWhere

  work_date	  = request("work_date")
  mg_ce_id	  = request("mg_ce_id")
  allow_yn  	= request("allow_yn")
  allow_sayou	= unescape( request("allow_sayou"))
   

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set weeksRs = Server.CreateObject("ADODB.Recordset")
	
	dbconn.open DbConnect

	sql = "UPDATE overtime                              "&chr(13)&_
	      "   SET allow_yn    = '" & allow_yn & "'      "&chr(13)&_
	      "     , allow_sayou = '" & allow_sayou & "'   "&chr(13)&_
	      " WHERE work_date = '" & work_date & "'       "&chr(13)&_ 
	      "   AND mg_ce_id  = '" & mg_ce_id & "'        "
  'Response.write sql
	Dbconn.execute sql
	
	
	dateNow = CDate(work_date) ' 일자변환
  week    = Weekday(dateNow) ' 요일  

  If  (week >= 4) Then
  		mGap = (week - 4) * -1  
  Else
  		mGap = (6 - (3-week)) * -1  
  End If

  ' 수요일 ~ 화요일(다음주)
  fDate = DateAdd("d", mGap, dateNow) 
  lDate = DateAdd("d", mGap + 6, dateNow) 
  
	' 해당 기준 주 내에서 작업시간 합을 구한다. (승인건만..)
  weeksSql = " SELECT ifnull(sum(delta_minute),0) total_minute          "&chr(13)&_
             "      , ifnull(Floor(sum(delta_minute)/60),0) floor_time  "&chr(13)&_
             "      , ifnull(Mod(sum(delta_minute),60),0)   mod_minute  "&chr(13)&_
             "   FROM overtime A                                        "&chr(13)&_ 
             "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'     "&chr(13)&_
             "    AND mg_ce_id = '"& mg_ce_id &"'                       "&chr(13)&_
             "    AND allow_yn = 'Y'                                    "&chr(13)
  'Response.write weeksSql&"<br>"
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
      alterTimeOff1   = 0
      alterTimeOff1_t = 0
      alterTimeOff1_m = 0
  end if
  
  ' 52시간 초과에 대한 대체휴가 총분은 해당주의 데이터에 일괄적용한다.
  sql = " UPDATE overtime                                              "&chr(13)&_
        "    SET alter_timeoff_minute_w = '"& alterTimeOff1 &"'        "&chr(13)&_
        "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'         "&chr(13)&_
        "    AND mg_ce_id = '"& mg_ce_id &"'                           "&chr(13)
  'Response.write "<pre>"&sql & "</pre><br>"
  dbconn.execute(sql)     
  
	

	result = "succ"

	Dbconn.close : Set Dbconn = Nothing

If Err.number<>0 Then
	result = "error"
End IF


If Trim(result&"")<>"" Then
	result = "{""result"":""" & result & """}"
End If

Response.write result
%>