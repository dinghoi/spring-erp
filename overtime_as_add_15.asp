<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim weeksRs
Set Dbconn=Server.CreateObject("ADODB.Connection")
Set weeksRs = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

acpt_no 			= request.form("acpt_no")
company 			= request.form("company")
dept 				= request.form("dept")
work_date1 			= request.form("work_date1")
week1 				= request.form("week1")
work_date2 			= request.form("work_date2")
week2 				= request.form("week2")
work_item 			= request.form("work_item") ' �۾�����
dev_inst_cnt 		= request.form("dev_inst_cnt")
ran_cnt 			= request.form("ran_cnt")
work_man_cnt 		= request.form("work_man_cnt")
from_hh				= request.form("from_hh")
from_mm				= request.form("from_mm")
to_hh				= request.form("to_hh")
to_mm				= request.form("to_mm")
work_gubun			= request.form("work_gubun")
reg_sw				= request.form("reg_sw")
you_yn				= request.form("you_yn")
alter_timeoff_date  = request.form("alter_timeoff_date") ' ��ü�޹��ð�
alter_timeoff_hh    = request.form("alter_timeoff_hh") ' ��ü�޹��ð�_��
alter_timeoff_mm    = request.form("alter_timeoff_mm") ' ��ü�޹��ð�_��

Dim dateNow, week, mGap, fDate, lDate, work_time

dateNow = Date()        ' ��������
week    = Weekday(Date) ' �������

If  (week >= 4) Then
		mGap = (week - 4) * -1  
Else
		mGap = (6 - (3-week)) * -1  
End If

' ������ ~ ȭ����(������)
fDate = DateAdd("d", mGap, dateNow) 
lDate = DateAdd("d", mGap + 6, dateNow)

'Response.write Date &" :"& Weekday(Date) &"<br>"
'Response.write fDate &" ~ "& lDate &"<br>"

weeksSql = " SELECT A.work_date                                                                   "&chr(13)&_
           "      , A.end_date                                                                    "&chr(13)&_
           "      , A.mg_ce_id                                                                    "&chr(13)&_
           "      , A.to_time                                                                     "&chr(13)&_
           "      , A.from_time                                                                   "&chr(13)&_
           "      , left(A.to_time,2)    totime                                                   "&chr(13)&_
           "      , left(A.from_time,2)  fromtime                                                 "&chr(13)&_
           "      , right(A.to_time,2)   tominute                                                 "&chr(13)&_
           "      , right(A.from_time,2) fromminute                                               "&chr(13)&_
           "      , A.mg_ce_id                                                                    "&chr(13)&_
           "      , (SELECT visit_date FROM as_acpt WHERE acpt_no = A.acpt_no ) AS visit_date     "&chr(13)&_
           "   FROM overtime A                                                                    "&chr(13)&_ 
           "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'                                 "&chr(13)&_
           "    AND mg_ce_id = '"& user_id &"'                                                    "&chr(13)
'Response.write weeksSql&"<br>"
weeksRs.Open weeksSql, Dbconn, 1

do until weeksRs.eof  
	                                 
    work_date  = weeksRs("work_date")	      
    end_date   = weeksRs("end_date")                           
	mg_ce_id   = weeksRs("mg_ce_id")                                 
	
	to_time    = weeksRs("to_time")
	from_time  = weeksRs("from_time")
	
	totime     = Cint( weeksRs("totime") )
	fromtime   = Cint( weeksRs("fromtime") )
	tominute   = Cint( weeksRs("tominute") )
	fromminute = Cint( weeksRs("fromminute") )
	
	'Response.write  ": " & from_time & " ~ " & to_time & ":"
	' �۾��������ڰ� ���� ��(�������϶�)�� ���۽ð��� ����ð����θ� �۾�����ð��� ����� �� �۾�����ð��� �����Ѵ�.
	if  IsNull(end_date) = True then 
	
		if to_time >= from_time then ' �������϶� (���ͽð� < �����ð�)
			
				if tominute >= fromminute then ' �������϶� (�л���)
						deltaminute = tominute - fromminute
				else ' ���� ������ �� ũ�� �ÿ��� 60�� �����´�. (�л���)
						deltaminute = (tominute+60)	- fromminute 
						totime =  totime - 1					
				end if		
				deltatime =  totime - fromtime '  (�û���)
					
				end_date = Cdate(work_date) ' ������..
				'Response.write deltatime&":"&deltaminute&"   "
						
		else ' ������ ������ (���ͽð� > �����ð�)
				
				deltatime =  (24 - fromtime)
				if 0 < fromminute then ' ���� ������ �ø� �����ϰ� 60���� �ܿ����� ����Ѵ�.
						deltatime = deltatime - 1
						deltaminute = 60 - fromminute
				else
						deltaminute = 0
				end if
				
				' �������� �ÿ� ���� ������ �úа� ���Ѵ�.
				deltatime   = deltatime + totime
				deltaminute = deltaminute + tominute
				
				if deltaminute >= 60 then ' ���� ���� 60�� �� �ʰ��ϸ� �ð��� �߰��ϰ� ���� 60 ���Ϸ� �����.
						deltatime   = deltatime + 1 
						deltaminute = deltaminute - 60
				end if
				
				end_date = DateAdd("d", 1, Cdate(work_date))  ' �������� ó��
			  'Response.write deltatime&":"&deltaminute&"<br>"		
		
		end if
		
		sqlupt = " UPDATE overtime                                                                       "&chr(13)&_
		         "    SET end_date     = '"&end_date&"'                                                  "&chr(13)&_
		         "      , delta_time   = concat( LPAD('"&deltatime&"',2,0), LPAD('"&deltaminute&"',2,0)) "&chr(13)&_
		         "      , delta_minute = "&deltatime&" * 60 +"&deltaminute&"                             "&chr(13)&_
	           "  WHERE work_date = '"&work_date&"'                                                    "&chr(13)&_
	           "    AND mg_ce_id  = '"&mg_ce_id&"'                                                     "&chr(13)         
		'Response.write "<pre>"&sqlupt&"</pre><br>"
		'Response.write sqlupt&"<br>"
		dbconn.execute(sqlupt)
				
	end if
	
	weeksRs.movenext()
loop

weeksRs.close

' �ش� ���� �� ������ �۾��ð� ���� ���Ѵ�. (����,�̽��� ������� �Ѵ�)
weeksSql = " SELECT ifnull(sum(delta_minute),0)   delta_minute     "&chr(13)&_
           "      , ifnull(sum(rest_minute),0)    rest_minute      "&chr(13)&_
           "   FROM overtime A                                     "&chr(13)&_ 
           "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'  "&chr(13)&_
           "    AND mg_ce_id = '"& user_id &"'                     "&chr(13)
'Response.write weeksSql&"<br>"
weeksRs.Open weeksSql, Dbconn, 1


if (weeksRs.eof or weeksRs.bof) then
  	delta_minute = 0 
  	rest_minute  = 0 
  	work_time    = 0
  	work_minute  = 0
else
  	delta_minute = Cint( weeksrs("delta_minute") ) ' �Ѱ���ð��� �Ѻ����� .. (����,�̽��� ������� �Ѵ�)
  	rest_minute  = Cint( weeksrs("rest_minute") )  ' ���ްԽð��� �Ѻ����� .. (����,�̽��� ������� �Ѵ�)
  	if (delta_minute > rest_minute) then
      delta_minute = delta_minute - rest_minute
    else
      delta_minute = 0
    end if
    work_time   = Fix(delta_minute / 60) ' ���۾��ð��� �÷� ..  (����,�̽��� ������� �Ѵ�)
    work_minute = delta_minute mod 60    ' ���۾��ð��� �÷� �������� ������ ..  (����,�̽��� ������� �Ѵ�)
end if

weeksRs.close

title_line = "["&user_name &"] ���� ���� ��Ư�� ��� (�� ����["&fDate&" ~ "&lDate&"] �� ��Ư�� �ð�: "&work_time&"�ð� "&work_minute&"�� [�� "& FormatNumber(delta_minute-rest_minute,0) &" ��] )"


' �ش� ���� �� ������ �۾��ð� ���� ���Ѵ�. (���ΰǸ�..)
weeksSql = " SELECT ifnull(sum(delta_minute),0)   delta_minute     "&chr(13)&_
           "      , ifnull(sum(rest_minute),0)    rest_minute      "&chr(13)&_
           "   FROM overtime A                                     "&chr(13)&_ 
           "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'  "&chr(13)&_
           "    AND mg_ce_id = '"& user_id &"'                     "&chr(13)&_
           "    AND allow_yn = 'Y'                                 "&chr(13)
'Response.write weeksSql&"<br>"
weeksRs.Open weeksSql, Dbconn, 1

if (weeksRs.eof or weeksRs.bof) then
  	delta_minute_aY = 0 
  	rest_minute_aY  = 0 
  	work_time_aY    = 0
  	work_minute_aY  = 0
else
  	delta_minute_aY = Cint( weeksrs("delta_minute") ) ' �Ѱ���ð��� �Ѻ����� .. (���ΰǸ�..)
  	rest_minute_aY  = Cint( weeksrs("rest_minute") )  ' ���ްԽð��� �Ѻ����� .. (���ΰǸ�..)
    if (delta_minute_aY > rest_minute_aY) then
      delta_minute_aY = delta_minute_aY - rest_minute_aY
    else
      delta_minute_aY = 0
    end if
    work_time_aY   = Fix(delta_minute_aY / 60) ' ���۾��ð��� �÷� ..  (���ΰǸ�..)
    work_minute_aY = delta_minute_aY mod 60    ' ���۾��ð��� �÷� �������� ������ ..  (���ΰǸ�..)
end if

weeksRs.close

'Response.write delta_minute_aY&"<br>"
if  delta_minute_aY > (12*60) then ' 12 �ð�(720��) �ʰ��ϸ� �ʰ��и� ��� ���
    alterTimeOff1   = delta_minute_aY - (12*60)
    alterTimeOff1_t = Fix(alterTimeOff1 / 60)
    alterTimeOff1_m = alterTimeOff1 mod 60
else
    alterTimeOff1   = 0
    alterTimeOff1_t = 0
    alterTimeOff1_m = 0
end if

sumAlterTimeOff2 = 0
sumAlterTimeOff3 = 0

' �ش� ���� �� ������ �۾��ð� ���� ���Ѵ�.  (���ΰǸ�..)
weeksSql = "     SELECT A.work_date                                          "&chr(13)&_
           "          , A.end_date                                           "&chr(13)&_
           "          , A.from_time                                          "&chr(13)&_
           "          , A.to_time                                            "&chr(13)&_
           "          , left( A.from_time,2 )  fromtime                      "&chr(13)&_
           "          , right( A.from_time,2 ) fromminute                    "&chr(13)&_
           "          , left( A.to_time,2 )    totime                        "&chr(13)&_
           "          , right( A.to_time,2 )   tominute                      "&chr(13)&_
           "          , DAYOFWEEK( A.work_date ) work_week                   "&chr(13)&_
           "          , DAYOFWEEK( A.end_date )  end_week                    "&chr(13)&_
           "          , ifnull(A.delta_minute,0)           delta_minute      "&chr(13)&_
           "          , ifnull(Floor(A.delta_minute/60),0) floor_time        "&chr(13)&_
           "          , ifnull(Mod(A.delta_minute,60),0)   mod_minute        "&chr(13)&_
           "          , ifnull(A.rest_minute,0)            rest_minute       "&chr(13)&_
           "          , H1.holiday_memo holiday1                             "&chr(13)&_ 
           "          , H2.holiday_memo holiday2                             "&chr(13)&_ 
           "       FROM overtime A                                           "&chr(13)&_ 
           "  LEFT JOIN holiday H1                                           "&chr(13)&_ 
           "         ON A.work_date = H1.holiday                             "&chr(13)&_ 
           "  LEFT JOIN holiday H2                                           "&chr(13)&_ 
           "         ON A.work_date = H2.holiday                             "&chr(13)&_ 
           "      WHERE A.work_date BETWEEN '"&fDate&"' AND '"&lDate&"'      "&chr(13)&_
           "        AND A.mg_ce_id = '"& user_id &"'                         "&chr(13)&_
           "        AND A.allow_yn = 'Y'                                     "&chr(13)
'Response.write "<pre>"&weeksSql&"</pre><br>"
weeksRs.Open weeksSql, Dbconn, 1

do until weeksRs.eof or weeksRs.bof  
    
    work_date = weeksRs("work_date")
    end_date  = weeksRs("end_date")  
  
    work_week = Cint(weeksRs("work_week"))
    end_week  = Cint(weeksRs("end_week"))
  
    to_time    = weeksRs("to_time")
    from_time  = weeksRs("from_time")
  
    fromtime   = Cint( weeksRs("fromtime") )
    fromminute = Cint( weeksRs("fromminute") )
    totime     = Cint( weeksRs("totime") )
    tominute   = Cint( weeksRs("tominute") )  
    
    holiday1   = weeksRs("holiday1")
    holiday2   = weeksRs("holiday2")
  
    if (work_week = 1) or (work_week = 7) or (holiday1 <> "") then ' ������ (�Ͽ���, �����, holiday���̹�����)
        work_date_type = "����"
    else
        work_date_type = "����"
    end if
    if (end_week = 1) or (end_week = 7) or (holiday2 <> "") then ' ������ (�Ͽ���, �����, holiday���̹�����)
        end_date_type = "����"
    else
        end_date_type = "����"
    end if
  
    'Response.write "("&work_date_type&"~"&end_date_type&") , "    
    'Response.write "("&work_date&" "&weeksRs("fromtime")&":"&weeksRs("fromminute")&"), "
    'Response.write "("&end_date&" "&weeksRs("totime")&":"&weeksRs("tominute")&") "
    'Response.write "("&weeksRs("floor_time")&":"&weeksRs("mod_minute")&" "&weeksRs("delta_minute")&") <br>"

    delta_minute = Cint( weeksRs("delta_minute") ) ' �Ѱ���ð� 
    rest_minute  = Cint( weeksRs("rest_minute") )  ' ���ްԽð�
  
    alterTimeOff2 = 0 ' ��ü�޹��ð�(����) (�д���)
    alterTimeOff3 = 0 ' ��ü�޹��ð�(����) (�д���)
    
    if  work_date_type = "����" and end_date_type = "����" and work_date = end_date then ' ���� ~ ���� (������) �̸�
        ' �����
        if totime >= 22 then ' ���ð��� 22�ø� ������
            alterTimeOff2 = ((totime - 22) * 60)  + tominute 
            
            if fromtime >= 22 then ' ���۽ð��� 22�ø� ������
                alterTimeOff2 = alterTimeOff2 - (((fromtime - 22) * 60) + fromminute) 
            end if
        end if               
    end if
    
    if  work_date_type = "����" and end_date_type = "����" and work_date < end_date then ' ���� ~ ���� (������) �̸�
        ' �����
        if fromtime < 22 then      ' ���۽ð��� 22�� �����̸�
          alterTimeOff2 = 3 * 60
        else                       ' ���۽ð��� 22�ø� ������
          alterTimeOff2 = ((24 - fromtime )*60) - fromminute  
        end if
        ' �����
        alterTimeOff2 = alterTimeOff2 + (totime *60) + tominute  
    end if
    
    if  work_date_type = "����" and end_date_type = "����" then  ' ���� ~ ���� (������) �̸�
        ' �����
        if fromtime < 22 then      ' ���۽ð��� 22�� �����̸�
          alterTimeOff2 = 3 * 60
        else                       ' ���۽ð��� 22�ø� ������
          alterTimeOff2 = ((24 - fromtime )*60) - fromminute  
        end if
        ' �޴���
        if totime >= 8 then        ' ���ð��� 8�ð��� ������
          alterTimeOff3 = alterTimeOff3 + ((totime-8) *60) + tominute  
        end if
    end if
    
    if  work_date_type = "����" and end_date_type = "����" then  ' ���� ~ ���� (������) �̸�
        ' �޴���
        if  (24*60 - (fromtime*60+fromminute) )  >= (8*60) then
            alterTimeOff3  =  (24*60 - (fromtime*60+fromminute) ) - (8*60)
        end if
        ' �����
        if totime >= 22 then      ' ���ð��� 22�ð��� ������
            alterTimeOff2  =  alterTimeOff2 + ((totime*60)+tominute) - 22*60
        end if
    end if
    
    if  work_date_type = "����" and end_date_type = "����" and work_date = end_date then ' ���� ~ ���� (������) �̸�
        ' �޴���
        if  ((totime*60)+tominute) - ((fromtime*60)+fromminute) >= (8*60) then
            alterTimeOff3  =  ((totime*60)+tominute) - ((fromtime*60)+fromminute) - (8*60)
        end if
    end if
    
    if  work_date_type = "����" and end_date_type = "����" and work_date < end_date then ' ���� ~ ���� (������) �̸�
        ' �޴���
        if  (24*60 - (fromtime*60+fromminute) + (totime*60+tominute)) >= (8*60) then
            alterTimeOff3  =  (24*60 - (fromtime*60+fromminute) + (totime*60+tominute)) - (8*60)
        end if
    end if
  
    if  alterTimeOff2 > rest_minute then 
        alterTimeOff2 = alterTimeOff2 - rest_minute
    else               			             
        alterTimeOff2 = 0
    end if
    if  alterTimeOff3 > rest_minute then 
        alterTimeOff3 = alterTimeOff3 - rest_minute
    else               			             
        alterTimeOff3 = 0
    end if
    
    sumAlterTimeOff2 = sumAlterTimeOff2 + alterTimeOff2
    sumAlterTimeOff3 = sumAlterTimeOff3 + alterTimeOff3    
  
  weeksRs.movenext()
loop
weeksRs.close  


' ���� 22���ʰ� (����) ....
alterTimeOff2 = sumAlterTimeOff2
alterTimeOff2_t = Fix(sumAlterTimeOff2 / 60)
alterTimeOff2_m = sumAlterTimeOff2 mod 60

' ���� 8�ð��ʰ� (����) ....
alterTimeOff3 = sumAlterTimeOff3
alterTimeOff3_t = Fix(sumAlterTimeOff3 / 60)
alterTimeOff3_m = sumAlterTimeOff3 mod 60
    

' �ش� ���� �� ������ �۾��ð��� ���Ѵ�.  (�̽��� ��)
weeksSql = " SELECT ifnull(sum(delta_minute),0)   delta_minute     "&chr(13)&_
           "      , ifnull(sum(rest_minute),0)    rest_minute      "&chr(13)&_
           "   FROM overtime A                                     "&chr(13)&_ 
           "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'  "&chr(13)&_
           "    AND mg_ce_id = '"& user_id &"'                     "&chr(13)&_
           "    AND allow_yn <> 'Y'                                "&chr(13)
'Response.write weeksSql&"<br>"
weeksRs.Open weeksSql, Dbconn, 1


if (weeksRs.eof or weeksRs.bof) then
  	delta_minute_aN = 0 
  	rest_minute_aN  = 0 
  	work_time_aN    = 0
  	work_minute_aN  = 0
else
  	delta_minute_aN = Cint( weeksrs("delta_minute") ) ' �Ѱ���ð��� �Ѻ����� .. (�̽��� ��)
  	rest_minute_aN  = Cint( weeksrs("rest_minute") )  ' ���ްԽð��� �Ѻ����� .. (�̽��� ��)
    if (delta_minute_aN > rest_minute_aN) then
      delta_minute_aN = delta_minute_aN - rest_minute_aN
    else
      delta_minute_aN = 0
    end if
    work_time_aN   = Fix(delta_minute_aN / 60) ' ���۾��ð��� �÷� ..  (�̽��� ��)
    work_minute_aN = delta_minute_aN mod 60    ' ���۾��ð��� �÷� �������� ������ ..  (�¹��� ��)
end if

weeksRs.close

'  SELECT mg_ce_id     ���̵�
'       , allow_yn     ����
'       , allow_sayou  �̽��λ���
'       , concat(work_date,' ',from_time)         ������_�ú�                                     
'       , concat(end_date,' ',to_time)             ������_�ú�                                     
'       , concat(delta_time,'(',delta_minute,')')  �۾�_�ú�
'       , concat(rest_time,'(',rest_minute,')')    �ް�_�ú�
'       , concat(alter_timeoff_date,' ',alter_timeoff_time)   ��ü�޹��ð�����
'       , concat( LPAD(Floor(alter_timeoff_minute_w/60),2,'0'), LPAD(Mod(alter_timeoff_minute_w,60),2,'0'),'(',alter_timeoff_minute_w,')') ��ü�޹��ð�_52�ʰ�_�ú�
'       , concat( LPAD(Floor(alter_timeoff_minute_d/60),2,'0'), LPAD(Mod(alter_timeoff_minute_d,60),2,'0'),'(',alter_timeoff_minute_d,')') ��ü�޹��ð�_228�ʰ�_�Ѻ�
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<script src="/java/jquery-1.9.1.js"></script>
		<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
        <script type="text/javascript" src="/java/js_form.js"></script>
        
		<script type="text/javascript">

            // '�߱��׸�' ���ý� overtime_code_search.asp���� ȣ���ϴ� �Լ� ������ ����..
            function fn_SelectedOvertimeCode( work_gubun )
            {
                //alert(work_gubun);
            }

			$( document ).ready(function() {				
			    <%
			    if acpt_no = "" then
				    %>alert("������ �ְ� 12�ð� ������ ��Ư���� �����ϸ�, ���ε��� �ʴ� �߱��� �������� �ʽ��ϴ�.");<%
			  	end if
			    %>
			    
			    <% if from_hh <> "" then %>$("#from_hh").val(<%=from_hh%>);<% end if %>
			    <% if from_mm <> "" then %>$("#from_mm").val(<%=from_mm%>);<% end if %>
			    <% if to_hh <> "" then %>$("#to_hh").val(<%=to_hh%>);<% end if %>
			    <% if to_mm <> "" then %>$("#to_mm").val(<%=to_mm%>);<% end if %>
			    
			    getDeltaTime();
			});

            // Ư������ ������ ���Ѵ�.
			function whatWeek( day )
			{
				var week = ['��', '��', 'ȭ', '��', '��', '��', '��'];
				var dayOfWeek = week[new Date( day ).getDay()];

				return dayOfWeek;
			}
			
			function getDeltaTime()
			{
                var datetime1 = document.frm.work_date1.value;
                var datetime2 = document.frm.work_date2.value;
                
                // ó���� '��','��' ���Ͽ��� �����̰�.....
                var dtype = ['��', '��', '��', '��', '��', '��', '��'];
                
                var datetime1type = dtype[new Date(datetime1).getDay()];
                var datetime2type = dtype[new Date(datetime2).getDay()];
                
                // holiday ���̺��� �������ϴ� ajax�� ȣ�� �� ��...			  
                var params = { "work_date1" : datetime1
                             , "work_date2" : datetime2
                             };
                $.ajax({
    					 url: "ajax_get_checkedHoliday_2date.asp"
    					,async: false
    					,type: 'post'
    					,data: params
    					,dataType: "json"
    					,contentType: "application/x-www-form-urlencoded; charset=euc-kr"
    					,beforeSend: function(jqXHR){
    							jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
    					}
    					,success: function(data)
    					{    					  
    						var result        = data.result;
    						var holiday_memo1 = data.holiday_memo1;
    						var holiday_memo2 = data.holiday_memo2;
    						
    						if( result=="succ")
    						{
    						    // holiday ���̺� �ش��ϴ� ��¥�� ������ �����̴�.
    						    $("#holiday1").val(holiday_memo1);
                                $("#holiday2").val(holiday_memo2);        
                                
                                if  (holiday_memo1!="") datetime1type = '��' ;
                                if  (holiday_memo2!="") datetime2type = '��' ;        				      
    						}
    						else if( result=="invalid" ){
    							alert("�Է��Ͻ� ������ ��Ȯ���� �ʽ��ϴ�.");
    						}
    						else if(result=="fail"){
    							alert("���� �����߽��ϴ�.");
    						}
    					}
    					,error: function(jqXHR, status, errorThrown){
    						alert("������ �߻��Ͽ����ϴ�.\n�����ڵ� : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
    					}
				});
			  
			  
				var week  = ['��', '��', 'ȭ', '��', '��', '��', '��'];
				
				var dayWeek1  = week[new Date(datetime1).getDay()];
                var year1     = eval(datetime1.substring(0, 4));
                var month1    = eval(datetime1.substring(5, 7));
                var second1   = eval(datetime1.substring(8, 10));
                var hh1       = eval(document.frm.from_hh.value); 
                var mm1       = eval(document.frm.from_mm.value);				
                
                var year2     = eval(datetime2.substring(0, 4));
                var month2    = eval(datetime2.substring(5, 7));
                var second2   = eval(datetime2.substring(8, 10));
                var hh2       = eval(document.frm.to_hh.value); 
                var mm2       = eval(document.frm.to_mm.value);

				var startDate = new Date(year1,month1-1,second1,hh1,mm1,00); //console.log(startDate.toLocaleString());
 				var endDate   = new Date(year2,month2-1,second2,hh2,mm2,00); //console.log(endDate.toLocaleString()); 				

 				var delta_minute = (endDate.getTime() - startDate.getTime()) / 60000; 
 				var rest_minute = 0; // �����ΰ�� 4�ð���  30�� ������ ���� �޽Ľð� 
 				
 			    if (isNaN(delta_minute))
 				{
                    $("#spanDeltaTime").text(" - �ð� - �� (���)");  
                    $("#spanRestTime").text(" 00�ð� 00�� (�ް�)");  
                    document.frm1.delta_time.value = "";  
                    document.frm1.delta_minute.value = "0";  
                    document.frm1.rest_time.value =  "";  
                    document.frm1.rest_minute.value =  "0";  
                        
                    $("#AlterTimeoffTo").text('');
                        
                    $("#alter_timeoff_date").val('');
                    $("#alter_timeoff_hh").val(0);
                    $("#alter_timeoff_mm").val(0);
                    
                    $("#alter_timeoff_date").prop('disabled', true);
                    $("#alter_timeoff_hh").prop('disabled', true);
                    $("#alter_timeoff_mm").prop('disabled', true);     				  

                    return 0 ;
 				}
 				else
 				{    
                    // �ްԽð� ����
                    if  (delta_minute > ( 4*60)) rest_minute += 30 ;
                    if  (delta_minute > ( 8*60)) rest_minute += 30 ;
                    if  (delta_minute > (12*60)) rest_minute += 30 ;
                    if  (delta_minute > (16*60)) rest_minute += 30 ;
                    if  (delta_minute > (20*60)) rest_minute += 30 ;
                    
                    document.frm1.rest_time.value = pad(parseInt(rest_minute / 60),2) + pad(rest_minute % 60,2);
                    document.frm1.rest_minute.value = rest_minute;  
                    
                    var delta_time =  pad(parseInt(delta_minute / 60),2) + pad(delta_minute % 60,2);
                    $("#spanDeltaTime").text(" "+pad(parseInt(delta_minute / 60),2)+"�ð� " + pad(delta_minute % 60,2)+"�� (���)");
                    
                    if (rest_minute ==0)
                    {
                        $("#spanRestTime").text(" 00�ð� 00�� (�ް�)");  
                    }
                    else
                    {
                        $("#spanRestTime").text(" "+pad(parseInt(rest_minute / 60),2)+"�ð� " + pad(rest_minute % 60,2)+"�� (�ް�)");
                    }
                    
                    
                    document.frm1.delta_time.value = delta_time;  
                    document.frm1.delta_minute.value = delta_minute;  
                    
                    // ���� �۾��ð��� �̹��� �۾��ð��� ���� 12�ð��� �ʰ��ϸ� �� �ʰ��и�.. (��ü�޹��ð� �� 52�ð��ʰ�)
                    // alllow_yn �� ����Ʈ�� X �̹Ƿ� delta_minute ���� ��꿡 ������� �ʴ´�.
                    /*
                    if  ((<%=delta_minute_aY%> + delta_minute) >= (12*60))
                    alterTimeOff1 = (<%=delta_minute_aY%> + delta_minute) - (12*60) ;
                    */   		
                    //������ ������� �⺻���� 7�� ��� �����߻� -> alterTimeOff1, 2, 3 ���� ���� �߰�[20201030_����ȣ]		
                    var alterTimeOff1;
                    var alterTimeOff2;
                    var alterTimeOff3;                    

                    if  (<%=delta_minute_aY%> >= (12*60)){
                        alterTimeOff1 = <%=delta_minute_aY%> - (12*60) ;
                    }else{
                        alterTimeOff1 = 0; 
                    }  				  
                    
                    alterTimeOff2 = 0; // (��ü�޹��ð� ���� 22���ʰ�)
                    alterTimeOff3 = 0; // (��ü�޹��ð� ���� 8�ð��ʰ�)
                    
                    if ((datetime1type == '��') && (datetime2type == '��') && (datetime1 == datetime2)) // ���� ~ ���� (������) �̸�
                    {
                        // �����
                        if (hh2 >= 22) // ���ð��� 22�ø� ������
                        {
                            alterTimeOff2 = ((hh2 - 22) * 60)  + mm2 ;
                            
                            if (hh1 >= 22) // ���۽ð��� 22�ø� ������
                            {
                                alterTimeOff2 -= ((hh1 - 22) * 60) + mm1 ;
                            }
                        }
                    }
                    
                    if ((datetime1type == '��') && (datetime2type == '��') && (datetime1 < datetime2)) // ���� ~ ���� (������) �̸�
                    {
                        // �����
                        
                        if (hh1 < 22)       // ���۽ð��� 22�� �����̸�
                        {
                            alterTimeOff2 = 3 * 60 ;
                        }
                        else              // ���۽ð��� 22�ø� ������
                        {
                            alterTimeOff2 = ((24 - hh1)*60) - mm1 ;  
                        }
                        
                        // �����
                        alterTimeOff2 += (hh2 *60) + mm2 ;                 
                    }
                    
                    if ((datetime1type == '��') && (datetime2type == '��')) // ���� ~ ���� (������) �̸�
                    {
                        // �����
                        if (hh1 < 22)      // ���۽ð��� 22�� �����̸�
                        {
                            alterTimeOff2 = 3 * 60 ; // 3 = 22 + 23 + 24 
                        }
                        else                  // ���۽ð��� 22�ø� ������
                        {
                            alterTimeOff2 = ((24 - hh1 )*60) - mm1 ; 
                        }              
                        // �޴���
                        if (hh2 >= 8)         // ���ð��� 8�ð��� ������
                        {
                            alterTimeOff3 += ((hh2-8) *60) + mm2 ;  
                        }
                    }
                    
                    if ((datetime1type == '��') && (datetime2type == '��')) // ���� ~ ���� (������) �̸�
                    {
                        // �޴���
                        if  ((24*60 - (hh1*60+mm1)) >= (8*60) )
                        {
                            alterTimeOff3  =  (24*60 - (hh1*60+mm1)) - (8*60) ;
                        }
                        // �����
                        if (hh2 >= 22)       // ���ð��� 22�ð��� ������
                        {
                            alterTimeOff2  =  alterTimeOff2 + ((hh2*60)+mm2) - 22*60 ;
                        }              
                    }
                    
                    if ((datetime1type == '��') && (datetime2type == '��') && (datetime1 == datetime2)) // ���� ~ ���� (������) �̸�
                    {
                        // �޴���
                        if  ((((hh2*60)+mm2) - (hh1*60)+mm1) >= (8*60) )
                        {
                            alterTimeOff3  = (((hh2*60)+mm2) - ((hh1*60)+mm1)) - (8*60)
                        }
                    }
                    
                    if ((datetime1type == '��') && (datetime2type == '��') && (datetime1 < datetime2)) // ���� ~ ���� (������) �̸�
                    {
                        // �޴���
                        if  (((24*60) - (hh1*60+mm1) + (hh2*60+mm2)) >= (8*60))
                        {
                            alterTimeOff3  =  ((24*60) - (hh1*60+mm1) + (hh2*60+mm2)) - (8*60)
                        }
                    }
                    
                    if  (alterTimeOff2 > rest_minute) alterTimeOff2 = alterTimeOff2 - rest_minute;
                    else               			      alterTimeOff2 = 0;
                    if  (alterTimeOff3 > rest_minute) alterTimeOff3 = alterTimeOff3 - rest_minute;
                    else               			      alterTimeOff3 = 0;
                    
                    // �� 52�ð� �ʰ�
                    $("#alterTimeOff1").text( alterTimeOff1 );
                    $("#alterTimeOff1_t").text( pad(parseInt(alterTimeOff1 / 60),2) );
                    $("#alterTimeOff1_m").text( pad(alterTimeOff1 % 60,2) );

                    // ���� 22�� �ʰ�
                    $("#alterTimeOff2").text( alterTimeOff2 );
                    $("#alterTimeOff2_t").text( pad(parseInt(alterTimeOff2 / 60),2) );
                    $("#alterTimeOff2_m").text( pad(alterTimeOff2 % 60,2) );
                    
                    // ���� 8�ð� �ʰ�
                    $("#alterTimeOff3").text( alterTimeOff3 );
                    $("#alterTimeOff3_t").text( pad(parseInt(alterTimeOff3 / 60),2) );
                    $("#alterTimeOff3_m").text( pad(alterTimeOff3 % 60,2) );
                    
                    // �� 52�ð� �ʰ� + ���� 22�� �ʰ� + ���� 8�ð� �ʰ�
                    $("#alter_timeoff_minute_w").val(0);
                    $("#alter_timeoff_minute_d").val(0);
            
                    if  (alterTimeOff1 > 0) 
                    {
                        $("#alter_timeoff_minute_w").val(alterTimeOff1);
                    }
                    if  (alterTimeOff2 + alterTimeOff3 > 0) 
                    {
                        $("#alter_timeoff_minute_d").val(alterTimeOff2+alterTimeOff3);
                    }
                    
                    if  (alterTimeOff1+alterTimeOff2+alterTimeOff3>0) 
                    {     				  
                        getAlterTimeOff()
                        
                        $("#alter_timeoff_date").prop('disabled', false);
                        $("#alter_timeoff_hh").prop('disabled', false);
                        $("#alter_timeoff_mm").prop('disabled', false);
                    }
                    else
                    {
                        $("#AlterTimeoffTo").text('');
                        
                        $("#alter_timeoff_date").val('');
                        $("#alter_timeoff_hh").val(0);
                        $("#alter_timeoff_mm").val(0);
                        
                        $("#alter_timeoff_date").prop('disabled', true);
                        $("#alter_timeoff_hh").prop('disabled', true);
                        $("#alter_timeoff_mm").prop('disabled', true);     				  
                    }

                    return delta_minute;
                }
			}
			
			// ��ü�޹��ð����� 
			function getAlterTimeOff()
			{
			    var alterTimeoff = document.frm.alter_timeoff_date.value;
			    var year         = eval(alterTimeoff.substring(0, 4));
                var month        = eval(alterTimeoff.substring(5, 7));
                var second       = eval(alterTimeoff.substring(8, 10));
                var hh           = eval(document.frm.alter_timeoff_hh.value); 
                var mm           = eval(document.frm.alter_timeoff_mm.value);
                var alterTimeoffDate = new Date(year,month-1,second,hh,mm,00);
                
                if (isNaN(alterTimeoffDate)) 
                {
                    $("#AlterTimeoffTo").text('');
                }
                else
                {  
                    // 52�ð� �ʰ� �а� (8�ð� �ʰ�, 22�� �ʰ�)�� �߰��Ѵ�.
                    alterTimeoffDate.setMinutes(alterTimeoffDate.getMinutes()
                                            + eval($("#alter_timeoff_minute_w").val())
                                            + eval($("#alter_timeoff_minute_d").val()) );

                    
                    $("#AlterTimeoffTo").text( alterTimeoffDate.getFullYear() 
                                            + '-' 
                                            + pad(alterTimeoffDate.getMonth() + 1,2) 
                                            + '-' 
                                            + pad(alterTimeoffDate.getDate() ,2) 
                                            + ' ' 
                                            + pad(alterTimeoffDate.getHours(),2) 
                                            + ':' 
                                            + pad(alterTimeoffDate.getMinutes(),2)
                                            );
                }
			}

			$(function() {

  				$( "#work_date1" ).datepicker({  defaultDate : "<%=work_date1%>",
  				                                 dateFormat: "yy-mm-dd",
  				                                 //minDate : 0,
  				                                 onSelect: function(dateText) { $( "#week1" ).val( whatWeek(this.value) ); getDeltaTime(); } ,
  				                                 beforeShow: function(i) { if ($(i).attr('readonly')) { return false; } }
  				                              });

  				$( "#work_date2" ).datepicker({  defaultDate : "<%=work_date2%>",
  				                                 dateFormat: "yy-mm-dd",
  				                                 //minDate : 0,
  				                                 onSelect: function(dateText) { $( "#week2" ).val( whatWeek(this.value) ); getDeltaTime(); } ,
  				                                 beforeShow: function(i) { if ($(i).attr('readonly')) { return false; } }
  				                              });

  				$( "#alter_timeoff_date" ).datepicker({  defaultDate : "<%=alter_timeoff_date%>",
          				                                 dateFormat: "yy-mm-dd",
          				                                 minDate : 0,
          				                                 onSelect: function(dateText) { getAlterTimeOff(); } ,
          				                                 beforeShow: function(i) {  }
          				                              });

			});
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {                
				if (chkfrm()) {                    
					regview();
					document.frm.submit ();   

                    console.log('true');                                 
				}
				else
				{                  
					reghide();     
                    
                    console.log('false');             
				}                 
			}
			function frmcheck1 () {
				document.frm1.submit ();
			}
			
			function chkfrm() 
			{				
				if(document.frm.acpt_no.value =="" || document.frm.acpt_no.value =="0") {
					alert('�ش� ���񽺸� �����ϼž� �մϴ� !!!');
					frm.acpt_no.focus();
					return false;
				}
				
				if(document.frm.work_date1.value == "") {
					alert('�۾��������ڰ� ����� �Ǿ� ���� �ʽ��ϴ� !!!');
					frm.work_date1.focus();
					return false;
				}
				/*
				else
				{
						if(document.frm.work_date1.value < "<%=dateNow%>") {
							alert('�۾��������ڴ� ������ ���� �̿��� �մϴ�. !!!');
							frm.work_date1.focus();
							return false;
						}
				}
				*/
				
				if(document.frm.work_date2.value == "") {
					alert('�۾������ڰ� ����� �Ǿ� ���� �ʽ��ϴ� !!!');
					frm.work_date2.focus();
					return false;
				}
				else
				{
                    if(document.frm.work_date2.value < document.frm.work_date1.value) {
                        alert('�۾������ڴ� �۾��������ں��� ���ų� ���� �̿��� �մϴ�. !!!');
                        frm.work_date2.focus();
                        return false;
                    }
                    else
                    {
                        if (
                            (document.frm.work_date2.value == document.frm.work_date1.value) && 
                            (eval(document.frm.to_hh.value) < eval(document.frm.from_hh.value))
                            ) 
                        {
                            alert('�۾����ð��� �۾����۽ð����� ���ų� ���� �̿��� �մϴ�. !!!');
                            frm.to_hh.focus();
                            return false;
                        }
                        else
                        {
                            if (
                                (document.frm.work_date2.value == document.frm.work_date1.value) && 
                                (eval(document.frm.to_hh.value) == eval(document.frm.from_hh.value)) && 
                                (eval(document.frm.to_mm.value) < eval(document.frm.from_mm.value))
                                ) 
                            {
                                    alert('�۾������� �۾����ۺк��� ���ų� ���� �̿��� �մϴ�. !!!');
                                    frm.to_mm.focus();
                                    return false;
                            }      				    
                        }
                    }
				}
					
				if(document.frm.work_man_cnt.value == "0") {
					alert('�۾��� ����� �Ǿ� ���� �ʽ��ϴ� !!!');
					frm.work_man_cnt.focus();
					return false;
				}
				
				if(document.frm.work_gubun.value =="") {
					alert('�߱��׸��� �����ϼ���');
					frm.work_gubun.focus();
					return false;
				}

				if (getDeltaTime() <= 0) {
						alert('�۾��������ڽð��� �۾��������ڽð����� �ڿ� ���;� �մϴ�.');
						return false;
				}

				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.you_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("������ ������ üũ�ϼ���");
					return false;
				}
				/** ����ڰ� ��ü�޹��ð��� ������ ���� ���� ���
				if ((document.frm.alter_timeoff_date.disabled == false) && (document.frm.alter_timeoff_date.value==''))
				{
				    alert('��ü�޹��ð� �������� �����ϴ�.');
						frm.alter_timeoff_date.focus();
						return false;
				}
				*/		

				document.frm.reg_sw.value = "Y";
				return true;
			}
			function regview() {
				//document.getElementById('reg_view').style.display = '';                
                $('#reg_view').show();                
			}
			function reghide() {
				//document.getElementById('reg_view').style.display = 'none';
                $('#reg_view').hide();
			}
			
			function pad(n, width) {
			  n = n + '';
			  return n.length >= width ? n : new Array(width - n.length + 1).join('0') + n;
			}
			
    </script>
    
	</head>
	<body>
		<div id="container">
			<div class="gView">
			  
                <h3 class="tit"><%=title_line%></h3>                
                <br>

				<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
                    <colgroup>
                        <col width="20%" >
                        <col width="30%" >
                        <col width="20%" >
                        <col width="30%" >
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>���ε� ��Ư�ٽð�</th>
                            <td class="left">
                                <%=work_time_aY&"�ð� "&work_minute_aY&"�� [�� "& FormatNumber(delta_minute_aY,0) &" ��]" %>
                            </td>
                            <th>�̽��ε� ��Ư�ٽð�</th>
                            <td class="left">
                                <%=work_time_aN&"�ð� "&work_minute_aN&"�� [�� "& FormatNumber(delta_minute_aN,0) &" ��]" %>
                            </td>
                        </tr>
                        <tr>
                            <th>���޽ð�����(����)</th>
                            <td class="left" colspan="3">
                                <strong>�� 52�ð� �ʰ� :</strong><%=alterTimeOff1_t&"�ð� "&alterTimeOff1_m&"�� [�� "& FormatNumber(alterTimeOff1,0) &" ��]" %>
                                + <strong>���� 22�� �ʰ� :</strong><%=alterTimeOff2_t&"�ð� "&alterTimeOff2_m&"�� [�� "& FormatNumber(alterTimeOff2,0) &" ��]" %>
                                + <strong>���� 8�ð� �ʰ� :</strong><%=alterTimeOff3_t&"�ð� "&alterTimeOff3_m&"�� [�� "& FormatNumber(alterTimeOff3,0) &" ��]" %>
                            </td>
                        </tr>					  
					</tbody>						  
				</table>  				
  			    <br>
  				
				<form method="post" name="frm" action="overtime_as_add_15.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="38%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<tbody>						  
							<tr>
                                <th>����NO</th>
                                <td class="left">
                                    <input name="acpt_no" type="text" id="acpt_no" style="width:80px" readonly="true" value="<%=acpt_no%>">
                                    <a href="#" class="btnType03" onClick="pop_Window('as_search.asp?work_item=<%=work_item%>&srchTy=ce','as_search','scrollbars=yes,width=800,height=400')">������ȸ</a>
                                    <br>
                                    <strong style="color:#02880a; font-size:11px">* �ش� �����۾��Ͽ� 2���̻��� as���ΰ�� ��ǥas�Ǹ� �����Ͽ� �ٹ��۾��ð��� �ջ��� �Է����ּ���</strong>
                                </td>
                                <th>ȸ��/������</th>
                                <td class="left">
                                    <input name="company" type="text" id="company" style="width:100px" readonly="true" value="<%=company%>">
                                    <input name="dept" type="text" id="dept" style="width:150px" readonly="true" value="<%=dept%>">
                                </td>
						    </tr>
							<tr>
                                <th>�۾���</th>
                                <td class="left">
                                    <input name="work_date1" id="work_date1" type="text" <% if acpt_no = "" then %>readonly="true"<% end if %> style="width:80px" value="<%=work_date1%>">
                                    <input name="week1" id="week1" type="text" style="width:13px" readonly="true" value="<%=week1%>">
                                    <input name="holiday1" id="holiday1" type="text" style="width:55px" readonly="true" value="<%=holiday1%>">                    
                                    ~
                                    <input name="work_date2" id="work_date2" type="text" <% if acpt_no = "" then %>readonly="true"<% end if %> style="width:80px" value="<%=work_date2%>">
                                    <input name="week2" id="week2" type="text" style="width:13px" readonly="true" value="<%=week2%>">
                                    <input name="holiday2" id="holiday2" type="text" style="width:55px" readonly="true" value="<%=holiday2%>">                    
                                </td>
                                <th>�۾�����</th>
                                <td class="left">
                                    <input name="work_item" type="text" style="width:100px" readonly="true" value="<%=work_item%>">
                                    &nbsp;
                                    <strong>* �����󱸺�</strong>
                                    <label><input type="radio" name="you_yn" value="Y" <% if you_yn = "Y" then %>checked<% end if %> style="width:20px" id="Radio1">����</label>
                                    <label><input type="radio" name="you_yn" value="N" <% if you_yn = "N" then %>checked<% end if %> style="width:20px" id="Radio2">����</label>
                                </td>
						    </tr>
							<tr>
                                <th>�۾�����</th>
                                <td class="left">
                                    ��ġ : <input name="dev_inst_cnt" type="text" id="dev_inst_cnt" size="3" readonly="true" style="text-align:right" value="<%=dev_inst_cnt%>">&nbsp;
                                    ������ : <input name="ran_cnt" type="text" id="ran_cnt" size="3" readonly="true" style="text-align:right" value="<%=ran_cnt%>">&nbsp;
                                    �۾��ο� : <input name="work_man_cnt" type="text" id="work_man_cnt" size="2" readonly="true" style="text-align:right" value="<%=work_man_cnt%>">
							 	</td>
                                <th>�۾��ð�</th>
                                <td class="left">
                                    <table >
                                    <tr>
                                        <td style="border-width: 0px;">
                                            <!-- <%=from_hh%> --> <!-- <%=from_mm%>" -->
                                            <select name="from_hh" id="from_hh" style="width:40px" onchange="getDeltaTime()">
                                            <%
                                            for count = 0 to 23 step 1
                                                %><option value="<%=count%>"><%=count%></option><%
                                            next
                                            %>
                                            </select>��        									
                                            <select name="from_mm" id="from_mm" style="width:40px" onchange="getDeltaTime()">
                                                <option value="0">0</option>
                                                <option value="30">30</option>
                                            </select>�� 
                                            
                                            ~
                                                
                                            <!-- <%=to_hh%> --> <!-- <%=to_mm%>" -->
                                            <select name="to_hh" id="to_hh" style="width:40px" onchange="getDeltaTime()">
                                            <%
                                            for count = 0 to 23 step 1
                                                %><option value="<%=count%>"><%=count%></option><%
                                                next
                                            %>
                                            </select>��        									
                                            <select name="to_mm" id="to_mm" style="width:40px" onchange="getDeltaTime()">
                                                <option value="0">0</option>
                                                <option value="30">30</option>
                                            </select>��
                                        </td>
                                        <td style="border-width: 0px;">
                                            &nbsp;<span id="spanDeltaTime"></span><br>
                                            &nbsp;<span id="spanRestTime"></span>
                                        </td>
                                    </tr>    
                                    </table>  
                                    <strong style="color:#02880a; font-size:11px">* ��Ư�� ��Ͻ� �� 12�ð��� �ʰ��Ͽ� ����� �� �����ϴ�.</strong>
                                </td>
						    </tr>
							<tr>
                                <th>�߱��׸�</th>
                                <td class="left">
                                    <input name="work_gubun" type="text" value="<%=work_gubun%>" readonly="true" style="width:150px">
                                    <a href="#" onClick="pop_Window('overtime_code_search.asp?gubun=<%="AS"%>','overtime_code_pop','scrollbars=yes,width=700,height=400')" class="btnType03">��ȸ</a>
                                </td>
                                <th>��ü�޹��ð����</th>
                                <td class="left">
                                    <strong>�� 52�ð� �ʰ� :</strong> <span id="alterTimeOff1_t"></span>�ð� <span id="alterTimeOff1_m"></span>�� [ <span id="alterTimeOff1"></span> ��]<br>  
                                    <strong>���� 22�� �ʰ�:</strong> <span id="alterTimeOff2_t"></span>�ð� <span id="alterTimeOff2_m"></span>�� [ <span id="alterTimeOff2"></span> ��]<br>  
                                    <strong>���� 8�ð� �ʰ� :</strong> <span id="alterTimeOff3_t"></span>�ð� <span id="alterTimeOff3_m"></span>�� [ <span id="alterTimeOff3"></span> ��]
                                </td>
                            </tr>
                            <tr>
							    <th>��ü�޹��ð�<br>���ۼ���</th>
								<td colspan="3" class="left">
                                    <input name="alter_timeoff_date" id="alter_timeoff_date" type="text" disabled="true" style="width:80px" value="<%=alter_timeoff_date%>">��
                                    <!-- <%=alter_timeoff_hh%> -->
                                    <select name="alter_timeoff_hh" id="alter_timeoff_hh" disabled="true" style="width:40px" onchange="getAlterTimeOff()">
                                    <%
                                    for count = 0 to 23 step 1
                                        %><option value="<%=count%>"><%=count%></option><%
                                    next
                                    %>
                                    </select>��
                                    <!-- <%=alter_timeoff_mm%>" -->
                                    <select name="alter_timeoff_mm" id="alter_timeoff_mm" disabled="true" style="width:40px" onchange="getAlterTimeOff()">
                                        <option value="0">0</option>
                                        <option value="30">30</option>
                                    </select>�� 
                                    ~ 
                                    <span id="AlterTimeoffTo"></span>
                                    
                                    <input type="hidden" name="alter_timeoff_to" id="alter_timeoff_to" value="">
								</td>
						    </tr>
						</tbody>
					</table>					
					
					<h3 class="stit">
						<div style="width:100%; text-align:left; ">* �۾��� ����&nbsp;&nbsp;<a href="#" class="btnType03" onClick="javascript:frmcheck();">�۾���Ȯ��</a>
							<strong style="color:#02880a; font-size:11px"></strong>
						</div>
					</h3>
					<input type="hidden" name="holi_id" value="<%=holi_id%>" ID="Hidden1">
					<input type="hidden" name="sign_yn" value="<%=sign_yn%>" ID="Hidden1">				
					<input type="hidden" name="reg_sw" value="<%=reg_sw%>" ID="reg_sw">
				</form>
				
				<!-- ��� -->
				<form method="post" name="frm1" action="overtime_as_add_15_save.asp">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td width="49%" valign="top">
								<table cellpadding="0" cellspacing="0" class="tableList">
									<colgroup>
										<col width="10%" >
										<col width="20%" >
										<col width="*" >
										<col width="30%" >
									</colgroup>
                                    <thead>
                                        <tr>
                                            <th class="first" scope="col">����</th>
                                        <th scope="col">�۾���</th>
                                        <th scope="col">����</th>
                                        <th scope="col">����ó</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                        if isnull(work_man_cnt) or work_man_cnt = "" then
                                            record_cnt = 0
                                            work_man_cnt = 0
                                        else
                                            record_cnt = int((int(work_man_cnt) + 1)/2)
                                        end if

                                        sql = "select * from ce_work where work_id = '2' and acpt_no ="&int(acpt_no)&" limit 0," &record_cnt
                                        'Response.write sql & "<br>"
                                        Rs.Open Sql, Dbconn, 1

                                        i = 0
                                        do until rs.eof
                                            i = i + 1
                                            sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
                                            set rs_memb=dbconn.execute(sql)

                                            if	rs_memb.eof or rs_memb.bof then
                                                mg_ce = "ERROR"
                                            else
                                                mg_ce = rs_memb("user_name")
                                            end if
                                        rs_memb.close()
                                        %>
                                        <tr>
                                            <td class="first"><%=i%></td>
                                            <td><%=mg_ce%></td>
                                            <td><%=rs("org_name")%></td>
                                            <td><%=rs("reside_place")%>&nbsp;</td>
                                        </tr>
                                        <%
                    		            rs.movenext()
                                            loop
                                            rs.close()
                                        %>
                                    </tbody>
                                </table>
                            </td>
                            <td width="2%"></td>
                            <td width="49%" valign="top">
                                <table cellpadding="0" cellspacing="0" class="tableList">
                                <colgroup>
                                    <col width="10%" >
                                    <col width="20%" >
                                    <col width="*" >
                                    <col width="30%" >
                                </colgroup>
                                <thead>
                                    <tr>
                                        <th class="first" scope="col">����</th>
                                        <th scope="col">�۾���</th>
                                        <th scope="col">����</th>
                                        <th scope="col">����ó</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    sql = "select * from ce_work where work_id = '2' and acpt_no ="&int(acpt_no)&" limit "&record_cnt&"," &work_man_cnt
                                    'Response.write sql & "<br>"
                                    Rs.Open Sql, Dbconn, 1

                                    i = record_cnt
                                    do until rs.eof
                                        i = i + 1

                                        sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
                                        set rs_memb=dbconn.execute(sql)

                                        if	rs_memb.eof or rs_memb.bof then
                                            mg_ce = "ERROR"
                                        else
                                            mg_ce = rs_memb("user_name")
                                        end if
                                        rs_memb.close()
                                        %>
                                        <tr>
                                            <td class="first"><%=i%></td>
                                            <td><%=mg_ce%></td>
                                            <td><%=rs("team")%></td>
                                            <td><%=rs("reside_place")%>&nbsp;</td>
                                        </tr>
                                        <%
                                        rs.movenext()
                                    loop
                                    rs.close()
                                    %>
                                </tbody>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <br>
                    
					<div align=center id="reg_view">
						<%						  
						' ce_work�� acpt_no�� �ش��ϴ� �۾��ڰ� ������ �ִ��� üũ�Ѵ�.						  
						sql = "select count(*) cntWork from ce_work where work_id = '2' and acpt_no ="&int(acpt_no)
                        'Response.write sql & "<br>"
                        Rs.Open Sql, Dbconn, 1
                        
                        cntWork = 0
                        if not rs.eof then
                            cntWork = Cint(rs("cntWork"))
                        end if
                                    
                        if reg_sw = "Y" and work_man_cnt > 0 and cntWork > 0 then
                            'if (work_time < 12) then
                            %>
                            <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck1();"></span>
                            <%
                            'end if
                        end if
                        %>
                        <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
                    </div>
                    
					<input type="hidden" name="work_man_cnt" value="<%=work_man_cnt%>">
					<input type="hidden" name="reg_sw" value="<%=reg_sw%>">
					<input type="hidden" name="acpt_no" value="<%=acpt_no%>">
					<input type="hidden" name="company" value="<%=company%>">
					<input type="hidden" name="dept" value="<%=dept%>">
					<input type="hidden" name="work_item" value="<%=work_item%>">
					
					<input type="hidden" name="work_date1" value="<%=work_date1%>">
					<input type="hidden" name="work_date2" value="<%=work_date2%>">
					
					<input type="hidden" name="from_hh" value="<%=from_hh%>">
					<input type="hidden" name="from_mm" value="<%=from_mm%>">
					<input type="hidden" name="to_hh" value="<%=to_hh%>">
					<input type="hidden" name="to_mm" value="<%=to_mm%>">
					
					<input type="hidden" name="delta_time" value="<%=delta_time%>">
					<input type="hidden" name="delta_minute" value="<%=delta_minute%>">
					
					<input type="hidden" name="rest_time" value="<%=rest_time%>">					
  			        <input type="hidden" name="rest_minute" value="<%=rest_minute%>">
					
					<input type="hidden" name="work_gubun" value="<%=work_gubun%>">
					<input type="hidden" name="you_yn" value="<%=you_yn%>">
					
					<input type="hidden" name="alter_timeoff_date" value="<%=alter_timeoff_date%>">
					<input type="hidden" name="alter_timeoff_hh" value="<%=alter_timeoff_hh%>">
					<input type="hidden" name="alter_timeoff_mm" value="<%=alter_timeoff_mm%>">
					<input type="hidden" name="alter_timeoff_minute_w" id="alter_timeoff_minute_w">
					<input type="hidden" name="alter_timeoff_minute_d" id="alter_timeoff_minute_d">
					
					<input type="hidden" name="fDate" value="<%=fDate%>">
					<input type="hidden" name="lDate" value="<%=lDate%>">
					
				</form>
			</div>
		</div>
	</body>
</html>
