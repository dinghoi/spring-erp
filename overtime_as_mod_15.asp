<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set weeksRs = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

	work_date  = request("work_date") ' ������ ��������
	mg_ce_id   = request("mg_ce_id")
	
  Dim dateNow, week, mGap, fDate, lDate, work_time
   
  dateNow = CDate(work_date) ' ���ں�ȯ
  week    = Weekday(dateNow) ' ����  

  If  (week >= 4) Then
  		mGap = (week - 4) * -1  
  Else
  		mGap = (6 - (3-week)) * -1  
  End If

  ' ������ ~ ȭ����(������)
  fDate = DateAdd("d", mGap, dateNow) 
  lDate = DateAdd("d", mGap + 6, dateNow)

  'Response.write dateNow &" :"& Weekday(dateNow) &"<br>"
  'Response.write fDate &" ~ "& lDate &"<br>"
  
  sql = "     SELECT A.work_date                                      "&chr(13)&_ 
        "          , A.mg_ce_id                                       "&chr(13)&_ 
        "          , A.user_name                                      "&chr(13)&_ 
        "          , A.user_grade                                     "&chr(13)&_ 
        "          , A.acpt_no                                        "&chr(13)&_ 
        "          , A.emp_company                                    "&chr(13)&_ 
        "          , A.bonbu                                          "&chr(13)&_ 
        "          , A.saupbu                                         "&chr(13)&_ 
        "          , A.team                                           "&chr(13)&_ 
        "          , A.org_name                                       "&chr(13)&_ 
        "          , A.reside_place                                   "&chr(13)&_ 
        "          , A.company                                        "&chr(13)&_ 
        "          , A.dept                                           "&chr(13)&_ 
        "          , A.work_item                                      "&chr(13)&_ 
        "          , A.end_date                                       "&chr(13)&_ 
        "          , A.from_time                                      "&chr(13)&_ 
        "          , A.to_time                                        "&chr(13)&_ 
        "          , left( A.from_time,2 )  fromtime                  "&chr(13)&_
        "          , right( A.from_time,2 ) fromminute                "&chr(13)&_
        "          , left( A.to_time,2 )    totime                    "&chr(13)&_
        "          , right( A.to_time,2 )   tominute                  "&chr(13)&_
        "          , ifnull(A.delta_time,0)   delta_time              "&chr(13)&_ 
        "          , ifnull(A.delta_minute,0) delta_minute            "&chr(13)&_ 
        "          , ifnull(A.rest_time,0)    rest_time               "&chr(13)&_ 
        "          , ifnull(A.rest_minute,0)  rest_minute             "&chr(13)&_ 
        "          , A.work_gubun                                     "&chr(13)&_ 
        "          , A.cost_detail                                    "&chr(13)&_ 
        "          , A.ask_amt                                        "&chr(13)&_ 
        "          , A.dev_inst_cnt                                   "&chr(13)&_ 
        "          , A.ran_cnt                                        "&chr(13)&_ 
        "          , A.work_man_cnt                                   "&chr(13)&_ 
        "          , A.alba_cnt                                       "&chr(13)&_ 
        "          , A.person_amt                                     "&chr(13)&_ 
        "          , A.overtime_amt                                   "&chr(13)&_ 
        "          , A.work_memo                                      "&chr(13)&_ 
        "          , A.sign_no                                        "&chr(13)&_ 
        "          , A.you_yn                                         "&chr(13)&_ 
        "          , A.cancel_yn                                      "&chr(13)&_ 
        "          , A.end_yn                                         "&chr(13)&_ 
        "          , A.cost_center                                    "&chr(13)&_ 
        "          , A.reg_id                                         "&chr(13)&_ 
        "          , A.reg_user                                       "&chr(13)&_ 
        "          , A.reg_date                                       "&chr(13)&_ 
        "          , A.mod_id                                         "&chr(13)&_ 
        "          , A.mod_user                                       "&chr(13)&_ 
        "          , A.mod_date                                       "&chr(13)&_ 
        "          , A.con_yn                                         "&chr(13)&_ 
        "          , A.mg_saupbu                                      "&chr(13)&_ 
        "          , A.alter_timeoff_date                             "&chr(13)&_ 
        "          , A.alter_timeoff_time                             "&chr(13)&_ 
        "          , A.alter_timeoff_minute_w                         "&chr(13)&_ 
        "          , A.alter_timeoff_minute_d                         "&chr(13)&_ 
        "          , A.allow_yn                                       "&chr(13)&_ 
        "          , A.allow_sayou                                    "&chr(13)&_ 
        "          , DAYOFWEEK( A.work_date ) work_week               "&chr(13)&_
        "          , DAYOFWEEK( A.end_date )  end_week                "&chr(13)&_
        "          , ifnull(A.delta_minute,0)           delta_minute  "&chr(13)&_
        "          , ifnull(Floor(A.delta_minute/60),0) floor_time    "&chr(13)&_
        "          , ifnull(Mod(A.delta_minute,60),0)   mod_minute    "&chr(13)&_
        "          , ifnull(A.rest_minute,0)            rest_minute   "&chr(13)&_
        "          , H1.holiday_memo holiday1                         "&chr(13)&_ 
        "          , H2.holiday_memo holiday2                         "&chr(13)&_ 
        "       FROM overtime A                                       "&chr(13)&_ 
        "  LEFT JOIN holiday H1                                       "&chr(13)&_ 
        "         ON a.work_date = H1.holiday                         "&chr(13)&_ 
        "  LEFT JOIN holiday H2                                       "&chr(13)&_ 
        "         ON a.work_date = H2.holiday                         "&chr(13)&_ 
        "      WHERE A.mg_ce_id  = '"& mg_ce_id &"'                   "&chr(13)&_
        "        AND A.work_date = '"& work_date &"'                  "&chr(13)
	'Response.write "<pre>"&sql&"</pre><br>"
	set rs = dbconn.execute(sql)
	
	if	rs.eof or rs.bof then    
    %>    
    <script type="text/javascript">
        alert('�ش��ϴ� ��Ư�� �ڷᰡ �����ϴ�.');
       //window.close();
    </script>
    <%
	else
  	sql="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
  	set rs_memb=dbconn.execute(sql)

  	if	rs_memb.eof or rs_memb.bof then
  		mg_ce = "ERROR"
  	else
  		mg_ce = rs_memb("user_name")
  	end if
  	rs_memb.close()						

  	if isnull(rs("acpt_no")) then
  		acpt_no = 0
  	else
  		acpt_no = rs("acpt_no")
  	end if
  	
  	
		strNowWeek = WeekDay(rs("work_date"))
		Select Case (strNowWeek)
		   Case 1	week1 = "��"
		   Case 2	week1 = "��"
		   Case 3	week1 = "ȭ"
		   Case 4	week1 = "��"
		   Case 5	week1 = "��"
		   Case 6	week1 = "��"
		   Case 7	week1 = "��"
		End Select		
  	
		strNowWeek = WeekDay(rs("end_date"))
		Select Case (strNowWeek)
		   Case 1	week2 = "��"
		   Case 2	week2 = "��"
		   Case 3	week2 = "ȭ"
		   Case 4	week2 = "��"
		   Case 5	week2 = "��"
		   Case 6	week2 = "��"
		   Case 7	week2 = "��"
		End Select		
  	
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
  'Response.write "<pre>"&weeksSql&"</pre><br>"
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
  			  'Response.write deltatime&":"&deltaminute&"   "
  		
  		end if
  		
  		sqlupt = " UPDATE overtime                                                                      "&chr(13)&_
  		        "    SET end_date     = '"&end_date&"'                                                  "&chr(13)&_
  		        "      , delta_time   = concat( LPAD('"&deltatime&"',2,0), LPAD('"&deltaminute&"',2,0)) "&chr(13)&_
  		        "      , delta_minute = "&deltatime&" * 60 +"&deltaminute&"                             "&chr(13)&_
  	          "  WHERE work_date = '"&work_date&"'                                                    "&chr(13)&_
  	          "    AND mg_ce_id  = '"&mg_ce_id&"'                                                     "&chr(13)         
  		'Response.write "<pre>"&sqlupt&"</pre><br>"
  		'Response.write  sqlupt&"<br>"	
  		dbconn.execute(sql)
  				
  	end if
  	weeksRs.movenext()
  loop

  weeksRs.close

  ' �ش� ���� �� ������ �۾��ð� ���� ���Ѵ�. (����,�̽��� ������� �Ѵ�) (�����ϰ��� ������� ����)
  weeksSql = " SELECT ifnull(sum(delta_minute),0)  delta_minute     "&chr(13)&_
             "      , ifnull(sum(rest_minute),0)   rest_minute      "&chr(13)&_
             "   FROM overtime A                                    "&chr(13)&_ 
             "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"' "&chr(13)&_
             "    AND work_date <> '"&work_date&"'                  "&chr(13)&_ 
             "    AND mg_ce_id = '"& user_id &"'                    "&chr(13)
  'Response.write weeksSql&"<br>"
  weeksRs.Open weeksSql, Dbconn, 1


  if (weeksRs.eof or weeksRs.bof) then
    	delta_minute = 0 
    	rest_minute  = 0
    	work_time    = 0
    	work_minute  = 0
  else
    	delta_minute = Cint( weeksRs("delta_minute") ) ' �Ѱ���ð��� �Ѻ����� .. (����,�̽��� ������� �Ѵ�)
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

  title_line = "��Ư�� ���� (�� ����["&fDate&" ~ "&lDate&"] �� ��Ư�� �ð�: "&work_time&"�ð� "&work_minute&"�� [�� "& FormatNumber(delta_minute,0) &" ��] )"


  ' �ش� ���� �� ������ �۾��ð� ���� ���Ѵ�. (���ΰǸ�..) (�����ϰ��� ������� ����)  work_date <> work_date
  weeksSql = " SELECT ifnull(sum(delta_minute),0)  delta_minute     "&chr(13)&_
             "      , ifnull(sum(rest_minute),0)   rest_minute      "&chr(13)&_
             "   FROM overtime A                                    "&chr(13)&_ 
             "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"' "&chr(13)&_
             "    AND work_date <> '"&work_date&"'                  "&chr(13)&_ 
             "    AND mg_ce_id = '"& user_id &"'                    "&chr(13)&_
             "    AND allow_yn = 'Y'                                "&chr(13)
  'Response.write "<pre>"&weeksSql&"</pre><br>"
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
'Response.write alterTimeOff1&"<br>"

  sumAlterTimeOff2 = 0
  sumAlterTimeOff3 = 0

  ' �ش� ���� �� ������ �۾��ð� ���� ���Ѵ�.  (���ΰǸ�..) (�����ϰ��� ������� ����)
  weeksSql = "     SELECT A.work_date                                      "&chr(13)&_
             "          , A.end_date                                       "&chr(13)&_
             "          , A.from_time                                      "&chr(13)&_
             "          , A.to_time                                        "&chr(13)&_
             "          , left( A.from_time,2 )  fromtime                  "&chr(13)&_
             "          , right( A.from_time,2 ) fromminute                "&chr(13)&_
             "          , left( A.to_time,2 )    totime                    "&chr(13)&_
             "          , right( A.to_time,2 )   tominute                  "&chr(13)&_
             "          , ifnull(DAYOFWEEK( A.work_date ),0) work_week     "&chr(13)&_
             "          , ifnull(DAYOFWEEK( A.end_date ),0)  end_week      "&chr(13)&_
             "          , ifnull(A.delta_minute,0)           delta_minute  "&chr(13)&_
             "          , ifnull(A.rest_minute,0)            rest_minute   "&chr(13)&_
             "          , ifnull(Floor(A.delta_minute/60),0) floor_time    "&chr(13)&_
             "          , ifnull(Mod(A.delta_minute,60),0)   mod_minute    "&chr(13)&_
             "          , H1.holiday_memo holiday1                         "&chr(13)&_ 
             "          , H2.holiday_memo holiday2                         "&chr(13)&_ 
             "       FROM overtime A                                       "&chr(13)&_ 
             "  LEFT JOIN holiday H1                                       "&chr(13)&_ 
             "         ON A.work_date = H1.holiday                         "&chr(13)&_ 
             "  LEFT JOIN holiday H2                                       "&chr(13)&_ 
             "         ON A.work_date = H2.holiday                         "&chr(13)&_ 
             "      WHERE A.work_date BETWEEN '"&fDate&"' AND '"&lDate&"'  "&chr(13)&_
             "        AND work_date <> '"&work_date&"'                     "&chr(13)&_ 
             "        AND A.mg_ce_id = '"& user_id &"'                     "&chr(13)&_
             "        AND A.allow_yn = 'Y'                                 "&chr(13)
  'Response.write weeksSql&"<br>"
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
      

  ' �ش� ���� �� ������ �۾��ð��� ���Ѵ�.  (�̽��� ��) (�����ϰ��� ������� ����)
  weeksSql = " SELECT ifnull(sum(delta_minute),0)   delta_minute     "&chr(13)&_
             "      , ifnull(sum(rest_minute),0)    rest_minute      "&chr(13)&_
             "   FROM overtime A                                     "&chr(13)&_ 
             "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'  "&chr(13)&_
             "    AND work_date <> '"&work_date&"'                   "&chr(13)&_ 
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


  	work_date1             = rs("work_date")
  	work_date2             = rs("end_date")
    mg_ce_id               = rs("mg_ce_id")
    company                = rs("company")
    dept                   = rs("dept")
    work_item              = rs("work_item")
    from_hh                = mid(rs("from_time"),1,2)
    from_mm                = mid(rs("from_time"),3,2)
    to_hh                  = mid(rs("to_time"),1,2)
    to_mm                  = mid(rs("to_time"),3,2)
    delta_minute           = rs("delta_minute")
    work_gubun             = rs("work_gubun")
    work_memo              = rs("work_memo")
    cancel_yn              = rs("cancel_yn")
    you_yn                 = rs("you_yn")
    end_yn                 = rs("end_yn")
    reg_id                 = rs("reg_id")
    reg_user               = rs("reg_user")
    reg_date               = rs("reg_date")
    mod_id                 = rs("mod_id")
    mod_user               = rs("mod_user")
    mod_date               = rs("mod_date")
    alter_timeoff_date     = rs("alter_timeoff_date") ' ��ü�޹��ð�
    alter_timeoff_time     = rs("alter_timeoff_time") ' ��ü�޹��ð�_��
    alter_timeoff_minute_w = rs("alter_timeoff_minute_w") ' ��ü�޹��ð�_52�ʰ��Ѻ�
    alter_timeoff_minute_d = rs("alter_timeoff_minute_d") ' ��ü�޹��ð�_8,22�ʰ��Ѻ�
    holiday1               = rs("holiday1")
    holiday2               = rs("holiday2")
    allow_yn               = rs("allow_yn")
  	rs.close()

  	sql = "select count(*) from overtime where acpt_no = "&int(acpt_no)
  	'Response.write "<pre>"&sql&"</pre><br>"
  	Set RsCount = Dbconn.Execute (sql)
  	tot_record = cint(RsCount(0)) 'Result.RecordCount

    if end_yn = "Y" then
    	  end_view = "����"
    else
      	end_view = "����"
    end if    

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

  <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
  "http://www.w3.org/TR/html4/loose.dtd">
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
  			    <% if to_hh <> "" then %>$("#to_hh").val(<%=to_hh%>);<% end if %>
  			    <% if to_mm <> "" then %>$("#to_m").val(<%=to_mm%>);<% end if %>
  			    <% if from_hh <> "" then %>$("#from_hh").val(<%=from_hh%>);<% end if %>
  			    <% if from_mm <> "" then %>$("#from_mm").val(<%=from_mm%>);<% end if %>
  			    
  			    getDeltaTime();
  			});
  			
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
   				  document.frm.delta_time.value = "";  
   				  document.frm.delta_minute.value = "0";  
   				  document.frm.rest_time.value = "";  
   				  document.frm.rest_minute.value = "0";  
   				  
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
     				
     				document.frm.rest_time.value = pad(parseInt(rest_minute / 60),2) + pad(rest_minute % 60,2);
     				document.frm.rest_minute.value = rest_minute;  
     				
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
     				
     				document.frm.delta_time.value = delta_time;  
     				document.frm.delta_minute.value = delta_minute;  
     				
     				<%
     				if allow_yn <> "Y" then
       				%>
       				if  (<%=delta_minute_aY%> >= (12*60))
       				  alterTimeOff1 = <%=delta_minute_aY%> - (12*60) ;
       				else
       				  alterTimeOff1 = 0;   				  
       				<%
     				else
       				%>  
       				// ���� �۾��ð��� �̹��� �۾��ð��� ���� 12�ð��� �ʰ��ϸ� �� �ʰ��и�.. (��ü�޹��ð� �� 52�ð��ʰ�)  (�����ϰ��� ���� �۾��ð� ���� ����)
       				if  ((<%=delta_minute_aY%> + (delta_minute - <%=delta_minute%>)) >= (12*60))
       				  alterTimeOff1 = (<%=delta_minute_aY%> + (delta_minute - <%=delta_minute%>)) - (12*60) ;
       				else
       				  alterTimeOff1 = 0;   				  
       				<%
     			  end if
     				%>  
     				  
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
                if  ((24*60 - (hh1*60+mm1))  >= (8*60) )
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
 				    else               			          alterTimeOff2 = 0;
			      if  (alterTimeOff3 > rest_minute) alterTimeOff3 = alterTimeOff3 - rest_minute;
 				    else               			          alterTimeOff3 = 0;
     				
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
  					document.frm.submit ();
  				}
  			}			
  			function chkfrm() {
  				if(document.frm.work_gubun.value =="") {
  					alert('�۾������� �����ϼ���');
  					frm.work_gubun.focus();
  					return false;
  				}
  					
  				if(document.frm.work_date1.value == "") {
  					alert('�۾��������ڰ� ����� �Ǿ� ���� �ʽ��ϴ� !!!');
  					frm.work_date1.focus();
  					return false;
  				}
  				else
  				{
  						if(document.frm.work_date1.value < "<%=dateNow%>") {
  							alert('�۾��������ڴ� ������ ���� �̿��� �մϴ�. !!!');
  							frm.work_date1.focus();
  							return false;
  						}
  				}
  				
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
  				  			
  				if(document.frm.work_memo.value =="" ) {
  					alert('�۾������� �Է��ϼ���');
  					frm.work_memo.focus();
  					return false;}  			
  				
  				if (confirm('�Է��Ͻðڽ��ϱ�?')==true) {
  					return true;
  				}
  				return false;
  				
  			}
  			function week_check() {
  			
  			a = document.frm.work_date.value.substring(0,4);
  			b = document.frm.work_date.value.substring(5,7);
  			c = document.frm.work_date.value.substring(8,10);
  			
  			var newDate = new Date(a,b-1,c); 
  			var s = newDate.getDay(); 
  			
  			switch(s) {
  				case 0: str = "�Ͽ���" ; break;
  				case 1: str = "������" ; break;
  				case 2: str = "ȭ����" ; break;
  				case 3: str = "������" ; break;
  				case 4: str = "�����" ; break;
  				case 5: str = "�ݿ���" ; break;
  				case 6: str = "�����" ; break;
  				}
  			
  				document.frm.week.value = str;			
  			}
  			function update_view() {
  					document.getElementById('cancel_col').style.display = '';
  					document.getElementById('info_col').style.display = '';
  			}
  			function delcheck() 
  			{
  				acpt_no    = document.frm.acpt_no.value;
  				tot_record = document.frm.tot_record.value;
  				
  				del_msg = "������ȣ " + acpt_no + "�� ��ü " + tot_record + "���� ������ �˴ϴ�. ���� �����Ͻðڽ��ϱ�?";
  				
  				if (confirm(del_msg)==true) {
  					document.frm.action = "overtime_del_ok.asp";
  					document.frm.submit();
    				return true;
    			}
  				return false;
  			}
  			
  			function pad(n, width) {
  			  n = n + '';
  			  return n.length >= width ? n : new Array(width - n.length + 1).join('0') + n;
  			}
        </script>
  	</head>
  	<body onload="update_view()">
  			<div id="container">
  				
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
  				
  				<form action="overtime_as_mod_15_save.asp" method="post" name="frm">
  				<div class="gView">
  					<table cellpadding="0" cellspacing="0" class="tableWrite">
  						<colgroup>
  							<col width="12%" >
  							<col width="38%" >
  							<col width="12%" >
  							<col width="*" >
  						</colgroup>
  						<tbody>
  							<tr>
  								<th class="first">�۾���</th>
  								<td class="left">
                      <input name="work_date1" id="work_date1" type="text" style="width:70px" value="<%=work_date1%>" readonly>
                      <input name="week1" id="week1" type="text" style="width:16px" readonly="true" value="<%=week1%>" readonly>
                      <input name="holiday1" id="holiday1" type="text" style="width:55px" readonly="true" value="<%=holiday1%>" readonly>
                      ~
    							  	<input name="work_date2" id="work_date2" type="text" style="width:70px" value="<%=work_date2%>">
                      <input name="week2" id="week2" type="text" style="width:16px" readonly="true" value="<%=week2%>">  
                      <input name="holiday2" id="holiday2" type="text" style="width:55px" readonly="true" value="<%=holiday2%>">                    
                  </td>
  								<th>�۾���</th>
  								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)
                      <input name="curr_date" type="hidden" id="curr_date" value="<%=now()%>">
                      <input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=mg_ce_id%>">
                  </td>
  							</tr>
  							<tr>
  								<th class="first">ȸ���</th>
  								<td class="left"><%=company%></td>
  								<th>�μ���</th>
  								<td class="left"><%=dept%></td>
  							</tr>
  							<tr>
  								<th class="first">�۾��׸�</th>
  								<td class="left">
  								    <%=work_item%><input name="work_item" type="hidden" id="work_item" value="<%=work_item%>">
              		</td>
  								<th>�۾��ð�</th>
  								<td class="left">
  								    <table>
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
                  </td>
  							</tr>
  							<tr>
  								<th class="first">�߱��׸�</th>
  								<td class="left">
                      <input name="work_gubun" type="text" value="<%=work_gubun%>" readonly="true" style="width:150px">
  				            <a href="#" onClick="pop_Window('overtime_code_search.asp?gubun=<%="AS"%>','overtime_code_pop','scrollbars=yes,width=700,height=400')" class="btnType03">��ȸ</a>
                  </td>
  								<th>�۾�����</th>
  								<td class="left">
  								    <input name="work_memo" type="text" id="work_memo" onKeyUp="checklength(this,50)"  style="width:100px" value="<%=work_memo%>">
    								  &nbsp;
                  	  <strong>* �����󱸺�</strong>
                  	  <label><input type="radio" name="you_yn" value="N" <% if you_yn = "N" then %>checked<% end if %> style="width:20px" id="Radio4">����</label>
  							      <label><input type="radio" name="you_yn" value="Y" <% if you_yn = "Y" then %>checked<% end if %> style="width:20px" id="Radio3">����</label>
  								</td>
  							</tr>  							
  							<tr id="cancel_col" style="display:none">
  								<th class="first">��ҿ���</th>
  								<td class="left">
  								    <label><input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:40px" ID="Radio1">���</label>
  				            <label><input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:40px" ID="Radio2">����</label>
  								</td>
                  <th>��������</th>
  								<td class="left"><%=end_view%><input name="end_yn" type="hidden" id="end_yn" value="<%=end_yn%>"></td>
  							</tr>
  							<tr id="info_col" style="display:none">
  								<th class="first">�������</th>
  								<td class="left"><%=reg_user%>&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
                  <th>��������</th>
  								<td class="left"><%=mod_user%>&nbsp;<%=mod_id%>(<%=mod_date%>)</td>
  							</tr>
  							<tr>
  							  <th>��ü�޹��ð����</th>
  								<td class="left" colspan="3">
  								    <strong>�� 52�ð� �ʰ� :</strong> <span id="alterTimeOff1_t"></span>�ð� <span id="alterTimeOff1_m"></span>�� [ <span id="alterTimeOff1"></span> ��]
                      &nbsp;
  								    <strong>���� 22�� �ʰ�:</strong> <span id="alterTimeOff2_t"></span>�ð� <span id="alterTimeOff2_m"></span>�� [ <span id="alterTimeOff2"></span> ��]
                      &nbsp;
  							  	  <strong>���� 8�ð� �ʰ� :</strong> <span id="alterTimeOff3_t"></span>�ð� <span id="alterTimeOff3_m"></span>�� [ <span id="alterTimeOff3"></span> ��]
  								</td>
  						  </tr>
  							<tr>
  							  <th class="first">��ü�޹��ð�<br>���ۼ���</th>
  							  <td class="left" colspan="3">
  							    <input name="alter_timeoff_date" id="alter_timeoff_date" type="text" style="width:80px" value="<%=alter_timeoff_date%>">
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

  							    <input type="hidden" name="reg_sw" value="<%=reg_sw%>" ID="reg_sw">
  							  </td>  							  
  						  </tr>
  						</tbody>
  					</table>
  				</div>
          
          <br>
          <div align=center>
            <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
            <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
  				  <%	
  					if (user_id = mg_ce_id or user_id = reg_id) then
  						if end_yn = "N" or end_yn = "C" then	
  				    %>
              <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
          		<%	
  						end if
  					end if	
  				  %>
          </div>
          
  				<input type="hidden" name="acpt_no" value="<%=acpt_no%>" ID="Hidden1">
          <input type="hidden" name="tot_record" value="<%=tot_record%>" ID="Hidden1">
          
          <input type="hidden" name="delta_time" value="<%=delta_time%>">
  			  <input type="hidden" name="delta_minute" value="<%=delta_minute%>">
          
          <input type="hidden" name="rest_time" value="<%=rest_time%>">
  			  <input type="hidden" name="rest_minute" value="<%=rest_minute%>">

  			  <input type="hidden" name="alter_timeoff_minute_w" id="alter_timeoff_minute_w">
					<input type="hidden" name="alter_timeoff_minute_d" id="alter_timeoff_minute_d">
					
					<input type="hidden" name="fDate" value="<%=fDate%>">
					<input type="hidden" name="lDate" value="<%=lDate%>">
					          
  			</form>
  		</div>				
  	</body>
  </html>
<%
  end if
%>
