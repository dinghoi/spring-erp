<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim from_date, to_date
Dim rsCount
Dim fDate, lDate, work_time

' ��Ư�� ���α��� ID ����Ʈ
allowerIDs = Array("100125","100029","100015","100031","100020") ' "����","�����","������","�ֱ漺','ȫ����'

view_c     = Request.form("view_c")
mg_ce      = Request.form("mg_ce")
from_date  = Request.form("from_date")
to_date    = Request.form("to_date")


Set Dbconn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsLoop = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
Set RsChk = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

view_condi = Request("view_condi")


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
  
  if  (from_date = "" ) then 
    from_date = fDate
  end if
  if  (to_date = "" ) then 
    to_date = lDate
  end if

if (view_condi = "") then
	view_condi = "���̿��������"
end if

SQLDEFAULT = " SELECT A.mg_ce_id                                                                    "&chr(13)&_
             "      , A.work_date                                                                   "&chr(13)&_
             "      , A.end_date                                                                    "&chr(13)&_
             "      , A.from_time                                                                   "&chr(13)&_
             "      , A.to_time                                                                     "&chr(13)&_
             "      , left(A.to_time,2)    totime                                                   "&chr(13)&_
             "      , left(A.from_time,2)  fromtime                                                 "&chr(13)&_
             "      , right(A.to_time,2)   tominute                                                 "&chr(13)&_
             "      , right(A.from_time,2) fromminute                                               "&chr(13)&_
             "      , A.acpt_no                                                                     "&chr(13)&_
             "      , A.user_name                                                                   "&chr(13)&_
             "      , A.emp_company                                                                 "&chr(13)&_
             "      , A.bonbu                                                                       "&chr(13)&_
             "      , A.saupbu                                                                      "&chr(13)&_
             "      , A.team                                                                        "&chr(13)&_
             "      , A.org_name                                                                    "&chr(13)&_
             "      , A.cost_detail                                                                 "&chr(13)&_ 
             "      , A.delta_minute                                                                "&chr(13)&_
             "      , Floor(ifnull(A.delta_minute,0)/60) floor_time                                 "&chr(13)&_
             "      , Mod(ifnull(A.delta_minute,0),60)   mod_minute                                 "&chr(13)&_             
             "      , A.alter_timeoff_date                                                          "&chr(13)&_
             "      , A.alter_timeoff_time                                                          "&chr(13)&_
             "      , left(A.alter_timeoff_time,2)  altertimeofftime                                "&chr(13)&_
             "      , right(A.alter_timeoff_time,2) altertimeoffminute                              "&chr(13)&_

             "      , A.alter_timeoff_minute_w                                                      "&chr(13)&_
             "      , A.alter_timeoff_minute_d                                                      "&chr(13)&_
             "      , DATE_FORMAT(date_add(A.alter_timeoff_date, interval (A.alter_timeoff_minute_d) minute), '%Y-%m-%d %I:%i') alter_timeoff_enddate1                          "&chr(13)&_
             "      , DATE_FORMAT(date_add(A.alter_timeoff_date, interval (A.alter_timeoff_minute_w+A.alter_timeoff_minute_d) minute), '%Y-%m-%d %I:%i') alter_timeoff_enddate2 "&chr(13)&_

             "      , (SELECT visit_date FROM as_acpt WHERE acpt_no = A.acpt_no ) AS visit_date     "&chr(13)&_ 
             "      , A.allow_yn                                                                    "&chr(13)&_
             "      , A.allow_sayou                                                                 "&chr(13)&_
		         "   FROM overtime A                                                                    "&chr(13)&_
		         "  WHERE work_date BETWEEN '"&from_date&"' AND '"&to_date&"'                           "&chr(13)
      if (Request("emp_company") <> "") then
				SQLAND1 = "  AND emp_company LIKE '%"&Request("emp_company")&"%'  "&chr(13)
			else
				SQLAND1 = ""&chr(13)
    	end if
    	if (Request("emp_bonbu") <> "") then
    		SQLAND2 = "  AND bonbu LIKE '%"&Request("emp_bonbu")&"%'  "&chr(13)
    	else
    		SQLAND2 = ""&chr(13)
    	end if
    	if (Request("emp_saupbu") <> "") then
    		SQLAND3 = "  AND saupbu LIKE '%"&Request("emp_saupbu")&"%'  "&chr(13)
    	else
    		SQLAND3 = ""&chr(13)
    	end if
    	if (Request("emp_team") <> "") then
				SQLAND4 = "  AND team LIKE '%"&Request("emp_team")&"%'  "&chr(13)
			else
				SQLAND4 = ""&chr(13)
    	end if
    	
' �����Ǻ�
posi_sql = " AND A.mg_ce_id = '" + user_id + "'"&chr(13)

if position = "����" then
	view_condi = "����"
end if

if position = "��Ʈ��" then
	if view_c = "total" then
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " AND (A.org_name = '��ȭ����ȣ��' or A.org_name = '��ȭ��������') "&chr(13)
		  else
			posi_sql = " AND A.org_name = '"&org_name&"'"&chr(13)
		end if
	  else
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " AND (A.org_name = '��ȭ����ȣ��' or A.org_name = '��ȭ��������') and M.user_name like '%"&mg_ce&"%'"&chr(13)
		  else
			posi_sql = " AND A.org_name = '"&org_name&"' and M.user_name like '%"&mg_ce&"%'"&chr(13)
		end if
	end if
end if

if position = "����" then
	if view_c = "total" then
		posi_sql = " AND A.team = '"&team&"'"&chr(13)
	  else
		posi_sql = " AND A.team = '"&team&"' and M.user_name like '%"&mg_ce&"%'"&chr(13)
	end if
end if

if position = "�������" or cost_grade = "2" then
	if view_c = "total" then
		posi_sql = " AND A.saupbu = '"&saupbu&"'"&chr(13)
	  else
		posi_sql = " AND A.saupbu = '"&saupbu&"' and M.user_name like '%"&mg_ce&"%'"&chr(13)
	end if
end if

if position = "������" or cost_grade = "1" then 
  	if view_c = "total" then
		posi_sql = " AND A.bonbu = '"&bonbu&"'"&chr(13)
 	  else
		posi_sql = " AND A.bonbu = '"&bonbu&"' and M.user_name like '%"&mg_ce&"%'"&chr(13)
	end if	 
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "��ü"
  	if view_c = "total" then
		posi_sql = ""&chr(13)
 	  else
		posi_sql = " AND M.user_name like '%"&mg_ce&"%'"&chr(13)
	end if	 
end if

base_sql = " SELECT A.mg_ce_id                                                                       "&chr(13)&_
           "         , A.work_date                                                                   "&chr(13)&_
           "         , A.end_date                                                                    "&chr(13)&_
           "         , A.from_time                                                                   "&chr(13)&_
           "         , A.to_time                                                                     "&chr(13)&_
           "         , left(A.to_time,2)    totime                                                   "&chr(13)&_
           "         , left(A.from_time,2)  fromtime                                                 "&chr(13)&_
           "         , right(A.to_time,2)   tominute                                                 "&chr(13)&_
           "         , right(A.from_time,2) fromminute                                               "&chr(13)&_
           "         , A.acpt_no                                                                     "&chr(13)&_
           "         , A.user_name                                                                   "&chr(13)&_
           "         , A.emp_company                                                                 "&chr(13)&_
           "         , A.bonbu                                                                       "&chr(13)&_
           "         , A.saupbu                                                                      "&chr(13)&_
           "         , A.team                                                                        "&chr(13)&_
           "         , A.org_name                                                                    "&chr(13)&_
           "         , A.cost_detail                                                                 "&chr(13)&_ 
           "         , A.delta_minute - A.rest_minute     delta_minute                               "&chr(13)&_
           "         , Floor(ifnull(A.delta_minute,0)/60) floor_time                                 "&chr(13)&_
           "         , Mod(ifnull(A.delta_minute,0),60)   mod_minute                                 "&chr(13)&_             
           "         , A.alter_timeoff_date                                                          "&chr(13)&_
           "         , A.alter_timeoff_time                                                          "&chr(13)&_
           "         , left(A.alter_timeoff_time,2)  altertimeofftime                                "&chr(13)&_
           "         , right(A.alter_timeoff_time,2) altertimeoffminute                              "&chr(13)&_
           "         , A.alter_timeoff_minute_w                                                      "&chr(13)&_
           "         , A.alter_timeoff_minute_d                                                      "&chr(13)&_
           "         , DATE_FORMAT(date_add(A.alter_timeoff_date, interval (A.alter_timeoff_minute_d) minute), '%Y-%m-%d %I:%i') alter_timeoff_enddate1                          "&chr(13)&_
           "         , DATE_FORMAT(date_add(A.alter_timeoff_date, interval (A.alter_timeoff_minute_w+A.alter_timeoff_minute_d) minute), '%Y-%m-%d %I:%i') alter_timeoff_enddate2 "&chr(13)&_
           "         , (SELECT visit_date FROM as_acpt WHERE acpt_no = A.acpt_no ) AS visit_date     "&chr(13)&_ 
           "         , A.allow_yn                                                                    "&chr(13)&_
           "         , A.allow_sayou                                                                 "&chr(13)&_
		       "      FROM overtime A                                                                    "&chr(13)&_
		       "INNER JOIN memb M                                                                        "&chr(13)&_
		       "        ON A.mg_ce_id = M.user_id                                                        "&chr(13)		         
date_sql = "  WHERE work_date BETWEEN '"&from_date&"' AND '"&to_date&"'                              "&chr(13)

sql = base_sql + date_sql + posi_sql + " ORDER BY A.org_name, M.user_name, A.work_date"
'Response.write "<pre>"&sql&"</pre><br>"
RsLoop.Open sql, Dbconn, 1

do until RsLoop.eof

  work_date  = RsLoop("work_date")	      
  end_date   = RsLoop("end_date")                           
	mg_ce_id   = RsLoop("mg_ce_id")                                 
	
	to_time    = RsLoop("to_time")
	from_time  = RsLoop("from_time")
	
	totime     = Cint( RsLoop("totime") )
	fromtime   = Cint( RsLoop("fromtime") )
	tominute   = Cint( RsLoop("tominute") )
	fromminute = Cint( RsLoop("fromminute") )

  'response.write IsNull(end_date)&"_ "
  
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
	  
  RsLoop.movenext()
Loop  
RsLoop.close()



Rs.CursorType = 3
Rs.CursorLocation = 3
Rs.LockType = 3
Rs.Open SQL, Dbconn, 1 

rsCount = Rs.RecordCount

title_line = "�� 52�ð� ��Ȳ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		
		<title><%=title_line %></title>
		
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		
		<script src="/java/jquery-1.9.1.js"></script>
		<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		
		<script type="text/javascript">
		  
		  $(document).ready(function () {
          $("input:radio").change(function () {
            
            if  ($(this).prop('name') == 'view_c') return ;
            
            var parent  = $(this).parent().parent().parent() ;
            var spanAllowSayou = parent.find("span[name='allowSayou']")            
            
            var allow_yn  = $(this).val(); // N or Y or X
            var allow_sayou = '' ;
            if  (allow_yn=='N')
            {
                allow_sayou = prompt('�̽��� ����','');
                if (allow_sayou == null) return;
            }
            var work_date = $(this).attr("work_date"); // �۾���
            var mg_ce_id  = $(this).attr("mg_ce_id"); // cd ���̵�

            var params = { "work_date" : work_date 
        								 , "mg_ce_id" : mg_ce_id
        								 , "allow_yn" : allow_yn
        								 , "allow_sayou" : escape(allow_sayou)
        								 };

            $.ajax({
        					 url: "ajax_set_overtime_allowYN.asp"
        					,type: 'post'
        					,data: params
        					,dataType: "json"
        					,contentType: "application/x-www-form-urlencoded; charset=euc-kr"
        					,beforeSend: function(jqXHR){
        							jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
        					}
        					,success: function(data){
        						var result = data.result;
        						if( result=="succ"){
        							alert("����ƽ��ϴ�.");
        				      
        				      spanAllowSayou.text(allow_sayou);			
        				      
        						}else if( result=="invalid" ){
        							alert("�Է��Ͻ� ������ ��Ȯ���� �ʽ��ϴ�.");
        						}else if(result=="fail"){
        							alert("���� �����߽��ϴ�.");
        						}
        					}
        					,error: function(jqXHR, status, errorThrown){
        						alert("������ �߻��Ͽ����ϴ�.\n�����ڵ� : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
        					}
    				});
          });
      });
      
			function getPageCode(){
				return "0 1";
			}
			
			$(function() {
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=from_date%>" );

				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "<%=to_date%>" );
			});	  

			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function frmcheckReset () {
				var fDate = "<%=DateAdd("d", (Weekday(Date)-2)*(-1), date)  %>";
				var lDate = "<%=DateAdd("d", (Weekday(Date)-2)*(-1)+6, date)%>";
				
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=from_date%>" );
				
				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "<%=to_date%>" );
				
				document.frm.submit ();
			}
			
			function chkfrm() {
				var fDate = $("#datepicker1").val();
				var lDate = $("#datepicker2").val();
				
				if (fDate = "")
				{
				  alert("�˻� ���۳������ �����ϴ�.");
					return false;
				}
				
				if (lDate = "")
				{
				  alert("�˻� ���������� �����ϴ�.");
					return false;
				}
				
				if ((fDate != "") && (lDate != "") && (fDate > lDate))
				{
					alert("�˻� ���۳������ ���� ����� ���� ���� �� �����ϴ�.");
					return false;
				}
				return true;
			}
			
			function condi_view() {
      <%
			if not (position = "����" and cost_grade <> "0") then
					%>
  				if (eval("document.frm.view_c[0].checked")) {
  					document.getElementById('mg_ce_view').style.display = 'none';
  				}	
  				if (eval("document.frm.view_c[1].checked")) {
  					document.getElementById('mg_ce_view').style.display = '';
  				}	
  				<% 
			end if 
			%>
			}
		</script>

	</head>
	
	<body onLoad="condi_view()">
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="overtime_stats.asp" method="post" name="frm" id="frm">
					<input type="hidden" id="emp_company"        name="emp_company"        value="<%=emp_company%>" />
					<input type="hidden" id="emp_bonbu"          name="emp_bonbu"          value="<%=emp_bonbu%>"   />
					<input type="hidden" id="emp_saupbu"         name="emp_saupbu"         value="<%=emp_saupbu%>"  />
					<input type="hidden" id="emp_team"           name="emp_team"           value="<%=emp_team%>"    />
					<input type="hidden" id="emp_reside_place"   name="emp_reside_place"   value="" />
					<input type="hidden" id="emp_reside_company" name="emp_reside_company" value="" />
					<input type="hidden" id="emp_org_level"      name="emp_org_level"      value="" />
					<input type="hidden" id="cost_center"        name="cost_center"        value="" />
					<input type="hidden" id="cost_group"         name="cost_group"         value="" />
					
					<fieldset class="srch">
						<legend>��ȸ����</legend>
								<p style="position:relative">
									&nbsp;
									<label><strong>��ȸ�Ⱓ : </strong></label>
									<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker1">
									 ~ 
									<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker2">
									&nbsp;
									&nbsp;
									<label><strong>��ȸ���� : </strong><%=position%></label>
								  <label><strong>��ȸ���� : </strong>
									<%
									if position = "����" and cost_grade <> "0" then
											Response.write view_condi ' "���̿��������"
									else
    									%>
    									<label><input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">������ü</label>
    									<label><input type="radio" name="view_c" value="reg_id" <% if view_c = "reg_id" then %>checked<% end if %> style="width:25px" onClick="condi_view()">���κ�</label>
    									<% 
									end if 
									%>
									</label>
									<label>
									<input name="mg_ce" type="text" value="<%=mg_ce%>" style="width:70px; display:none" id="mg_ce_view">
								  </label>
									
									<span style="position:absolute;right:105px; cursor: pointer; width:100px" class="btnType04" onclick="javascript:frmcheck();">�˻�</span>
									<span style="position:absolute;right:5px; cursor: pointer; width:80px" class="btnType04" onclick="javascript:frmcheckReset();">�˻��ʱ�ȭ</span>
              	</p>
					</fieldset>
					<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
							<colgroup>
								<col width="5%" />
								<col width="10%" />
								<col width="*" />
								<col width="11%" />								
								<col width="6%" />
								<col width="11%" />
								<%
  								find = False 
                  For i = 0 To uBound(allowerIDs)
                    if  user_id = allowerIDs(i) then
                        find =True 
                    end if
                  Next
                  
                  if find = True then
                      width = 16 
                  else   
                      width = 4 
                  end if
  							%>
								<col width="<%=width%>%" />
								<col width="15%" />
							</colgroup>
							<thead>
								<tr>
									<th class="first" scope="col">���</th>
									<th scope="col">�̸�</th>
									<th scope="col">�����/��/����</th>
									<th scope="col">��Ư�� ��¥</th>									
									<th scope="col">�� �ð�</th>
									<th scope="col">��ü�ް�</th>
  								<th scope="col">����<br>����</th>
								  <th scope="col">�̽��λ���</th>
								</tr>
							</thead>
							<tbody>
								<%
								if (rsCount > 0) then 
									 	do until rs.eof
								%>
								<tr>
									<td class="first"><%=Rs("mg_ce_id")%></td>
									<td><%=Rs("user_name")%></td>
									<td style="text-align:left;">&nbsp;&nbsp;<%=Rs("emp_company") & "&nbsp;>&nbsp;" & Rs("bonbu") & "&nbsp;>&nbsp;" &  Rs("team") %></td>
									<td><%=Rs("work_date")%>&nbsp;<%=Rs("fromtime")%>:<%=Rs("fromminute")%> 
									    <br> ~ 
									    <%=Rs("end_date")%>&nbsp;<%=Rs("totime")%>:<%=Rs("tominute")%>
									</td>
	              	<td><%=Rs("floor_time")%>�� <%=Rs("mod_minute")%>��</td>
	              	<td>
	              	    <%
	              	    if Rs("alter_timeoff_date") <> "" then '����ڰ� ��ü�ް��������� �Է����� ���	              	        
	              	        %>
	              	        <%=Rs("alter_timeoff_date")%>&nbsp;<%=Rs("altertimeofftime")%>:<%=Rs("altertimeoffminute")%>
	              	        <br> ~ 
	              	        <%
	              	        if CInt(Rs("alter_timeoff_minute_w")) > 0 then ' 52�ð� �ʰ����� ���
	              	          
  	              	        dateNow = CDate(Rs("work_date")) ' ���ں�ȯ
                            week    = Weekday(dateNow)       ' ����  

                            If  (week >= 4) Then
                            		mGap = (week - 4) * -1  
                            Else
                            		mGap = (6 - (3-week)) * -1  
                            End If

                            fDate = DateAdd("d", mGap, dateNow) 
                            lDate = DateAdd("d", mGap + 6, dateNow)
  	              	      
  	              	        chkSql =  " SELECT count(*) last_cnt                                "&chr(13)&_
  	              	                  "   FROM overtime                                         "&chr(13)&_
  	              	                  "  WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"'    "&chr(13)&_
                                      "    AND mg_ce_id  = '"& mg_ce_id &"'                     "&chr(13)&_
                                      "    AND length(alter_timeoff_date) > 0                   "&chr(13)&_
                                      "    AND work_date > '"& Rs("work_date") &"'              "&chr(13)
                            'Response.write "<pre>"&chkSql&"</pre><br>"
                            RsChk.Open chkSql, Dbconn, 1
  
                            last_cnt = 0
	                          If not (RsChk.bof or RsChk.eof) Then
	                              last_cnt = CInt(RsChk("last_cnt"))
	                          end if
	                          RsChk.close()
	                          
	                          if  (last_cnt = 0) then  ' ������ 52�ð� �ʰ����� ���
                              Response.write Rs("alter_timeoff_enddate2") ' �� 52�ð� �ʰ� + (���� 22�� �ʰ� + ���� 8�ð� �ʰ�)
                            else                     ' 52�ð� �ʰ��������� ���������� �ƴѰ��
                              Response.write Rs("alter_timeoff_enddate1") ' (���� 22�� �ʰ� + ���� 8�ð� �ʰ�)
                            end if
                          else ' 52�ð� �ʰ����� �ƴ� ���
                            Response.write Rs("alter_timeoff_enddate1") ' (���� 22�� �ʰ� + ���� 8�ð� �ʰ�)
                          end if
	              	    end if
	              	    %>									    
									</td>
  								<td>
  								<%
  								find = False 
                  For i = 0 To uBound(allowerIDs)
                    if  user_id = allowerIDs(i) then
                        find =True 
                    end if
                  Next
                  
                  if find = True then
  								%>
	              	    <label><input type="radio" name="allow_yn_<%=Rs("mg_ce_id")%>_<%=Rs("work_date")%>" mg_ce_id="<%=Rs("mg_ce_id")%>" work_date="<%=Rs("work_date")%>" value="Y" <% if Rs("allow_yn") = "Y" then %>checked="checked"<% end if %> style="width:20px" id="Radio1">����</label>
                	    <label><input type="radio" name="allow_yn_<%=Rs("mg_ce_id")%>_<%=Rs("work_date")%>" mg_ce_id="<%=Rs("mg_ce_id")%>" work_date="<%=Rs("work_date")%>" value="N" <% if Rs("allow_yn") = "N" then %>checked="checked"<% end if %> style="width:20px" id="Radio2">�̽���</label>
                	    <label><input type="radio" name="allow_yn_<%=Rs("mg_ce_id")%>_<%=Rs("work_date")%>" mg_ce_id="<%=Rs("mg_ce_id")%>" work_date="<%=Rs("work_date")%>" value="X" <% if Rs("allow_yn") = "X" then %>checked="checked"<% end if %> style="width:20px" id="Radio3">��Ȯ��</label>
  								<%
  							  else
  							  %>
                      <%=Rs("allow_yn")%>
  							  <%
  							  end if
  								%>
  								</td>
  								<td>
  								    <span name ="allowSayou"><%=Rs("allow_sayou")%></span>
  								</td>
								</tr>
								<% 
										Rs.movenext()
										Loop
										Rs.close()
									else 
								%>
								<tr>
								  <td colspan="8">���ǿ� ��ġ�ϴ� �����Ͱ� �������� �ʽ��ϴ�.</td>
								</tr>
								<% end if %>
							</tbody>
						</table>
					</div>

					<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  	<tr>
				    	<td width="15%">
					      <div class="btnCenter">
                    <a href="overtime_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&view_c=<%=view_c%>&mg_ce=<%=mg_ce%>" class="btnType04">�����ٿ�ε�</a>
					      </div>
              </td>
				    	<td width="85%"></td>
			      </tr>
				  </table>
				  
				</form>
			</div>				
		</div>        				
	</body>
</html>

