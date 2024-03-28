<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim from_date, to_date
Dim rsCount
Dim fDate, lDate, work_time

' 야특근 승인권자 ID 리스트
allowerIDs = Array("100125","100029","100015","100031","100020") ' "강명석","이재원","전간수","최길성','홍건형'

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


	dateNow = Date()        ' 현재일자
  week    = Weekday(Date) ' 현재요일

  If  (week >= 4) Then
  		mGap = (week - 4) * -1  
  Else
  		mGap = (6 - (3-week)) * -1  
  End If

  ' 수요일 ~ 화요일(다음주)
  fDate = DateAdd("d", mGap, dateNow) 
  lDate = DateAdd("d", mGap + 6, dateNow)
  
  if  (from_date = "" ) then 
    from_date = fDate
  end if
  if  (to_date = "" ) then 
    to_date = lDate
  end if

if (view_condi = "") then
	view_condi = "케이원정보통신"
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
    	
' 포지션별
posi_sql = " AND A.mg_ce_id = '" + user_id + "'"&chr(13)

if position = "팀원" then
	view_condi = "본인"
end if

if position = "파트장" then
	if view_c = "total" then
		if org_name = "한화생명호남" then
			posi_sql = " AND (A.org_name = '한화생명호남' or A.org_name = '한화생명전북') "&chr(13)
		  else
			posi_sql = " AND A.org_name = '"&org_name&"'"&chr(13)
		end if
	  else
		if org_name = "한화생명호남" then
			posi_sql = " AND (A.org_name = '한화생명호남' or A.org_name = '한화생명전북') and M.user_name like '%"&mg_ce&"%'"&chr(13)
		  else
			posi_sql = " AND A.org_name = '"&org_name&"' and M.user_name like '%"&mg_ce&"%'"&chr(13)
		end if
	end if
end if

if position = "팀장" then
	if view_c = "total" then
		posi_sql = " AND A.team = '"&team&"'"&chr(13)
	  else
		posi_sql = " AND A.team = '"&team&"' and M.user_name like '%"&mg_ce&"%'"&chr(13)
	end if
end if

if position = "사업부장" or cost_grade = "2" then
	if view_c = "total" then
		posi_sql = " AND A.saupbu = '"&saupbu&"'"&chr(13)
	  else
		posi_sql = " AND A.saupbu = '"&saupbu&"' and M.user_name like '%"&mg_ce&"%'"&chr(13)
	end if
end if

if position = "본부장" or cost_grade = "1" then 
  	if view_c = "total" then
		posi_sql = " AND A.bonbu = '"&bonbu&"'"&chr(13)
 	  else
		posi_sql = " AND A.bonbu = '"&bonbu&"' and M.user_name like '%"&mg_ce&"%'"&chr(13)
	end if	 
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "전체"
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
  
  ' 작업종료일자가 없을 때(구버젼일때)만 시작시간과 종료시간으로만 작업경과시간을 계산한 후 작업종료시간을 삽입한다.
  if  IsNull(end_date) = True then 
	
		if to_time >= from_time then ' 정상적일때 (부터시간 < 까지시간)
			
				if tominute >= fromminute then ' 정상적일때 (분빼기)
						deltaminute = tominute - fromminute
				else ' 분이 까지가 더 크면 시에서 60을 빌려온다. (분빼기)
						deltaminute = (tominute+60)	- fromminute 
						totime =  totime - 1					
				end if		
				deltatime =  totime - fromtime '  (시빼기)
					
				end_date = Cdate(work_date) ' 같은날..
				'Response.write deltatime&":"&deltaminute&"   "
						
		else ' 까지가 작을때 (부터시간 > 까지시간)
				
				deltatime =  (24 - fromtime)
				if 0 < fromminute then ' 분이 있으면 시를 차감하고 60까지 잔여분을 계산한다.
						deltatime = deltatime - 1
						deltaminute = 60 - fromminute
				else
						deltaminute = 0
				end if
				
				' 자정까지 시와 분을 까지의 시분과 더한다.
				deltatime   = deltatime + totime
				deltaminute = deltaminute + tominute
				
				if deltaminute >= 60 then ' 더한 분이 60분 을 초과하면 시간을 추가하고 분을 60 이하로 맞춘다.
						deltatime   = deltatime + 1 
						deltaminute = deltaminute - 60
				end if
				
				end_date = DateAdd("d", 1, Cdate(work_date))  ' 다음날로 처리
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

title_line = "주 52시간 현황보기"
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
                allow_sayou = prompt('미승인 사유','');
                if (allow_sayou == null) return;
            }
            var work_date = $(this).attr("work_date"); // 작업일
            var mg_ce_id  = $(this).attr("mg_ce_id"); // cd 아이디

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
        							alert("변경됐습니다.");
        				      
        				      spanAllowSayou.text(allow_sayou);			
        				      
        						}else if( result=="invalid" ){
        							alert("입력하신 정보가 정확하지 않습니다.");
        						}else if(result=="fail"){
        							alert("저장 실패했습니다.");
        						}
        					}
        					,error: function(jqXHR, status, errorThrown){
        						alert("에러가 발생하였습니다.\n상태코드 : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
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
				  alert("검색 시작년월일이 없습니다.");
					return false;
				}
				
				if (lDate = "")
				{
				  alert("검색 종료년월일이 없습니다.");
					return false;
				}
				
				if ((fDate != "") && (lDate != "") && (fDate > lDate))
				{
					alert("검색 시작년월일이 종료 년월일 보다 작을 수 없습니다.");
					return false;
				}
				return true;
			}
			
			function condi_view() {
      <%
			if not (position = "팀원" and cost_grade <> "0") then
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
						<legend>조회영역</legend>
								<p style="position:relative">
									&nbsp;
									<label><strong>조회기간 : </strong></label>
									<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker1">
									 ~ 
									<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker2">
									&nbsp;
									&nbsp;
									<label><strong>조회권한 : </strong><%=position%></label>
								  <label><strong>조회범위 : </strong>
									<%
									if position = "팀원" and cost_grade <> "0" then
											Response.write view_condi ' "케이원정보통신"
									else
    									%>
    									<label><input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">조직전체</label>
    									<label><input type="radio" name="view_c" value="reg_id" <% if view_c = "reg_id" then %>checked<% end if %> style="width:25px" onClick="condi_view()">개인별</label>
    									<% 
									end if 
									%>
									</label>
									<label>
									<input name="mg_ce" type="text" value="<%=mg_ce%>" style="width:70px; display:none" id="mg_ce_view">
								  </label>
									
									<span style="position:absolute;right:105px; cursor: pointer; width:100px" class="btnType04" onclick="javascript:frmcheck();">검색</span>
									<span style="position:absolute;right:5px; cursor: pointer; width:80px" class="btnType04" onclick="javascript:frmcheckReset();">검색초기화</span>
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
									<th class="first" scope="col">사번</th>
									<th scope="col">이름</th>
									<th scope="col">사업부/본/지사</th>
									<th scope="col">야특근 날짜</th>									
									<th scope="col">총 시간</th>
									<th scope="col">대체휴가</th>
  								<th scope="col">승인<br>여부</th>
								  <th scope="col">미승인사유</th>
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
	              	<td><%=Rs("floor_time")%>시 <%=Rs("mod_minute")%>분</td>
	              	<td>
	              	    <%
	              	    if Rs("alter_timeoff_date") <> "" then '사용자가 대체휴가시작일을 입력했을 경우	              	        
	              	        %>
	              	        <%=Rs("alter_timeoff_date")%>&nbsp;<%=Rs("altertimeofftime")%>:<%=Rs("altertimeoffminute")%>
	              	        <br> ~ 
	              	        <%
	              	        if CInt(Rs("alter_timeoff_minute_w")) > 0 then ' 52시간 초과건을 경우
	              	          
  	              	        dateNow = CDate(Rs("work_date")) ' 일자변환
                            week    = Weekday(dateNow)       ' 요일  

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
	                          
	                          if  (last_cnt = 0) then  ' 마지막 52시간 초과건을 경우
                              Response.write Rs("alter_timeoff_enddate2") ' 주 52시간 초과 + (평일 22시 초과 + 휴일 8시간 초과)
                            else                     ' 52시간 초과건이지만 마지막건인 아닌경우
                              Response.write Rs("alter_timeoff_enddate1") ' (평일 22시 초과 + 휴일 8시간 초과)
                            end if
                          else ' 52시간 초과건이 아닌 경우
                            Response.write Rs("alter_timeoff_enddate1") ' (평일 22시 초과 + 휴일 8시간 초과)
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
	              	    <label><input type="radio" name="allow_yn_<%=Rs("mg_ce_id")%>_<%=Rs("work_date")%>" mg_ce_id="<%=Rs("mg_ce_id")%>" work_date="<%=Rs("work_date")%>" value="Y" <% if Rs("allow_yn") = "Y" then %>checked="checked"<% end if %> style="width:20px" id="Radio1">승인</label>
                	    <label><input type="radio" name="allow_yn_<%=Rs("mg_ce_id")%>_<%=Rs("work_date")%>" mg_ce_id="<%=Rs("mg_ce_id")%>" work_date="<%=Rs("work_date")%>" value="N" <% if Rs("allow_yn") = "N" then %>checked="checked"<% end if %> style="width:20px" id="Radio2">미승인</label>
                	    <label><input type="radio" name="allow_yn_<%=Rs("mg_ce_id")%>_<%=Rs("work_date")%>" mg_ce_id="<%=Rs("mg_ce_id")%>" work_date="<%=Rs("work_date")%>" value="X" <% if Rs("allow_yn") = "X" then %>checked="checked"<% end if %> style="width:20px" id="Radio3">미확인</label>
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
								  <td colspan="8">조건에 일치하는 데이터가 존재하지 않습니다.</td>
								</tr>
								<% end if %>
							</tbody>
						</table>
					</div>

					<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  	<tr>
				    	<td width="15%">
					      <div class="btnCenter">
                    <a href="overtime_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&view_c=<%=view_c%>&mg_ce=<%=mg_ce%>" class="btnType04">엑셀다운로드</a>
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

