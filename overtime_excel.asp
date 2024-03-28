<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

' 야특근 승인권자 ID 리스트
allowerIDs = Array("100125","100029","100015","100031","100020","100018") ' "강명석","이재원","전간수","최길성','홍건형','송지영'

from_date = Request("from_date")
to_date   = Request("to_date")
view_c    = Request("view_c")
mg_ce     = Request("mg_ce")

'Response.write mg_ce

savefilename = "야특근 현황("+from_date+"_"+to_date+").xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsChk = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

' 포지션별
posi_sql = " AND mg_ce_id = '" + user_id + "'"&chr(13)

if position = "팀원" then
	view_condi = "본인"
end if

if position = "파트장" then
	if view_c = "total" then
		if org_name = "한화생명호남" then
			posi_sql = " AND (org_name = '한화생명호남' or org_name = '한화생명전북') "&chr(13)
		else
			posi_sql = " AND org_name = '"&org_name&"'"&chr(13)
		end if
	  else
		if org_name = "한화생명호남" then
			posi_sql = " AND (org_name = '한화생명호남' or org_name = '한화생명전북') and user_name like '%"&mg_ce&"%'"&chr(13)
		else
			posi_sql = " AND org_name = '"&org_name&"' and user_name like '%"&mg_ce&"%'"&chr(13)
		end if
	end if
end if

if position = "팀장" then
	if view_c = "total" then
		posi_sql = " AND team = '"&team&"'"&chr(13)
	else
		posi_sql = " AND team = '"&team&"' and user_name like '%"&mg_ce&"%'"&chr(13)
	end if
end if

if position = "사업부장" or cost_grade = "2" then
	if view_c = "total" then
		'posi_sql = " AND saupbu = '"&saupbu&"'"&chr(13)
        posi_sql = " and A.saupbu = emp_master.emp_saupbu "&chr(13)
	else
        'posi_sql = " AND saupbu = '"&saupbu&"' and user_name like '%"&mg_ce&"%'"&chr(13)
        posi_sql = " AND A.saupbu = emp_master.emp_saupbu and user_name like '%"&mg_ce&"%'"&chr(13)
	end if
end if

if position = "본부장" or cost_grade = "1" then 
	if view_c = "total" then
	  posi_sql = " AND bonbu = '"&bonbu&"'"&chr(13)
  else
	  posi_sql = " AND bonbu = '"&bonbu&"' and user_name like '%"&mg_ce&"%'"&chr(13)
	end if	 
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "전체"
  	if view_c = "total" then
		posi_sql = ""
 	else
		posi_sql = " AND user_name like '%"&mg_ce&"%'"&chr(13)
	end if	 
end if

base_sql = " SELECT A.mg_ce_id                                                                    "&chr(13)&_
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
           "      , ifnull(A.delta_minute,0) delta_minute                                         "&chr(13)&_
           "      , ifnull(A.rest_minute,0)  rest_minute                                          "&chr(13)&_
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
           "      , A.cancel_yn                                                                   "&chr(13)&_
           "      , A.you_yn                                                                      "&chr(13)&_
           "      , A.reside_place                                                                "&chr(13)&_
           "      , A.user_name                                                                   "&chr(13)&_
           "      , A.user_grade                                                                  "&chr(13)&_
           "      , A.company                                                                     "&chr(13)&_
		   "      , A.dept                                                                        "&chr(13)&_
		   "      , A.cost_center                                                                 "&chr(13)&_
		   "      , A.work_gubun                                                                  "&chr(13)&_
		   "      , A.work_memo                                                                   "&chr(13)&_
		   "      , A.overtime_amt                                                                "&chr(13)&_
           "   FROM overtime A                                                                    "&chr(13)&_
           "inner join emp_master                                                                 "&chr(13)&_           
           "        ON emp_master.emp_no = A.mg_ce_id                                             "&chr(13)
date_sql = "  WHERE work_date BETWEEN '"& from_date &"' AND '"& to_date &"'                       "&chr(13)

sql = base_sql + date_sql + posi_sql &_
    " ORDER BY org_name, user_name, work_date"

'Response.write "<pre>"&sql&"</pre><br>"

Rs.Open Sql, Dbconn, 1
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">회사</th>
								<th scope="col">본부</th>
								<th scope="col">사업부</th>
								<th scope="col">팀</th>
								<th scope="col">조직명</th>
								<th scope="col">상주처</th>
								<th scope="col">사번</th>
								<th scope="col">작업자</th>
								<th scope="col">야특근 시작</th>
								<th scope="col">야특근 끝</th>
								<th scope="col">총시간</th>
								<th scope="col">대체휴가</th>
								<th scope="col">AS NO</th>
								<th scope="col">회사</th>
								<th scope="col">조직명</th>
								<th scope="col">비용유형</th>
								<th scope="col">야특근구분</th>
								<th scope="col">작업내역</th>
								<th scope="col">신청금액</th>
								<th scope="col">유무상</th>
								<th scope="col">지급</th>
								<th scope="col">승인</th>
								<th scope="col">미승인사유</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
						  
						    delta_minute = Cint( Rs("delta_minute") ) ' 총경과시간을 총분으로 .. 
                            rest_minute  = Cint( Rs("rest_minute") )  ' 총휴게시간을 총분으로 .. 
                            
                            if (delta_minute > rest_minute) then
                                delta_minute = delta_minute - rest_minute
                            else
                                delta_minute = 0
                            end if
                            work_time   = Fix(delta_minute / 60) ' 총작업시간을 시로 ..
                            work_minute = delta_minute mod 60    ' 총작업시간을 시로 나눈몫인 분으로 ..

							if  rs("cancel_yn") = "Y" then
								cancel_yn = "취소"
							  else
								cancel_yn = "지급"
							end if
							if rs("acpt_no") = 0 or rs("acpt_no") = null then
								acpt_no = "없음"
							  else
								acpt_no = rs("acpt_no")
							end if 

							if rs("you_yn") = "Y" then
								you_view = "유상"
							else
							 	you_view = "무상"
							end if
                            %>
                            <tr>
                                <td class="first"><%=rs("emp_company")%></td>
                                <td><%=rs("bonbu")%></td>
                                <td><%=rs("saupbu")%></td>
                                <td><%=rs("team")%></td>
                                <td><%=rs("org_name")%></td>
                                <td><%=rs("reside_place")%></td>
                                <td><%=rs("mg_ce_id")%></td>
                                <td><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%></td>								
                                
                                <td><%=Rs("work_date")%>&nbsp;<%=Rs("fromtime")%>:<%=Rs("fromminute")%></td>
                                <td><%=Rs("end_date")%>&nbsp;<%=Rs("totime")%>:<%=Rs("tominute")%></td>
                                <td><%=work_time%>시간 <%=work_minute%>분</td>
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
                                                "    AND mg_ce_id  = '"& rs("mg_ce_id") &"'               "&chr(13)&_
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
								<td><%=acpt_no%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("cost_center")%></td>
								<td><%=rs("work_gubun")%></td>
								<td><%=rs("work_memo")%></td>
								<%
  								find = False 
                                For i = 0 To uBound(allowerIDs)
                                    if  user_id = allowerIDs(i) then
                                        find =True 
                                    end if
                                Next
                                
                                if find = True then
                                    %><td class="right"><%=formatnumber(rs("overtime_amt"),0)%></td><%
                                end if
  							    %>								
								<td><%=you_view%></td>
                                <td><%=cancel_yn%></td>
                                								
								<td><%=Rs("allow_yn")%></td>
								<td>
								    <span name ="allowSayou"><%=Rs("allow_sayou")%></span>
								</td>
							</tr>
						    <%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

