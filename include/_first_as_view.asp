<!--#include virtual="/common/inc_top.asp"--><!--설정 파일-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"--><!--nkpmg_user.asp 변수 선언-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" --><!--사용자 정의 함수-->
<%
'=========================================
'author : 허정호
'modify date : 20201126
'Desc :
'	include file 추가
'	변수 선언 추가 및 사용 객체 소멸 처리
'=========================================

'=========================================
'### DB Connect
'=========================================
Dim DBConn

Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DBConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder

Set objBuilder = New StringBuilder

'=========================================
'### Request Param
'=========================================
Dim com_tab
Dim com_sum(7)
Dim ok_sum(7)
Dim mi_sum(7)
'dim com_cnt(7,9)
'dim com_in(7,9)
'dim sum_cnt(9)
'dim sum_in(9)
Dim com_cnt(7,10)
Dim com_in(7,10)
Dim sum_cnt(10)
Dim sum_in(10)
Dim company_tab(150)
Dim end_tab(11)
Dim mi_tab(11)
Dim curr_mi_tab(11)
Dim mi_in

Dim sql, Rs

Dim rs_etc, rs_trade
Dim type_sql, mg_group_sql

'Dim rs_wek

Dim i, j
Dim curr_day, curr_date, to_date, as_type
Dim company
Dim tot_cnt
Dim dd, a, d, com_date
Dim title_line
Dim k
Dim sido

'공휴일 배열 처리 =============================
Dim rs_hol
Dim strHoli, holiCnt, idx

objBuilder.Append "SELECT holiday FROM holiday ORDER BY holiday ASC "
Set rs_hol = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

strHoli = rs_hol.GetRows()
holiCnt = UBound(strHoli, 2)

'For i=0 To holiCnt
'	Response.write strHoli(0, i)
'Next

rs_hol.Close
Set rs_hol = Nothing
'공휴일 배열 처리 =============================

com_tab = array("본사", "부산지사", "대구지사", "대전지사", "광주지사", "전주지사", "강원지사", "제주지사")

For i = 0 To 7
	com_sum(i) = 0
	ok_sum(i) = 0
	mi_sum(i) = 0
	'for j = 0 to 9
	For j = 0 To 10
		com_cnt(i,j) = 0
		com_in(i,j) = 0
		sum_cnt(j) = 0
		sum_in(j) = 0
	Next
Next

For i = 0 To 11
	end_tab(i) = 0
	mi_tab(i) = 0
	curr_mi_tab(i) = 0
Next

curr_day = DateValue(Mid(CStr(Now()), 1, 10))
curr_date = DateValue(Mid(DateAdd("h", 12, Now()), 1, 10))
to_date = Mid(CStr(Now()), 1, 10)
as_type = "방문처리"
company = "전체"
mg_group = "1"

'Set Rs_etc = Server.CreateObject("ADODB.Recordset")
'Set rs_trade = Server.CreateObject("ADODB.Recordset")

type_sql = " (as_type ='방문처리') and "
'type_sql = " (as_acpt.as_type ='방문처리') and "
mg_group_sql = " (mg_group ='1') and "

tot_cnt = 0

' 미처리건
'sql = "select as_acpt.sido, as_acpt.as_process, Cast(acpt_date as date) as acpt_day, CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) as com_date, count(*) as err_cnt from as_acpt"
'sql = sql + " WHERE "+type_sql+mg_group_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
'sql = sql + " GROUP BY sido, as_process, Cast(acpt_date as date), CAST((as_acpt.acpt_date + interval 10 DAY_HOUR) as date) Order By as_acpt.sido Asc"


'sql = "select as_acpt.sido, as_acpt.as_process, Cast(request_date as date) as acpt_day, CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date) as com_date, count(*) as err_cnt from as_acpt"
'sql = sql + " WHERE "+type_sql+mg_group_sql+" (as_process = '접수' or as_process = '입고' or as_process = '연기')"
'sql = sql + " AND CAST(request_date AS DATE) <= now()"
'sql = sql + " GROUP BY sido, as_process, Cast(request_date as date), CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date) Order By as_acpt.sido Asc"

' 충북제천시와 단양군이 대전지사에서 강원지사로 배정이 변경됨 (2018-11-16)  정상원 과장 요구
sql = " select *                                                                                                                    "&chr(13)&_
      " from                                                                                                                        "&chr(13)&_
      " (                                                                                                                           "&chr(13)&_
      " select as_acpt.sido                                                                                                         "&chr(13)&_
      "      , as_acpt.as_process                                                                                                   "&chr(13)&_
      "      , Cast(request_date as date) as acpt_day                                                                               "&chr(13)&_
      "      , CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date) as com_date                                              "&chr(13)&_
      "      , count(*) as err_cnt                                                                                                  "&chr(13)&_
      "   from as_acpt                                                                                                              "&chr(13)&_
      "  WHERE (as_type ='방문처리') and (mg_group ='1') and (as_process = '접수' or as_process = '입고' or as_process = '연기')    "&chr(13)&_
      "    AND CAST(request_date AS DATE) <= now()                                                                                  "&chr(13)&_
      "    and (sido <> '충북' and sido <> '강원')                                                                                  "&chr(13)&_
      "  GROUP BY sido                                                                                                              "&chr(13)&_
      "          ,as_process                                                                                                        "&chr(13)&_
      "          ,Cast(request_date as date), CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date)                           "&chr(13)&_
      " union all                                                                                                                   "&chr(13)&_
      " select '충북'                                                                                                               "&chr(13)&_
      "      , as_acpt.as_process                                                                                                   "&chr(13)&_
      "      , Cast(request_date as date) as acpt_day                                                                               "&chr(13)&_
      "      , CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date) as com_date                                              "&chr(13)&_
      "      , count(*) as err_cnt                                                                                                  "&chr(13)&_
      "   from as_acpt                                                                                                              "&chr(13)&_
      "  WHERE (as_type ='방문처리') and (mg_group ='1') and (as_process = '접수' or as_process = '입고' or as_process = '연기')    "&chr(13)&_
      "    AND CAST(request_date AS DATE) <= now()                                                                                  "&chr(13)&_
      "    and (sido = '충북' and (gugun <> '제천시' and gugun <> '단양군'))                                                        "&chr(13)&_
      "  GROUP BY sido                                                                                                              "&chr(13)&_
      "          ,as_process                                                                                                        "&chr(13)&_
      "          ,Cast(request_date as date), CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date)                           "&chr(13)&_
      " union all                                                                                                                   "&chr(13)&_
      " select '강원'                                                                                                               "&chr(13)&_
      "      , as_acpt.as_process                                                                                                   "&chr(13)&_
      "      , Cast(request_date as date) as acpt_day                                                                               "&chr(13)&_
      "      , CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date) as com_date                                              "&chr(13)&_
      "      , count(*) as err_cnt                                                                                                  "&chr(13)&_
      "   from as_acpt                                                                                                              "&chr(13)&_
      "  WHERE (as_type ='방문처리') and (mg_group ='1') and (as_process = '접수' or as_process = '입고' or as_process = '연기')    "&chr(13)&_
      "    AND CAST(request_date AS DATE) <= now()                                                                                  "&chr(13)&_
      "    and (sido = '강원' or (gugun = '제천시' or gugun = '단양군'))                                                            "&chr(13)&_
      "  GROUP BY sido                                                                                                              "&chr(13)&_
      "          ,as_process                                                                                                        "&chr(13)&_
      "          ,Cast(request_date as date), CAST((as_acpt.request_date + interval 10 DAY_HOUR) as date)                           "&chr(13)&_
      "  ) a                                                                                                                        "&chr(13)&_
      "  Order By sido Asc, acpt_day ASC                                                                                                           "

' 방문일 변경checking
' select a.*, if(trim(old) <> '' and old <> new,'변동','') visit_changed
' from
' (
' select a.*, concat(visit_date_old,visit_time_old ) old
'      , concat(d.visit_date,d.visit_time ) new
' from as_acpt a
' left join as_mod_visit_datetime  d
'  on a.acpt_no = d.acpt_no
' where a.visit_date = '2019-01-01'
' ) a
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open Sql, Dbconn, 1

Do Until Rs.EOF
'	com_tab(i) = rs("sido")
	Select Case Rs("sido")
		Case "서울": i = 0
		Case "경기": i = 0
		Case "인천": i = 0
		Case "부산": i = 1
		Case "울산": i = 1
		Case "경남": i = 1
		Case "대구": i = 2
		Case "경북": i = 2
		Case "대전": i = 3
		Case "충남": i = 3
		Case "충북": i = 3
		Case "세종": i = 3
		Case "광주": i = 4
		Case "전남": i = 4
		Case "전북": i = 4 ' 5 ->4  전북이 광주지사로 편입 (2018.09.27 변경)
		Case "강원": i = 6
		Case "제주": i = 7
	End Select

	dd = DateDiff("d", Rs("com_date"), curr_date)

	If dd < 0 Then
		dd = 0
	End If

	If CStr(curr_day) = CStr(Rs("acpt_day")) Then
		dd = 0
	End If

	'휴일 계산
	If dd > 0 Then
		a = DateDiff("d", Rs("acpt_day"), curr_day)
		'b = datepart("w",rs("acpt_day"))
		'bb = datepart("w", curr_day)
		'if bb = 1 then
		'	a = a -1
		'end if
		'c = a + b
		d = a
		'if a > 1 then
		'	if c > 7 then
		'		d = a - 2
		'	end if
		'end if

		com_date = DateValue(Rs("acpt_day"))

		'주말 체크
		Do Until com_date > curr_day
			'쿼리 사용 대신 내부 함수로 처리[허정호_20201126]
			'sql_hol = "select * from (select DAYOFWEEK('" + cstr(com_date) + "') as  dayweeks ) A where A.dayweeks in (1,7)"
			'Set rs_wek=DbConn.Execute(SQL_hol)

			'If rs_wek.eof or rs_wek.bof Then
			'	d = d
			'Else
			'	d = d -1
			'End If

			If Weekday(CStr(com_date)) = "1" Or Weekday(CStr(com_date)) = "7" Then
				d = d - 1
			End If

 			com_date = DateAdd("d", 1, com_date)

			'rs_wek.close()
		Loop

'		visit_date = rs("visit_date")
		com_date = datevalue(rs("acpt_day"))
'		act_date = com_date

		'공휴일 체크
		Do Until com_date > curr_day
			'쿼리 사용 대신 배열로 처리
			'sql_hol = "select * from holiday where holiday = '" + cstr(com_date) + "'"
			'Set rs_hol=DbConn.Execute(SQL_hol)
			'if rs_hol.eof or rs_hol.bof then
			'	d = d
			'  else
			'	d = d -1
			'end If

			'공휴일 배열과 비교 후 해당 일자가 있을 경우
			For idx=0 To holiCnt
				If strHoli(0, idx) = CStr(com_date) Then
					d = d - 1

					Exit For
				End If
			Next

			com_date = dateadd("d",1,com_date)
			'rs_hol.close()
		Loop

' 2012-02-06
		If d = 1 Then
			curr_hh = Int(DatePart("h",Now()))

			If rs("acpt_day") <> rs("com_date") And curr_hh < 12 Then
				d = 0
			End If
		End If

' 2012-02-06 end
		If d = 0 Then '당일
			j = 5
		ElseIf d = 1 Then '익일
			j = 6
		ElseIf d = 2 Then '2일
			j = 7
'		  elseif d > 2 and d < 7  then
'			j = 8
'		  else
'			j = 9
		ElseIf d = 3 Then '3일
			j = 8
		ElseIf d = 4 Then '4일
			j = 9
		Else  '5일이상
			j = 10
		End If

		com_cnt(i, j) = com_cnt(i, j) + CLng(Rs("err_cnt"))

		If rs("as_process") = "입고" Then
			com_in(i,j) = com_in(i,j) + CLng(Rs("err_cnt"))
		End If
	  Else
' 휴일 계산 끝
		com_cnt(i,5) = com_cnt(i,5) + CLng(Rs("err_cnt"))
		'com_cnt(i,6) = com_cnt(i,6) + clng(rs("err_cnt"))

		If rs("as_process") = "입고" Then
			com_in(i,5) = com_in(i,5) + CLng(Rs("err_cnt"))
			'com_in(i,6) = com_in(i,6) + clng(rs("err_cnt"))
		End If
	End If
	tot_cnt = tot_cnt + CLng(Rs("err_cnt"))

	Rs.MoveNext()
Loop

Rs.close()
Set Rs = Nothing

title_line = "방문처리 지사별 미처리 현황 (요청일 기준)"

DBConn.Close
Set DBConn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<!-- <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" /> -->
		<link href="/include/style.css" type="text/css" rel="stylesheet">

	    <script src="/java/jquery-1.9.1.js"></script>
	    <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

    <script type="text/javascript">

      function setCookie(cname, cvalue, exdays) {
          var d = new Date();
          d.setTime(d.getTime() + (exdays*24*60*60*1000));
          var expires = "expires="+ d.toUTCString();
          document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/";
      }

      // '오늘만 이 창을 열지 않음' 클릭
      function closePop()
      {
       	setCookie('first_as_view', 'first_as_view', 1);
        self.close();
      }

    </script>

	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
				<div class="gView" style="position: relative;">
					<h3 class="stit">* 현재시간 : <%=Now()%></h3>
					<table cellpadding="0" cellspacing="0" class="tableList3" style="width:1000px;">
						<colgroup>
							<col width="*" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">지사</th>

								<th colspan="2" style=" border-left:1px solid #e3e3e3;border-bottom:1px solid #e3e3e3;" scope="col">당일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">익일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">2일</th>
								<!--
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">3일~6일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">7일이상</th>
								-->
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">3일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">4일</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">5일이상</th>
								<th colspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">소계</th>
								<th rowspan="2" style="border-bottom:1px solid #e3e3e3;" scope="col">백분율</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">건수</th>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">건수</th>
							  <th style=" border-left:1px solid #e3e3e3;" scope="col">입고</th>
						  </tr>
						</thead>
						<tbody>
						<%
                    	If tot_cnt > 0 Then
                        	k = 0
                      	Else
                        	k = 7
                    	End If

		'--------------------------------------여기 확인
                    	For i = k To 7
                        	If com_tab(i) <> "" Then

								For j = 0 To 4
									ok_sum(i) = ok_sum(i) + com_cnt(i,j)
									sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)
								Next
								'for j = 5 to 9
								For j = 5 To 10
									mi_sum(i) = mi_sum(i) + com_cnt(i,j)
									sum_cnt(j) = sum_cnt(j) + com_cnt(i,j)
									sum_in(j) = sum_in(j) + com_in(i,j)
								Next
								com_sum(i) = ok_sum(i) + mi_sum(i)

								sido = com_tab(i)
							End If
						Next
		'--------------------------------------여기 확인
                		%>
							<tr>
                              <th>계</th>
                              <th class="right"><%=FormatNumber(CLng(sum_cnt(5)),0)%></a></th>
                              <th class="right"><%=sum_in(5)%></th>
                              <th class="right"><%=FormatNumber(CLng(sum_cnt(6)),0)%></a></th>
                              <th class="right"><%=sum_in(6)%></th>
                              <th class="right"><%=FormatNumber(CLng(sum_cnt(7)),0)%></a></th>
                              <th class="right"><%=sum_in(7)%></th>
                              <th class="right"><%=FormatNumber(CLng(sum_cnt(8)),0)%></a></th>
                              <th class="right"><%=sum_in(8)%></th>
                              <th class="right"><%=FormatNumber(CLng(sum_cnt(9)),0)%></a></th>
                              <th class="right"><%=sum_in(9)%></th>
                              <th class="right"><%=FormatNumber(CLng(sum_cnt(10)),0)%></a></th>
                              <th class="right"><%=sum_in(10)%></th>
                              <th class="right"><%=FormatNumber(CLng(sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)+sum_cnt(10)),0)%></th>
                              <th class="right"><%=sum_in(5)+sum_in(6)+sum_in(7)+sum_in(8)+sum_in(9)+sum_in(10)%></th>
                              <th class="right">
                              <% If tot_cnt = 0 Then %>
                                    0%
                              <% Else %>
								<%=FormatNumber(((sum_cnt(0)+sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)+sum_cnt(10))/tot_cnt * 100),2)%>%
                              <% End If %>
                              &nbsp;
                              </th>
							</tr>
						<%
                    	If tot_cnt > 0 Then
                        	k = 0
                      	Else
                        	k = 7
                    	End If

                    	For i = k To 7
                        	If com_tab(i) <> "" Then
                        	  ' 전북지사가 없어짐 (2018.09.27 변경)
                        		If i <> 5 Then
                		%>
							<tr>
                              <td><%=com_tab(i)%></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=0%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=FormatNumber(CLng(com_cnt(i,5)),0)%></td>
                              <td class="right"><%=com_in(i,5)%></td>
                              <td class="right"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=1%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=FormatNumber(CLng(com_cnt(i,6)),0)%></td>
                              <td class="right"><%=com_in(i,6)%></td>
                              <td class="right" bgcolor="#FFFF88"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=2%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=FormatNumber(CLng(com_cnt(i,7)),0)%></strong></td>
                              <td class="right"><strong><%=com_in(i,7)%></strong></td>
                              <td class="right" bgcolor="#FFBE7D"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=3%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=FormatNumber(CLng(com_cnt(i,8)),0)%></strong></td>
                              <td class="right"><strong><%=com_in(i,8)%></strong></td>
                              <td class="right" bgcolor="#FF8080"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=4%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=FormatNumber(CLng(com_cnt(i,9)),0)%></strong></td>
                              <td class="right"><strong><%=com_in(i,9)%></strong></td>

<!-- 추가 4일 -->
                              <td class="right" bgcolor="#FF8080"><a  href="#" onClick="pop_Window('day_michulri_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>&days=<%=5%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><strong><%=FormatNumber(CLng(com_cnt(i,10)),0)%></strong></td>
                              <td class="right"><strong><%=com_in(i,10)%></strong></td>
<!-- 추가 4일 -->

                              <td class="right"><a  href="#" onClick="pop_Window('as_michulri_popup_request.asp?from_date=<%="1900-01-01"%>&to_date=<%=to_date%>&sido=<%=com_tab(i)%>&company=<%=company%>&as_type=<%=as_type%>&mg_group=<%=mg_group%>','as_mi_popup','scrollbars=yes,width=800,height=600')"><%=FormatNumber(CLng(mi_sum(i)),0)%></td>
                              <td class="right"><%=com_in(i,5)+com_in(i,6)+com_in(i,7)+com_in(i,8)+com_in(i,9)+com_in(i,10)%></td>
                              <td class="right">
                              <% If tot_cnt = 0 Then %>
								0%
                              <% Else %>
								<%=FormatNumber((com_sum(i)/tot_cnt * 100),2)%>%
                              <% End If %>
                              &nbsp;
                              </td>
							</tr>
                		<%
                			End If ' 전북지사가 없어짐 (2018.09.27 변경)
						   End If
						Next
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	충북제천시와 단양군이 대전지사에서 강원지사로 배정

	<table cellpadding="0" cellspacing="0" style="width:1000px;">
  <TR>
    <TD width="585" height="25" valign="middle">
      <div align="right">
      <span class="style1"><strong>오늘만 이 창을 열지 않음</strong></span>
      <input name="todayPop" type="checkbox" id="todayPop" onClick="closePop()" value="checkbox">
      </div>
    </TD>
  </TR>
  </TABLE>

	</body>
</html>
