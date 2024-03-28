<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim com_tab(15)
Dim com_sum(15)
Dim ok_sum(15)
Dim mi_sum(15)
Dim com_cnt(15,7)
Dim sum_cnt(7)
Dim company_tab(150)
Dim end_tab(8)
Dim mi_tab(8)

Dim from_date, to_date, curr_day, curr_date, sido, mg_ce, mg_ce_id
Dim company, as_type, days, win_sw, dis_days
Dim memo01, memo02, com_sql, type_sql, i, in_cnt, acpt_cnt, yun_cnt
Dim title_line, rsAs, arrAs

from_date = f_Request("from_date")
to_date = f_Request("to_date")
sido = f_Request("sido")
mg_ce = f_Request("mg_ce")
mg_ce_id = f_Request("mg_ce_id")
mg_group = f_Request("mg_group")
company = f_Request("company")
as_type = f_Request("as_type")
days = Int(f_Request("days"))

title_line = "기간별 미처리 현황"

curr_day = DateValue(Mid(CStr(Now()), 1, 10))
curr_date = DateValue(Mid(DateAdd("h",12,now()),1,10))

win_sw = "back"
dis_days = CStr(days) & "일"

'if days = 3 then
'	dis_days = "3~6일"
'end if

'if days = 7 then
If days = 5 Then
	dis_days = "5일이상"
End If

If company = "" Then
	company = "전체"
	as_type = "전체"
End If

If mg_ce = "" Then
	memo01 = "시도"
	memo02 = sido
Else
	memo01 = "담당자"
	memo02 = mg_ce
End If

If company = "전체" Then
	com_sql = ""
Else
  	com_sql = "company ='"&company&"' AND "
End If

If as_type = "전체" Then
	type_sql = ""
Else
  	type_sql = "as_type ='"&as_type&"' AND "
End If

i = 0
in_cnt = 0
acpt_cnt = 0
yun_cnt = 0

objBuilder.Append "SELECT acpt_no, request_date, as_process, company, dept, sido, gugun, as_type "
objBuilder.Append "FROM as_acpt "
objBuilder.Append "WHERE "&com_sql&type_sql&" (as_process = '접수' OR as_process = '입고' OR as_process = '연기') "
objBuilder.Append "	AND (Cast(request_date as date) >= '"&from_date&"' AND Cast(request_date as date) <= '"&to_date&"') "

' 미처리건
If mg_ce = "" Then
	Select Case sido
		Case "총괄", "계"
			objBuilder.Append ""
		Case "본사"
			objBuilder.Append " AND sido IN ('서울', '경기', '인천') "
		Case "부산지사"
			objBuilder.Append "	AND sido IN ('부산', '경남', '울산') "
		Case "대구지사"
			objBuilder.Append "	AND sido IN ('대구', '경북') "
		Case "대전지사"
			objBuilder.Append "	AND sido IN ('대전', '충남', '충북', '세종') "
			objBuilder.Append "	AND (GUGUN <> '제천시' AND GUGUN <> '단양군') "	 ' 충북제천시와 단양군이 대전지사에서 강원지사로 배정이 변경됨 (2019.01.18)  정상원 과장 요구
		Case "광주지사"
			objBuilder.Append "	AND sido IN ('광주', '전남', '전북') "
		Case "제주지사"
			objBuilder.Append "	AND sido = '제주' "
		Case "강원지사"
			objBuilder.Append "	AND sido = '강원' "
			objBuilder.Append "	OR (GUGUN = '제천시' OR GUGUN = '단양군') "	 ' 충북제천시와 단양군이 대전지사에서 강원지사로 배정이 변경됨 (2019.01.18)  정상원 과장 요구
		Case Else
			objBuilder.Append "	AND sido = '"&sido&"' "
	End Select
Else
	If mg_ce <> "총괄" Then
		objBuilder.Append " AND mg_ce_id = '"&mg_ce_id&"' "
	End If
End If

Set rsAs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsAs.EOF Then
	arrAs = rsAs.getRows()
End If
rsAs.Close() : Set rsAs = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>기간별 미처리 현황</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction(){
		  		 window.close();
			}
        </script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="container">
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th><%=memo01%></th>
							  <td class="left"><%=memo02%></td>
							  <th>회사</th>
							  <td class="left"><%=company%></td>
							  <th>처리유형</th>
							  <td class="left"><%=as_type%></td>
							</tr>
                            <tr>
							  <th>기간</th>
							  <td class="left"><%=dis_days%></td>
							  <td colspan="4">
								<a href = "/main/excel/day_michulri_excel_request.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=sido%>&company=<%=company%>&as_type=<%=as_type%>&mg_ce=<%=mg_ce%>&mg_ce_id=<%=mg_ce_id%>&mg_group=<%=mg_group%>&days=<%=days%>" class="btnType04">엑셀다운로드</a>
							  </td>
					      	</tr>
						</tbody>
					</table>
					<br>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="15%" >
							<col width="5%" >
							<col width="18%" >
							<col width="25%" >
							<col width="*" >
							<col width="10%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">요청일자</th>
								<th scope="col">상태</th>
								<th scope="col">회사명</th>
								<th scope="col">부서명</th>
								<th scope="col">지역</th>
								<th scope="col">처리유형</th>
								<th scope="col">조회</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim seq, as_acpt_no, as_request_date, as_process
						Dim as_company, as_dept, as_sido, as_gugun, as_as_type
						'Int date_len
						Dim len_date, hangle, bit01, bit02, bit03, l
						Dim com_date, dd, a, d
						Dim rs_week, rs_hol, acpt_date, date_to_date, curr_hh, acpt_hh

						If IsArray(arrAs) Then
							seq = 0
							For l = LBound(arrAs) To UBound(arrAs, 2)
								as_acpt_no = arrAs(0, l)
								as_request_date = arrAs(1, l)
								as_process = arrAs(2, l)
								as_company = arrAs(3, l)
								as_dept = arrAs(4, l)
								as_sido = arrAs(5, l)
								as_gugun = arrAs(6, l)
								as_as_type = arrAs(7, l)

	                            seq = seq + 1

								com_date = DateValue(Mid(DateAdd("h", 10, as_request_date), 1, 10))
								'com_date = datevalue(mid(rs("acpt_date"),1,10))
								dd = DateDiff("d", com_date, curr_date)
								'ddd = dd

								'휴일 계산
								If dd < 0 Then
									dd = 0
								End If

								If CStr(curr_day) = CStr(Mid(as_request_date, 1, 10)) Then
									dd = 0
								End If

								If dd > 0 Then
									com_date = DateValue(Mid(as_request_date, 1, 10))
									'a = dd
									a = DateDiff("d", com_date, curr_day)
									'b = datepart("w", com_date)
									'bb = datepart("w", curr_day)
									'if bb = 1 then
									'    a = a -1
									'end if
									'c = a + b
									d = a
									'if a > 1 then
									'    if c > 7 then
									'        d = a - 2
									'    end if
									'end if

									Do Until com_date > curr_day
										objBuilder.Append "SELECT dayweeks FROM (SELECT DAYOFWEEK('"&CStr(com_date)&"') AS dayweeks) A WHERE A.dayweeks IN (1, 7) "
										Set rs_week = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										If rs_week.EOF Or rs_week.BOF Then
											d = d
										Else
											d = d -1
										End If

										com_date = DateAdd("d", 1, com_date)
										rs_week.Close()
									Loop
									Set rs_week = Nothing

									'visit_date = rs("visit_date")
									'com_date = datevalue(rs("acpt_date"))
									'act_date = com_date

									com_date = DateValue(Mid(as_request_date, 1, 10))

									Do Until com_date > curr_day
										objBuilder.Append "SELECT holiday FROM holiday WHERE holiday = '"&CStr(com_date)&"' "

										Set rs_hol = DbConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										If rs_hol.EOF Or  rs_hol.BOF Then
											d = d
										Else
											d = d -1
										End If

										com_date = DateAdd("d", 1, com_date)
										rs_hol.Close()
									Loop
									Set rs_hol = Nothing

									' 2012-02-06
									If d = 1 Then
										curr_hh = Int(DatePart("h", Now()))
										acpt_hh = int(DatePart("h", as_request_date))

										If acpt_hh > 13 And curr_hh < 12 Then
											d = 0
										End If
									End If

									' 2012-02-06 end
									dd = d
									'if d > 2 and d < 7 then
									'    dd = 3
									'end if
									'if d > 6 then
										'dd = 7
									If d > 4 Then
										dd = 5
									End If
								  Else
								' 휴일 계산 끝
									dd = 0
								End If

								'date_len=len(rs("acpt_date"))

								acpt_date = as_request_date
								len_date = Len(acpt_date)
								bit01 = Left(acpt_date, 10)
								'bit01 = Replace(bit01,"-",".")
								bit03 = Left(Right(acpt_date, 5), 2)
								hangle = Mid(acpt_date, 12, 2)

								If len_date = 22 Then
									bit02 = Mid(acpt_date, 15, 2)
								Else
									bit02 = "0" & Mid(acpt_date, 15, 1)
								End If

								If hangle = "오후" And bit02 <> 12 Then
									bit02 = bit02 + 12
								End If

								date_to_date = bit01 & " " &bit02 & ":" & bit03
								acpt_date = Mid(date_to_date, 3)
								'acpt_date = replace(acpt_date,"-","/")
								acpt_date = as_request_date

								If dd = days Then
									If as_process = "접수" Then
										acpt_cnt = acpt_cnt + 1
									End If

									If as_process = "연기" Then
										yun_cnt = yun_cnt + 1
									End If

									If as_process = "입고" Then
										in_cnt = in_cnt + 1
									End If

									i = i + 1
                        %>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=acpt_date%></td>
								<td><%=as_process%></td>
								<td><%=as_company%></td>
								<td><%=as_dept%></td>
								<td><%=as_sido%>&nbsp;<%=as_gugun%></td>
								<td><%=as_as_type%></td>
								<td>
									<a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=as_acpt_no%>&win_sw=<%=win_sw%>','asview_pop','scrollbars=yes,width=800,height=700')">조회</a>
								</td>
							</tr>
							<%
                                End If
								Next
							End If
							DBConn.Close() : Set DBConn = Nothing
                            %>
						</tbody>
					</table>
					<br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>접수</th>
							  <td class="left"><%=acpt_cnt%></td>
							  <th>연기</th>
							  <td class="left"><%=yun_cnt%></td>
							  <th>입고</th>
							  <td class="left"><%=in_cnt%></td>
					      	</tr>
						</tbody>
					</table>
					<br>
				</form>
				</div>
			</div>
	</body>
</html>