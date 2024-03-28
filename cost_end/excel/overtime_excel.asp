<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim allowerIDs
Dim from_date, to_date, view_c, mg_ce, savefilename
Dim posi_sql, view_condi, view_grade, rsOT

' 야특근 승인권자 ID 리스트
allowerIDs = Array("100125","100029","100015","100031","100020","100018") ' "강명석","이재원","전간수","최길성','홍건형','송지영'

from_date = Request("from_date")
to_date = Request("to_date")
view_c = Request("view_c")
mg_ce = Request("mg_ce")

savefilename = "야특근 현황("&from_date&"_"&to_date&").xls"

'엑셀 다운로드 지정
Call ViewExcelType(savefilename)

' 포지션별
posi_sql = " AND mg_ce_id = '"&user_id&"'"

If position = "팀원" Then
	view_condi = "본인"
End If

If position = "파트장" Then
	If view_c = "total" Then
		If org_name = "한화생명호남" Then
			posi_sql = "AND (org_name = '한화생명호남' OR org_name = '한화생명전북') "
		Else
			posi_sql = "AND org_name = '"&org_name&"' "
		End If
	Else
		If org_name = "한화생명호남" Then
			posi_sql = "AND (org_name = '한화생명호남' OR org_name = '한화생명전북') AND user_name LIKE '%"&mg_ce&"%' "
		Else
			posi_sql = "AND org_name = '"&org_name&"' AND user_name LIKE '%"&mg_ce&"%' "
		End If
	End If
End If

If position = "팀장" Then
	If view_c = "total" Then
		posi_sql = "AND ovrt.team = '"&team&"' "
	Else
		posi_sql = "AND ovrt.team = '"&team&"' AND user_name LIKE '%"&mg_ce&"%' "
	End If
End If

If position = "사업부장" Or cost_grade = "2" Then
	If view_c = "total" Then
        posi_sql = "AND ovrt.saupbu = emtt.emp_saupbu "
	Else
        posi_sql = "AND ovrt.saupbu = emtt.emp_saupbu AND user_name LIKE '%"&mg_ce&"%' "
	End If
End If

If position = "본부장" Or cost_grade = "1" Then
	If view_c = "total" Then
	  posi_sql = "AND ovrt.bonbu = '"&bonbu&"' "
	Else
	  posi_sql = "AND ovrt.bonbu = '"&bonbu&"' AND user_name LIKE '%"&mg_ce&"%' "
	End If
End If

view_grade = position

If cost_grade = "0" Then
	view_grade = "전체"

  	If view_c = "total" Then
		posi_sql = ""
 	Else
		posi_sql = " AND user_name LIKE '%"&mg_ce&"%' "
	End If
End If

objBuilder.Append "SELECT ovrt.mg_ce_id, ovrt.work_date, ovrt.end_date, "
objBuilder.Append "	ovrt.from_time, ovrt.to_time, "
objBuilder.Append "	LEFT(ovrt.to_time, 2) AS totime,"
objBuilder.Append "	LEFT(ovrt.from_time, 2) AS fromtime, "
objBuilder.Append "	RIGHT(ovrt.to_time, 2) AS tominute, "
objBuilder.Append "	RIGHT(ovrt.from_time, 2) AS fromminute,"
objBuilder.Append "	ovrt.acpt_no, ovrt.user_name, ovrt.cost_detail,"
objBuilder.Append "	IFNULL(ovrt.delta_minute, 0) AS delta_minute,"
objBuilder.Append "	IFNULL(ovrt.rest_minute, 0) AS rest_minute,"
objBuilder.Append "	ovrt.alter_timeoff_date, ovrt.alter_timeoff_time, "
objBuilder.Append "	LEFT(ovrt.alter_timeoff_time, 2) AS altertimeofftime, "
objBuilder.Append "	RIGHT(ovrt.alter_timeoff_time, 2) AS altertimeoffminute, "
objBuilder.Append "	ovrt.alter_timeoff_minute_w, ovrt.alter_timeoff_minute_d, "
objBuilder.Append "	DATE_FORMAT(DATE_ADD(ovrt.alter_timeoff_date, "
objBuilder.Append "		INTERVAL(ovrt.alter_timeoff_minute_d) MINUTE), "
objBuilder.Append "		'%Y-%m-%d %I:%i') AS alter_timeoff_enddate1, "
objBuilder.Append "	DATE_FORMAT(DATE_ADD(ovrt.alter_timeoff_date, "
objBuilder.Append "	INTERVAL(ovrt.alter_timeoff_minute_w + ovrt.alter_timeoff_minute_d) MINUTE), "
objBuilder.Append "	'%Y-%m-%d %I:%i') AS alter_timeoff_enddate2, "
objBuilder.Append "	(SELECT visit_date FROM as_acpt WHERE acpt_no = ovrt.acpt_no) AS visit_date, "
objBuilder.Append "	ovrt.allow_yn, ovrt.allow_sayou, ovrt.cancel_yn, ovrt.you_yn, ovrt.reside_place, "
objBuilder.Append "	ovrt.user_name, ovrt.user_grade, ovrt.company, ovrt.dept, ovrt.cost_center, "
objBuilder.Append "	ovrt.work_gubun, ovrt.work_memo, ovrt.overtime_amt, "
objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "	eomt.org_name "
objBuilder.Append "FROM overtime AS ovrt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON ovrt.mg_ce_id = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE work_date BETWEEN '"&from_date&"' AND '"&to_date&"' "
objBuilder.Append posi_sql
objBuilder.Append "ORDER BY eomt.org_name, ovrt.user_name, ovrt.work_date "

'response.write objBuilder.ToString()
'Response.end
Set rsOT = Server.CreateObject("ADODB.RecordSet")
rsOT.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
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
						Dim delta_minute, rest_minute, work_time, work_minute
						Dim cancel_yn, acpt_no, you_view, find, i
						Dim dateNow, week, mGap, fDate, lDate, rsChk
						Dim last_cnt

						Set rsChk = Server.CreateObject("ADODB.RecordSet")

						Do Until rsOT.EOF
						    delta_minute = CInt(rsOT("delta_minute")) ' 총경과시간을 총분으로 ..
                            rest_minute  = CInt(rsOT("rest_minute"))  ' 총휴게시간을 총분으로 ..

                            If delta_minute > rest_minute Then
                                delta_minute = delta_minute - rest_minute
                            Else
                                delta_minute = 0
                            End If

                            work_time = Fix(delta_minute / 60) ' 총작업시간을 시로 ..
                            work_minute = delta_minute Mod 60    ' 총작업시간을 시로 나눈몫인 분으로 ..

							If rsOT("cancel_yn") = "Y" Then
								cancel_yn = "취소"
							Else
								cancel_yn = "지급"
							End If

							If rsOT("acpt_no") = 0 Or rsOT("acpt_no") = null Then
								acpt_no = "없음"
							Else
								acpt_no = rsOT("acpt_no")
							End If

							If rsOT("you_yn") = "Y" Then
								you_view = "유상"
							Else
							 	you_view = "무상"
							End If
                            %>
                            <tr>
                                <td class="first"><%=rsOT("org_company")%></td>
                                <td><%=rsOT("org_bonbu")%></td>
                                <td><%=rsOT("org_saupbu")%></td>
                                <td><%=rsOT("org_team")%></td>
                                <td><%=rsOT("org_name")%></td>
                                <td><%=rsOT("reside_place")%></td>
                                <td><%=rsOT("mg_ce_id")%></td>
                                <td><%=rsOT("user_name")%>&nbsp;<%=rsOT("user_grade")%></td>

                                <td><%=rsOT("work_date")%>&nbsp;<%=rsOT("fromtime")%>:<%=rsOT("fromminute")%></td>
                                <td><%=rsOT("end_date")%>&nbsp;<%=rsOT("totime")%>:<%=rsOT("tominute")%></td>
                                <td><%=work_time%>시간 <%=work_minute%>분</td>
                                <td>
								<%
                                If rsOT("alter_timeoff_date") <> "" Then '사용자가 대체휴가시작일을 입력했을 경우
                                %>
                                    <%=rsOT("alter_timeoff_date")%>&nbsp;<%=rsOT("altertimeofftime")%>:<%=rsOT("altertimeoffminute")%>
                                    <br> ~
                                    <%
                                    If CInt(rsOT("alter_timeoff_minute_w")) > 0 Then ' 52시간 초과건을 경우
                                        dateNow = CDate(rsOT("work_date")) ' 일자변환
										week = Weekday(dateNow)	' 요일

										If week >= 4 Then
											mGap = (week - 4) * -1
										Else
											mGap = (6 - (3 - week)) * -1
										End If

										fDate = DateAdd("d", mGap, dateNow)
										lDate = DateAdd("d", mGap + 6, dateNow)

										objBuilder.Append "SELECT COUNT(*) AS last_cnt "
										objBuilder.Append "FROM overtime "
										objBuilder.Append "WHERE work_date BETWEEN '"&fDate&"' AND '"&lDate&"' "
										objBuilder.Append "	AND mg_ce_id  = '"&rsOT("mg_ce_id")&"' "
										objBuilder.Append "	AND LENGTH(alter_timeoff_date) > 0 "
										objBuilder.Append "	AND work_date > '"&rsOT("work_date")&"' "

										rsChk.Open objBuilder.ToString(), DBConn, 1
										objBuilder.Clear()

										last_cnt = 0

										If Not (rsChk.BOF Or rsChk.EOF) Then
											last_cnt = CInt(RsChk("last_cnt"))
										End If

										rsChk.Close()

										If last_cnt = 0 Then  ' 마지막 52시간 초과건을 경우
											Response.Write rsOT("alter_timeoff_enddate2") ' 주 52시간 초과 + (평일 22시 초과 + 휴일 8시간 초과)
										Else
											Response.Write rsOT("alter_timeoff_enddate1") ' (평일 22시 초과 + 휴일 8시간 초과)
										End If
                                    Else ' 52시간 초과건이 아닌 경우
										Response.write rsOT("alter_timeoff_enddate1") ' (평일 22시 초과 + 휴일 8시간 초과)
                                    End If
                                End If
                                %>
								</td>
								<td><%=acpt_no%></td>
								<td><%=rsOT("company")%></td>
								<td><%=rsOT("dept")%></td>
								<td><%=rsOT("cost_center")%></td>
								<td><%=rsOT("work_gubun")%></td>
								<td><%=rsOT("work_memo")%></td>
								<%
  								find = False

                                For i = 0 To UBound(allowerIDs)
                                    if  user_id = allowerIDs(i) then
                                        find =True
                                    end if
                                Next

                                If find = True Then
                                %>
								<td class="right"><%=formatnumber(rsOT("overtime_amt"),0)%></td>
								<%
                                End If
  							    %>
								<td><%=you_view%></td>
                                <td><%=cancel_yn%></td>
								<td><%=rsOT("allow_yn")%></td>
								<td><span name ="allowSayou"><%=rsOT("allow_sayou")%></span></td>
							</tr>
						    <%
							rsOT.MoveNext()
						Loop
						rsOT.Close() : Set rsOT = Nothing
						Set rsChk = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>