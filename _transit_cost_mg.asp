<!--#include virtual="/common/inc_top.asp" -->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"--><!--nkpmg_user.asp 변수 선언-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" --><!--사용자 정의 함수-->
<%
'===================================================
'DB Connection
'===================================================
Dim DBConn, RsCount, rs_sum, tranRs

Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'StringBuilder Object
'===================================================
Dim objBuilder

'StringBuffer 형식 사용[허정호_20201123]
Set objBuilder = New StringBuilder

'===================================================
'Request & Param
'===================================================
Dim ck_sw
Dim Page, pgsize, start_page, stpage
Dim run_month, transit_id, view_c, view_d, use_man
Dim from_date, end_date, to_date, sign_month
Dim posi_sql
Dim view_condi
Dim view_grade, transit_sql, base_sql, date_sql, order_sql
Dim total_record, total_page
Dim sum_far, sum_oil_price, sum_fare, sum_repair_cost, sum_parking, sum_toll
Dim title_line
Dim car_gubun, run_km, cancel_view, start_view, end_view
Dim chk_slip_month, chk_reg_month, chk_reg_day, bgcolor
Dim intstart, intend, first_page, i

ck_sw = Request("ck_sw") '페이징 구분값
Page = Request("page")

If ck_sw = "y" Then
	run_month = Request("run_month")
	transit_id = Request("transit_id")
	view_c = Request("view_c")
	view_d = Request("view_d")
	use_man = Request("use_man")
Else
	run_month = Request.Form("run_month")
	transit_id = Request.Form("transit_id")
	view_c = Request.Form("view_c")
	view_d = Request.Form("view_d")
	use_man = Request.Form("use_man")
End If

'toString 함수 적용[허정호_20201123]
'If view_d = "" Then
If toString(view_d, "") = "" Then
    view_d = "run"
End If

If run_month = "" Then
	run_month = Mid(CStr(Now()), 1, 4) + Mid(CStr(Now()), 6, 2)
	transit_id = "차량"
    view_c = "total"
    view_d = "run"
	use_man = ""
End If

from_date = Mid(run_month, 1, 4) + "-" + Mid(run_month, 5, 2) + "-01"

'중복 사용으로 주석 처리[허정호_20201123]
end_date = DateAdd("m", 1, from_date)

to_date = CStr(DateAdd("d", -1, end_date))
sign_month = run_month

pgsize = 10 ' 화면 한 페이지

If toString(Page, "") = "" Then
	Page = 1
	start_page = 1
End If
stpage = Int((page - 1) * pgsize)

'전체 쿼리 수정[허정호_20201123]
' 포지션별
posi_sql = "AND tc.mg_ce_id = '" + user_id + "'"

If position = "팀원" Then
	view_condi = "본인"
End If

If position = "파트장" Then
	If view_c = "total" Then
		If org_name = "한화생명호남" Then
			posi_sql = "AND (tc.org_name = '한화생명호남' OR tc.org_name = '한화생명전북') "
		Else
			posi_sql = "AND tc.org_name = '"&org_name&"' "
		End If
	Else
		If org_name = "한화생명호남" Then
			posi_sql = "AND (tc.org_name = '한화생명호남' OR tc.org_name = '한화생명전북') AND mb.user_name LIKE '%"&use_man&"%' "
		Else
			posi_sql = "AND tc.org_name = '"&org_name&"' AND mb.user_name LIKE '%"&use_man&"%' "
		End If
	End If
End If

If position = "팀장" Then
	If view_c = "total" Then
		posi_sql = "AND tc.team = '"&team&"' "
	Else
		posi_sql = "AND tc.team =  '"&team&"' AND mb.user_name LIKE '%"&use_man&"%' "
	End If
End If

If position = "사업부장" Or cost_grade = "2" Then
    If view_c = "total" Then
		posi_sql = "AND tc.saupbu = em.emp_saupbu "
    Else
		posi_sql = "AND tc.saupbu = em.emp_saupbu AND mb.user_name LIKE '%"&use_man&"%' "
    End If
End If

If position = "본부장" Or cost_grade = "1" Then
  	If view_c = "total" Then
		posi_sql = "AND tc.bonbu = '"&bonbu&"' "
 	Else
		posi_sql = "AND tc.bonbu = '"&bonbu&"' AND mb.user_name LIKE '%"&use_man&"%' "
	End If
End If

view_grade = position

If cost_grade = "0" Then
	view_grade = "전체"

  	If view_c = "total" Then
		posi_sql = ""
 	Else
		posi_sql = "AND mb.user_name LIKE '%"&use_man&"%' "
	End If
End If

If transit_id = "차량" Then
	transit_sql = "AND (tc.car_owner = '개인' OR tc.car_owner = '회사') "
Else
	transit_sql = "AND (tc.car_owner = '대중교통') "
End If

If view_d = "run" Then
   date_sql = "WHERE (tc.run_date >= '" + from_date  + "' AND tc.run_date <= '" + to_date  + "') "
	order_sql = "ORDER BY mb.user_name ASC, tc.run_date DESC, tc.run_seq DESC "
End If

If view_d = "reg" Then
	date_sql = "WHERE (tc.reg_date >= '" + from_date  + " 00:00:00' AND tc.reg_date <='" + to_date  + " 23:59:59') "
	order_sql = "ORDER BY mb.user_name ASC, tc.reg_date DESC, tc.run_seq DESC "
End If

'차량운행정보 개수 조회
objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM transit_cost AS tc "
objBuilder.Append "INNER JOIN memb AS mb ON tc.mg_ce_id = mb.user_id "
objBuilder.Append "INNER JOIN emp_master AS em ON mb.user_id = em.emp_no "
objBuilder.Append date_sql & posi_sql & transit_sql

Set RsCount = DBConn.Execute(objBuilder.ToString())

total_record = CInt(RsCount(0)) 'Result.RecordCount

'레코드 객체 제거[허정호_20201123]
objBuilder.Clear()
RsCount.Close()
Set RsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

objBuilder.Append "SELECT SUM(far) AS far, "
objBuilder.Append "SUM(oil_price) AS oil_price, "
objBuilder.Append "SUM(fare) AS fare, "
objBuilder.Append "SUM(repair_cost) AS repair_cost, "
objBuilder.Append "SUM(parking) AS parking, "
objBuilder.Append "SUM(toll) AS toll "
objBuilder.Append "FROM transit_cost AS tc "
objBuilder.Append "INNER JOIN  memb AS mb ON tc.mg_ce_id = mb.user_id "
objBuilder.Append "INNER JOIN emp_master AS em ON mb.user_id = em.emp_no "
objBuilder.Append date_sql & posi_sql & transit_sql
objBuilder.Append "AND cancel_yn = 'N' "

Set rs_sum = DBConn.Execute (objBuilder.ToString())

If rs_sum("far") = "" Or IsNull(rs_sum("far")) Then
	sum_far         = 0
	sum_oil_price   = 0
	sum_fare        = 0
	sum_repair_cost = 0
	sum_parking     = 0
	sum_toll        = 0
Else
	sum_far         = rs_sum("far")
	sum_oil_price   = rs_sum("oil_price")
	sum_fare        = rs_sum("fare")
	sum_repair_cost = rs_sum("repair_cost")
	sum_parking     = rs_sum("parking")
	sum_toll        = rs_sum("toll")
End If

objBuilder.Clear()

rs_sum.Close()
Set rs_sum = Nothing

' 조건별 조회.........
' 차량운행정보 조회 리스트
objBuilder.Append "SELECT tc.run_date, tc.mg_ce_id, tc.run_seq, tc.user_name, tc.car_owner, "
objBuilder.Append "tc.oil_kind, tc.start_company, tc.start_point, tc.start_km, tc.end_company, "
objBuilder.Append "tc.end_point, tc.end_km, tc.far, tc.transit, tc.oil_price, "
objBuilder.Append "tc.fare, tc.run_memo, tc.repair_cost, tc.parking, tc.toll, "
objBuilder.Append "tc.cancel_yn, tc.end_yn, tc.reg_date "
objBuilder.Append "FROM transit_cost AS tc "
objBuilder.Append "INNER JOIN memb AS mb ON tc.mg_ce_id = mb.user_id "
objBuilder.Append "INNER JOIN emp_master AS em ON mb.user_id = em.emp_no "
objBuilder.Append date_sql & posi_sql & transit_sql & order_sql & " "
objBuilder.Append "LIMIT " & stpage & ", "  & pgsize

Response.write objBuilder.tostring()

Set tranRs = Server.CreateObject("ADODB.Recordset")
tranRs.Open objBuilder.ToString(), DBConn, 1

title_line = "교통비 관리"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.run_month.value == "") {
					alert ("운행년월을 입력하세요");
					return false;
				}
				return true;
			}

			function condi_view() {
                <%If position <> "팀원" Or cost_grade = "0" Then %>
                    if (eval("document.frm.view_c[0].checked")) {
                        document.getElementById('use_man_view').style.display = 'none';
                    }
                    if (eval("document.frm.view_c[1].checked")) {
                        document.getElementById('use_man_view').style.display = '';
                    }
                <%End If %>
			}
		</script>
	</head>
	<body onLoad="condi_view()">
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="transit_cost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건검색</dt>
                        <dd>
                            <p>
								<label>
                                    <input type="radio" name="view_d" value="run" <% If view_d = "run" Then %>checked <% End If %> style="width:20px">
                                    <strong>운행년월&nbsp;</strong>
                                    <input type="radio" name="view_d" value="reg" <% If view_d = "reg" Then %>checked <% End If %> style="width:20px">
                                    <strong>발급년월&nbsp;</strong>

                                    : <input name="run_month" type="text" value="<%=run_month%>" style="width:70px">
                                    (예201401)
								</label>
								<label>
                              	<input type="radio" name="transit_id" value="차량" <% If transit_id = "차량" Then %>checked <% End If %> style="width:20px">
                                차량운행일지
                                <input type="radio" name="transit_id" value="대중" <% If transit_id = "대중" Then %>checked <% End If %> style="width:20px">
                                대중교통비
								</label>
								<label><strong>조회권한:</strong><%=view_grade%></label>
								<label>
								<strong>조회범위:</strong>
                                <% If position = "팀원" and cost_grade <> "0" Then %>
                                    <%=view_condi%>
                                <% Else	%>
                                    <input type="radio" name="view_c" value="total" <% If view_c = "total" Then %>checked <% End If %> style="width:20px" onClick="condi_view()">
                                    조직전체
                                    <input type="radio" name="view_c" value="reg_id" <% If view_c = "reg_id" Then %>checked <% End If %> style="width:20px" onClick="condi_view()">
                                    개인별
								<% End If	%>
                                </label>
								<label>
                                	<input name="use_man" type="text" value="<%=use_man%>" style="width:70px; display:none" id="use_man_view">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="6%" >
							<col width="6%" >
							<col width="4%" >
							<col width="17%" >
							<col width="17%" >
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="3%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="4%" >
							<col width="3%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">운행자</th>
                                <th rowspan="2" scope="col">운행일자</th>
                                <th rowspan="2" scope="col">발급일자</th>
								<th rowspan="2" scope="col">차량<br>구분</th>
								<th rowspan="2" scope="col">출발지</th>
								<th rowspan="2" scope="col">도착지</th>
								<th rowspan="2" scope="col">운행목적</th>
								<th rowspan="2" scope="col">시작KM</th>
								<th rowspan="2" scope="col">도착KM</th>
								<th rowspan="2" scope="col">거리</th>
								<th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">경 비 </th>
								<th rowspan="2" scope="col">지급</th>
								<th rowspan="2" scope="col">수정</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">수리비</th>
								<th scope="col">대중교통</th>
								<th scope="col">주유금액</th>
								<th scope="col">주차료</th>
								<th scope="col">통행료</th>
							</tr>
						</thead>
						<tbody>
						<%

						Do Until tranRs.EOF
							If tranRs("car_owner") = "대중교통" Then
								car_gubun = tranRs("transit")
							Else
								car_gubun = tranRs("car_owner") + "<br>" + tranRs("oil_kind")
							End If
							run_km = tranRs("far")

							If tranRs("cancel_yn") = "Y" Then
								cancel_view = "취소"
							Else
							  	cancel_view = "지급"
							End If

							If tranRs("start_km") = "" Or IsNull(tranRs("start_km")) Then
								start_view = 0
							Else
							  	start_view = tranRs("start_km")
							End If
							If tranRs("end_km") = "" Or IsNull(tranRs("end_km")) Then
								end_view = 0
							Else
							  	end_view = tranRs("end_km")
							End If

                            ' 5일 이후 지연 입력건 검출...
                            chk_slip_month = Mid(tranRs("run_date"), 1, 7)
                            chk_reg_month = Mid(tranRs("reg_date"), 1, 7)
                            chk_reg_day = Mid(tranRs("reg_date"), 9, 2)

                            If ((chk_slip_month < chk_reg_month) And (chk_reg_day > "05")) Then
                                bgcolor = "burlywood"
                            Else
                                bgcolor = "#f8f8f8"
                            End If
                            %>
                            <tr style="background-color: <%=bgcolor%>;">
								<td class="first"><%=tranRs("user_name")%></td>
                                <td><%=tranRs("run_date")%></td>
                                <td><%=Mid(tranRs("reg_date"), 1, 10)%></td>
								<td><%=car_gubun%></td>
								<td><%=tranRs("start_company")%>-<%=tranRs("start_point")%></td>
								<td><%=tranRs("end_company")%>-<%=tranRs("end_point")%></td>
								<td><%=tranRs("run_memo")%>&nbsp;</td>
								<td class="right"><%=FormatNumber(start_view, 0)%></td>
								<td class="right"><%=FormatNumber(end_view, 0)%></td>
								<td class="right"><%=FormatNumber(run_km, 0)%></td>
								<td class="right"><%=FormatNumber(tranRs("repair_cost"), 0)%></td>
								<td class="right"><%=FormatNumber(tranRs("fare"), 0)%></td>
								<td class="right"><%=FormatNumber(tranRs("oil_price"), 0)%></td>
								<td class="right"><%=FormatNumber(tranRs("parking"), 0)%></td>
								<td class="right"><%=FormatNumber(tranRs("toll"), 0)%></td>
								<td><%=cancel_view%></td>
								<td>
                                    <%If tranRs("end_yn") <> "Y" Then %>
                                        <%If tranRs("car_owner") = "대중교통" Then %>
                                            <%If tranRs("mg_ce_id") = user_id Then%>
                                                <a href="#" onClick="pop_Window('mass_transit_add.asp?mg_ce_id=<%=tranRs("mg_ce_id")%>&run_date=<%=tranRs("run_date")%>&run_seq=<%=tranRs("run_seq")%>&u_type=<%="U"%>','mass_transit_add_pop','scrollbars=yes,width=850,height=350')">수정</a>
                                            <%Else%>
                                                <a href="#" onClick="pop_Window('mass_transit_cancel.asp?mg_ce_id=<%=tranRs("mg_ce_id")%>&run_date=<%=tranRs("run_date")%>&run_seq=<%=tranRs("run_seq")%>&u_type=<%="U"%>','mass_transit_cancel_pop','scrollbars=yes,width=850,height=350')">수정</a>
                                            <%End If%>
                                        <%Else%>
                                            <%If tranRs("mg_ce_id") = user_id Then%>
                                                <a href="#" onClick="pop_Window('car_drive_add.asp?mg_ce_id=<%=tranRs("mg_ce_id")%>&run_date=<%=tranRs("run_date")%>&run_seq=<%=tranRs("run_seq")%>&u_type=<%="U"%>','car_drive_add_pop','scrollbars=yes,width=900,height=500')">수정</a>
                                            <%Else%>
                                                <a href="#" onClick="pop_Window('car_drive_cancel.asp?mg_ce_id=<%=tranRs("mg_ce_id")%>&run_date=<%=tranRs("run_date")%>&run_seq=<%=tranRs("run_seq")%>&u_type=<%="U"%>','car_drive_cancel_pop','scrollbars=yes,width=900,height=470')">수정</a>
                                            <%End If%>
                                        <%End If%>
                                    <%Else%>
                                        마감
                                    <%End If%>
                                </td>
							</tr>
						    <%
							tranRs.MoveNext()
						Loop

						objBuilder.Clear()
						tranRs.close()
						Set tranRs = Nothing

						DBConn.Close()
						Set DBConn = Nothing

						If total_record <> 0 Then
						%>
							<tr>
								<th class="first">계</th>
								<th colspan="3"><%=total_record%>&nbsp;건</th>
								<th colspan="13">
									주행거리 :&nbsp;<%=FormatNumber(sum_far, 0)%>&nbsp;KM&nbsp;&nbsp;,
									&nbsp;&nbsp;수리비 :&nbsp;<%=FormatNumber(sum_repair_cost, 0)%>&nbsp;&nbsp;,
									&nbsp;&nbsp;대중교통비 :&nbsp;<%=FormatNumber(sum_fare, 0)%>&nbsp;&nbsp;,
									&nbsp;&nbsp;주유금액 :&nbsp;<%=FormatNumber(sum_oil_price, 0)%>&nbsp;&nbsp;,
									&nbsp;&nbsp;주차비 :&nbsp;<%=FormatNumber(sum_parking, 0)%>&nbsp;&nbsp;,
									&nbsp;&nbsp;통행료 :&nbsp;<%=FormatNumber(sum_toll, 0)%>
								</th>
							</tr>
						<%
						End If
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (Int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                If intend > total_page Then
                    intend = total_page
                End If
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
                    <div class="btnCenter">
                        <a href="transit_cost_excel.asp?run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>" class="btnType04">엑셀다운로드</a>
                    </div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="transit_cost_mg.asp?page=<%=first_page%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[처음]</a>
                        <%If intstart > 1 Then %>
                            <a href="transit_cost_mg.asp?page=<%=intstart -1%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[이전]</a>
                        <%End If %>
                        <%For i = intstart To intend %>
                            <%If i = Int(page) Then %>
                                <b>[<%=i%>]</b>
                            <%Else %>
                                <a href="transit_cost_mg.asp?page=<%=i%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <%End If %>
                        <%Next %>
                        <%If intend < total_page Then%>
                            <a href="transit_cost_mg.asp?page=<%=intend+1%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[다음]</a>
                            <a href="transit_cost_mg.asp?page=<%=total_page%>&run_month=<%=run_month%>&view_c=<%=view_c%>&view_d=<%=view_d%>&use_man=<%=use_man%>&transit_id=<%=transit_id%>&ck_sw=<%="y"%>">[마지막]</a>
                            <%Else %>
                            [다음]&nbsp;[마지막]
                        <%End If %>
                    </div>
                    </td>
				    <td width="25%">
                    <div class="btnCenter">
                        <a href="#" onClick="pop_Window('car_drive_add.asp','car_drive_add_popup','scrollbars=yes,width=900,height=450')" class="btnType04">차량운행일지입력</a>
                        <a href="#" onClick="pop_Window('mass_transit_add.asp','mass_transit_add_popup','scrollbars=yes,width=850,height=300')" class="btnType04">대중교통비입력</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>
