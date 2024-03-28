<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim from_date
Dim to_date

' 야특근 승인권자 ID 리스트
allowerIDs = Array("100125","100029","100015","100031","100020","100018") ' "강명석","이재원","전간수","최길성','홍건형','송지영'

treeDayAgo = DateAdd("d",-60,now())

view_c     = Request.form("view_c")
mg_ce      = Request.form("mg_ce")
from_date  = Request.form("from_date")
to_date    = Request.form("to_date")

work_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)

If view_c = "" Then
	view_c = "total"
End If

if from_date = "" then
    from_date = mid(work_month,1,4) + "-" + mid(work_month,5,2) + "-01"
end if
if to_date = "" then
    to_date = cstr(dateadd("d",-1, dateadd("m",1,datevalue(from_date)) ))
end if


Set Dbconn  = Server.CreateObject("ADODB.Connection")
Set Rs      = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 포지션별
posi_sql = " and overtime.mg_ce_id = '" + user_id + "'"

if position = "팀원" then
	view_condi = "본인"
end if

if position = "파트장" then
	if view_c = "total" then
		if org_name = "한화생명호남" then
			posi_sql = " and (overtime.org_name = '한화생명호남' or overtime.org_name = '한화생명전북') "
		  else
			posi_sql = " and overtime.org_name = '"&org_name&"'"
		end if
	  else
		if org_name = "한화생명호남" then
			posi_sql = " and (overtime.org_name = '한화생명호남' or overtime.org_name = '한화생명전북') and memb.user_name like '%"&mg_ce&"%'"
		  else
			posi_sql = " and overtime.org_name = '"&org_name&"' and memb.user_name like '%"&mg_ce&"%'"
		end if
	end if
end if

if position = "팀장" then
	if view_c = "total" then
		posi_sql = " and overtime.team = '"&team&"'"
	  else
		posi_sql = " and overtime.team = '"&team&"' and memb.user_name like '%"&mg_ce&"%'"
	end if
end if

if position = "사업부장" or cost_grade = "2" then
	if view_c = "total" then
        'posi_sql = " and overtime.saupbu = '"&saupbu&"'"
        posi_sql = " and overtime.saupbu = emp_master.emp_saupbu "&chr(13)
	  else
        'posi_sql = " and overtime.saupbu = '"&saupbu&"' and memb.user_name like '%"&mg_ce&"%'"
        posi_sql = " and overtime.saupbu = emp_master.emp_saupbu and memb.user_name like '%"&mg_ce&"%'"&chr(13)
	end if
end if

if position = "본부장" or cost_grade = "1" then
  	if view_c = "total" then
		posi_sql = " and overtime.bonbu = '"&bonbu&"'"
 	else
		posi_sql = " and overtime.bonbu = '"&bonbu&"' and memb.user_name like '%"&mg_ce&"%'"
	end if
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "전체"
  	if view_c = "total" then
		posi_sql = ""
 	else
		posi_sql = " and memb.user_name like '%"&mg_ce&"%'"
	end if
end if

base_sql = "    SELECT overtime.cancel_yn                            "&chr(13)&_
           "         , overtime.acpt_no                              "&chr(13)&_
           "         , overtime.you_yn                               "&chr(13)&_
           "         , overtime.org_name                             "&chr(13)&_
           "         , overtime.user_name                            "&chr(13)&_
           "         , overtime.user_grade                           "&chr(13)&_
           "         , overtime.mg_ce_id                             "&chr(13)&_
           "         , overtime.work_date                            "&chr(13)&_
           "         , overtime.company                              "&chr(13)&_
           "         , overtime.dept                                 "&chr(13)&_
           "         , overtime.work_gubun                           "&chr(13)&_
           "         , overtime.work_memo                            "&chr(13)&_
           "         , overtime.overtime_amt                         "&chr(13)&_
           "         , overtime.end_yn                               "&chr(13)&_
           "         , overtime.reg_id                               "&chr(13)&_
           "         , overtime.allow_yn                             "&chr(13)&_
           "         , ifnull(overtime.delta_minute,0) delta_minute  "&chr(13)&_
           "         , ifnull(overtime.rest_minute,0) rest_minute    "&chr(13)&_
           "         , memb.user_name                                "&chr(13)&_
           "         , memb.user_grade                               "&chr(13)&_
		   "         , emp_org_mst.org_name                          "&chr(13)&_
           "      FROM overtime                                      "&chr(13)&_
           "INNER JOIN memb                                          "&chr(13)&_
           "        ON overtime.mg_ce_id = memb.user_id              "&chr(13)&_
           "inner join emp_master                                    "&chr(13)&_
           "        ON emp_master.emp_no = overtime.mg_ce_id         "&chr(13)&_
		   "inner join emp_org_mst									 "&chr(13)&_
		   "        ON emp_org_mst.org_code = emp_master.emp_org_code        "&chr(13)
date_sql = "     WHERE work_date >= '" + from_date  + "' "&chr(13)&_
           "       AND work_date <= '" + to_date  + "'   "&chr(13)

sql = base_sql & date_sql & posi_sql & chr(13)&_
    " ORDER BY overtime.org_name, memb.user_name, work_date"


Rs.Open Sql, Dbconn, 1

title_line = "야특근 관리 (※ 소속직원 검색 안될 경우 인사팀에 요청하여 소속본부, 사업부 일치를 요청 하시기 바랍니다.) 작업일3일후 수정불가"
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

				if ((fDate != "") && (lDate != "") && (fDate > lDate)) {
					alert("검색 시작년월일이 종료 년월일 보다 작을 수 없습니다.");
					return false;
				}

				return true;
			}

			function condi_view()
            {
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
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="overtime_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건검색</dt>
						<dd>
							<p style="position:relative">

                                &nbsp;&nbsp;<strong>작업년월&nbsp;</strong>
                                <input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker1">
                                    ~
                                <input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker2">

                                <label><strong>조회권한 : </strong><%=view_grade%></label>
                                <label><strong>조회범위 : </strong>
                                <%
                                if position = "팀원" and cost_grade <> "0" then
                                        Response.write view_condi
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
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                                <span style="position:absolute;right:5px; cursor: pointer;" class="btnType04" onclick="pop_Window('overtime_stats.asp?','asview_pop','scrollbars=yes,width=1200,height=700')">주 52시간 현황보기</span>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" />
							<col width="7%" />
							<col width="7%" />
							<col width="11%" />
							<col width="5%" />
							<col width="11%" />
							<col width="11%" />
							<col width="13%" />
							<col width="13%" />
							<%
                            find = False
                            For i = 0 To uBound(allowerIDs)
                            if  user_id = allowerIDs(i) then
                                find =True
                            end if
                            Next

                            if find = True then
                                %><col width="7%" /><%
                            end if
							%>
							<col width="5%" />
							<col width="5%" />
							<col width="4%" />
							<col width="4%" />
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">조직명</th>
								<th scope="col">작업자</th>
								<th scope="col">근무일자</th>
								<th scope="col">총 시간</th>
								<th scope="col">AS NO</th>
								<th scope="col">회사</th>
								<th scope="col">조직명</th>
								<th scope="col">야특근구분</th>
								<th scope="col">작업내역</th>
								<%
  								find = False
                                For i = 0 To uBound(allowerIDs)
                                    if  user_id = allowerIDs(i) then
                                        find =True
                                    end if
                                Next

                                if find = True then
                                    %><th scope="col">신청금액</th><%
                                end if
  							    %>
								<th scope="col">유무상</th>
								<th scope="col">지급</th>
								<th scope="col">수정</th>
								<th scope="col">승인</th>
							</tr>
						</thead>
						<tbody>
						<%
                        cost_sum = 0
                        end_sum = 0
                        cancel_sum = 0

                        do until rs.eof

                            delta_minute = Cint( Rs("delta_minute") ) ' 총경과시간을 총분으로 .. (승인,미승인 관계없이 둘다)
                                rest_minute  = Cint( Rs("rest_minute") )  ' 총휴게시간을 총분으로 .. (승인,미승인 관계없이 둘다)
                                if (delta_minute > rest_minute) then
                                delta_minute = delta_minute - rest_minute
                            else
                                delta_minute = 0
                            end if
                            work_time   = Fix(delta_minute / 60) ' 총작업시간을 시로 ..  (승인,미승인 관계없이 둘다)
                            work_minute = delta_minute mod 60    ' 총작업시간을 시로 나눈몫인 분으로 ..  (승인,미승인 관계없이 둘다)

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

                            cost_sum = cost_sum + rs("overtime_amt")
                            if rs("cancel_yn") = "Y" then
                                cancel_sum = cancel_sum + rs("overtime_amt")
                            else
                                end_sum = end_sum + rs("overtime_amt")
                            end if
                            if rs("you_yn") = "Y" then
                                you_view = "유상"
                                else
                                you_view = "무상"
                            end if
                            %>
                            <tr>
                                <td class="first"><%=rs("org_name")%></td>
                                <td><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%><input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=rs("mg_ce_id")%>"></td>
                                <td><%=rs("work_date")%><input name="work_date" type="hidden" id="work_date" value="<%=rs("work_date")%>"></td>
                                <td><%=work_time%>시간 <%=work_minute%>분</td>
                                <td>
                                    <%
                                        if acpt_no = "없음" then
                                            Response.write acpt_no
                                        else
                                    %>
                                    <a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=acpt_no%>','asview_pop','scrollbars=yes,width=800,height=700')"><%=acpt_no%></a>
                                    <% end if	%>
                                </td>
                                <td><%=rs("company")%></td>
                                <td><%=rs("dept")%></td>
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
                                <td>
                                <%
                                if rs("end_yn") = "Y" then
                                    Response.write "마감"
                                else
                                    if treeDayAgo < rs("work_date") then
                                        if rs("mg_ce_id") = user_id or rs("reg_id") = user_id then ' ce와 등록자가 동일인인 경우..
                                            if rs("acpt_no") = 0 then ' AS접수 번호가 없다면 '한진'
                                                %><a href="#" onClick="pop_Window('overtime_hanjin_add.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime__hanjinadd_popup','scrollbars=yes,width=1100,height=500')">수정</a><%
                                            else
                                                if rs("work_date") > "2014-12-31" then
                                                    %><a href="#" onClick="pop_Window('overtime_as_mod_15.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>','overtime_as_mod_15_popup','scrollbars=yes,width=1000,height=660')">수정</a><%
                                                else
                                                    %><a href="#" onClick="pop_Window('overtime_add.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime_add_popup','scrollbars=yes,width=750,height=300')">수정</a><%
                                                end if
                                            end if
                                        else
                                            %><a href="#" onClick="pop_Window('overtime_cancel.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime_cancel_popup','scrollbars=yes,width=750,height=300')">수정</a><%
                                        end if
                                    else
                                        Response.write "불가" ' 3일 초과시 수정불가
                                    end if
                                end if
                                %>
                                </td>
                                <td><%=rs("allow_yn")%></td>
                                </tr>
                                <%
                                rs.movenext()
                            loop
                            rs.close()
                            %>
							<!-- 합계 항목이 건수로 계산되어 미노출 처리(인사에서 급여에 포함해서 지급 처리함)[허정호_20210705]
                            <tr>
								<th colspan="2" class="first">합 계</th>
                                <th colspan="3">신청금액 :&nbsp;<%''=formatnumber(cost_sum,0)%></th>
                                <th colspan="3">지급금액 :&nbsp;<%''=formatnumber(end_sum,0)%></th>
                                <%
  								'find = False
                                'For i = 0 To uBound(allowerIDs)
                                '    if  user_id = allowerIDs(i) then
                                '        find =True
                                '    end if
                                'Next

                                'if find = True then
                                '    width = 6
                                'else
                                '    width = 5
                                'end if
                                %>
                                <th colspan="<%''=width%>">취소금액 :&nbsp;<%''=formatnumber(cancel_sum,0)%></th>
						    </tr>-->
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                if intend > total_page then
                    intend = total_page
                end if
                %>

				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
				    <td width="15%">
					    <div class="btnCenter">
                        <a href="/cost/excel/overtime_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&view_c=<%=view_c%>&mg_ce=<%=mg_ce%>" class="btnType04">엑셀다운로드</a>
					    </div>
                    </td>
				    <td width="85%">
    					<div class="btnRight">
                            <!-- 2019.02.07 지현주 요청 오승택,부종민 "한진 제주공항","한진 제주지점","수도권사업부 의 경우 '한진당직스케쥴'로 디스플레이 하도록 수정 -->
                            <!-- 2019.02.08 정맹석 요청 버그수정 -->
                            <!-- [<%=cost_grade%>] [<%=saupbu%>] [<%=org_name%>] -->
                            <% if cost_grade = "0" or (saupbu <> "KAL지원사업부" and saupbu <> "공항지원사업부" and ((org_name<>"한진 제주공항" and org_name<>"한진 제주지점") or saupbu<>"수도권사업부")) then	%>
                                <a href="#" onClick="pop_Window('overtime_as_add_15.asp','overtime_as_add_15_popup','scrollbars=yes,width=1000,height=660')" class="btnType04">A/S연동 야특근등록</a>
                            <% end if	%>
                            <% if cost_grade = "0" or saupbu = "KAL지원사업부" or saupbu = "공항지원사업부" or org_name = "국회사무처" or org_name="한진 부산지점" or org_name="한진 테크센터" or org_name="한진 울산공항" or org_name="한진 포항공항" or org_name="한진 대전지점" or org_name="한진 광주지점" or org_name="한진 부산공항" or org_name="한진 대구공항" or (org_name="한진 청주공항" and saupbu="충청사업부") or ((org_name="한진 제주공항" or org_name="한진 제주지점" ) and saupbu="수도권사업부") Or org_name = "한진" then	%>
                                <a href="#" onClick="pop_Window('overtime_hanjin_add.asp','overtime_hanjin_as_add_popup','scrollbars=yes,width=1100,height=500')" class="btnType04"> 한진당직및스케줄등록</a>
                            <% end if	%>
    					</div>
                    </td>
			    </tr>
				</table>
			</form>
		</div>
	</div>
	</body>
</html>
