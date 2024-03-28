<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
Dim strNowWeek

work_date = Request.form("work_date")
work_item = Request.form("work_item")
view_c = Request.form("view_c")
acpt_no = Request.form("acpt_no")

'//2017-10-31 야특근 검색(srchTy:ce)
srchTy = Request.form("srchTy")
If Trim(srchTy&"")="" Then srchTy = Request.QueryString("srchTy")

if work_date = "" then
	work_date = mid(cstr(now()),1,10)
	work_item = ""
	acpt_no = ""
	view_c = "acpt"
end if

'//2017-10-31 담당CE만 야특근비 등록 가능하게 쿼리 변경
if view_c = "acpt" Then
		
		If srchTy = "ce" Then
				SQL = "select * from as_acpt where overtime <> 'Y' and ( acpt_no = '"&acpt_no&"') and as_process = '완료' and mg_ce_id='" & user_id & "' ORDER BY acpt_no ASC"
		Else
				SQL = "select * from as_acpt where overtime <> 'Y' and ( acpt_no = '"&acpt_no&"') and as_process = '완료' ORDER BY acpt_no ASC"			
		End If
Else
		
		If srchTy = "ce" Then
				SQL = "select * from as_acpt where overtime <> 'Y' and ( as_type = '"&work_item&"') and visit_date = '"&work_date&"' and as_process = '완료' and mg_ce_id='" & user_id & "' ORDER BY acpt_no ASC"
		Else
				SQL = "select * from as_acpt where overtime <> 'Y' and ( as_type = '"&work_item&"') and visit_date = '"&work_date&"' and as_process = '완료' ORDER BY acpt_no ASC"				
		End If
End If

Rs.open SQL, Dbconn, 1
'Response.write SQL

title_line = "A/S 검색"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>CE 검색</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=work_date%>" );
			});	  
			function as_list(acpt_no,company,dept,work_date,work_item,dev_inst_cnt,ran_cnt,work_man_cnt,week)
			{
				opener.document.frm.acpt_no.value = acpt_no;
				opener.document.frm.company.value = company;
				opener.document.frm.dept.value = dept;
				opener.document.frm.work_date1.readOnly = false;
				opener.document.frm.work_date2.readOnly = false;
				opener.document.frm.work_date1.value = work_date.substring(0, 10);
				opener.document.frm.week1.value = week;				
				opener.document.frm.work_date2.value = work_date.substring(0, 10);
				opener.document.frm.week2.value = week;				
				opener.document.frm.work_item.value = work_item;
				opener.document.frm.dev_inst_cnt.value = dev_inst_cnt;
				opener.document.frm.ran_cnt.value = ran_cnt;
				opener.document.frm.work_man_cnt.value = work_man_cnt;
				
				// 19시 부터시작,종료 시간을 셋팅한다.					
				opener.document.frm.from_hh.value = 19;				
				opener.document.frm.to_hh.value = 19;				
				
				window.close();
			}
			function frmcheck () {
				//if (formcheck(document.frm) && chkfrm()) {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
//				if(document.frm.work_item.value =="") {
//					alert('작업내용을 선택하세요');
//					frm.work_item.focus();
//					return false;}
				if(document.frm.end_date.value >= document.frm.work_date.value) {
					alert('마감된 날자입니다');
					frm.work_date.focus();
					return false;}
				{
					return true;
				}
			}
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('work1').style.display = 'none';
					document.getElementById('work2').style.display = 'none';
					document.getElementById('acpt1').style.display = '';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('work1').style.display = '';
					document.getElementById('work2').style.display = '';
					document.getElementById('acpt1').style.display = 'none';
				}	
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_search.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
 								<label><input type="radio" name="view_c" value="acpt" <% if view_c = "acpt" then %>checked<% end if %> style="width:25px" onClick="condi_view()">접수번호</label>
                                <label><input type="radio" name="view_c" value="work" <% if view_c = "work" then %>checked<% end if %> style="width:25px" onClick="condi_view()">작업내용</label>
                                <label id="work1">
								<strong>작업내용</strong>
                                <select name="work_item" id="work_item" style="width:100px">
                              		<option value="">선택</option>
								    <option value="방문처리" <%If work_item = "방문처리" then %>selected<% end if %>>방문처리</option>
								    <option value="원격처리" <%If work_item = "원격처리" then %>selected<% end if %>>원격처리</option>
								    <option value="신규설치" <%If work_item = "신규설치" then %>selected<% end if %>>신규설치</option>
								    <option value="신규설치공사" <%If work_item = "신규설치공사" then %>selected<% end if %>>신규설치공사</option>
								    <option value="이전설치" <%If work_item = "이전설치" then %>selected<% end if %>>이전설치</option>
								    <option value="이전설치공사" <%If work_item = "이전설치공사" then %>selected<% end if %>>이전설치공사</option>
								    <option value="랜공사" <%If work_item = "랜공사" then %>selected<% end if %>>랜공사</option>
								    <option value="이전랜공사" <%If work_item = "이전랜공사" then %>selected<% end if %>>이전랜공사</option>
								    <option value="장비회수" <%If work_item = "장비회수" then %>selected<% end if %>>장비회수</option>
								    <option value="예방점검" <%If work_item = "예방점검" then %>selected<% end if %>>예방점검</option>
								    <option value="야특근" <%If work_item = "야특근" then %>selected<% end if %>>야특근</option>
								    <option value="기타" <%If work_item = "기타" then %>selected<% end if %>>기타</option>
                                </select>
								</label>
								<label id="work2">
								<strong>작업일자</strong>
                                	<input name="work_date" type="text" value="<%=work_date%>" style="width:70px" id="datepicker" id="work_date">
								</label>
								<label id="acpt1">
								<strong>접수번호</strong>
                                	<input name="acpt_no" type="text" value="<%=acpt_no%>" style="width:70px" id="acpt_no">
								</label>
								<strong>마감:</strong><%=end_date%>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="12%" >
							<col width="22%" >
							<col width="*" >
							<col width="20%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">접수번호</th>
								<th scope="col">회사</th>
								<th scope="col">조직명</th>
								<th scope="col">담당자</th>
								<th scope="col">설치</th>
								<th scope="col">공사</th>
								<th scope="col">인원</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
							strNowWeek = WeekDay(rs("visit_date"))
							Select Case (strNowWeek)
							   Case 1	week = "일"
							   Case 2	week = "월"
							   Case 3	week = "화"
							   Case 4	week = "수"
							   Case 5	week = "목"
							   Case 6	week = "금"
							   Case 7	week = "토"
							End Select							
						%>
							<tr>
								<td class="first">
									<a href="#" onClick="as_list('<%=rs("acpt_no")%>','<%=rs("company")%>','<%=rs("dept")%>','<%=rs("visit_date")%>','<%=rs("as_type")%>','<%=rs("dev_inst_cnt")%>','<%=rs("ran_cnt")%>','<%=rs("work_man_cnt")%>','<%=week%>');"><%=rs("acpt_no")%></a>
								</td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("mg_ce")%>&nbsp;(<%=rs("mg_ce_id")%>)</td>
								<td><%=rs("dev_inst_cnt")%></td>
								<td><%=rs("ran_cnt")%></td>
								<td><%=rs("work_man_cnt")%></td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 then
							%>
							<tr>
								<td class="first" colspan="7">내역이 없습니다</td>
							</tr>
            				<%
						end if
						%>
						</tbody>
					</table>
				</div>				
				<input type="hidden" name="end_date" value="<%=end_date%>" ID="end_date">
				<input type="hidden" name="srchTy" value="<%=srchTy%>" ID="srchTy">
			</form>
		</div>        				
	</body>
</html>
