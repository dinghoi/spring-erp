<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim from_date
Dim to_date

work_month=Request.form("work_month")
view_c=Request.form("view_c")
mg_ce=Request.form("mg_ce")

If work_month = "" Then
	work_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	view_c = "total"
	mg_ce = ""
End If

from_date = mid(work_month,1,4) + "-" + mid(work_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
sign_month = work_month

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 포지션별
posi_sql = " and overtime.mg_ce_id = '" + user_id + "'"

if position = "팀원" then
	view_condi = "본인"
end if

'if position = "파트장" then
'	if view_c = "total" then
'		posi_sql = " and overtime.org_name = '"&org_name&"'"
'	  else
'		posi_sql = " and overtime.org_name = '"&org_name&"' and memb.user_name like '%"&mg_ce&"%'"
'	end if
'end if

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

if position = "사업부장" then
	if view_c = "total" then
		posi_sql = " and overtime.saupbu = '"&saupbu&"'"
	  else
		posi_sql = " and overtime.saupbu = '"&saupbu&"' and memb.user_name like '%"&mg_ce&"%'"
	end if
end if

if position = "본부장" then 
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

base_sql = "select overtime.* , memb.user_name, memb.user_grade  FROM overtime INNER JOIN memb ON overtime.mg_ce_id = memb.user_id "
date_sql = " where work_date >= '" + from_date  + "' and work_date <= '" + to_date  + "'"

sql = base_sql + date_sql + posi_sql + " order by overtime.org_name, user_name, work_date"
Rs.Open Sql, Dbconn, 1

title_line = "야특근 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				if (document.frm.work_month.value == "") {
					alert ("발생년월을 입력하세요.");
					return false;
				}	
				return true;
			}
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('mg_ce_view').style.display = 'none';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('mg_ce_view').style.display = '';
				}	
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
                            <p>
								<label>
								&nbsp;&nbsp;<strong>작업년월&nbsp;</strong>(예201401) : 
                                	<input name="work_month" type="text" value="<%=work_month%>" style="width:70px">
								</label>
								<label>
								<strong>조회권한 : </strong><%=view_grade%>
								</label>
								<label>
								<strong>조회범위 : </strong>
							<% if position = "팀원" and cost_grade <> "0" then %>
								<%=view_condi%>
							<%   else	%>
                              	<input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                조직전체
                                <input type="radio" name="view_c" value="reg_id" <% if view_c = "reg_id" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                개인별
							<% end if %>
								</label>
								<label>
                                	<input name="mg_ce" type="text" value="<%=mg_ce%>" style="width:70px; display:none" id="mg_ce_view">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="7%" >
							<col width="7%" >
							<col width="5%" >
							<col width="11%" >
							<col width="11%" >
							<col width="13%" >
							<col width="*" >
							<col width="7%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">조직명</th>
								<th scope="col">작업자</th>
								<th scope="col">근무일자</th>
								<th scope="col">AS NO</th>
								<th scope="col">회사</th>
								<th scope="col">조직명</th>
								<th scope="col">야특근구분</th>
								<th scope="col">작업내역</th>
								<th scope="col">신청금액</th>
								<th scope="col">유무상</th>
								<th scope="col">지급</th>
								<th scope="col">수정</th>
							</tr>
						</thead>
						<tbody>
						<%
						cost_sum = 0
						end_sum = 0
						cancel_sum = 0
						do until rs.eof

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
								<td>
						<% if acpt_no = "없음" then	%>
								<%=acpt_no%>
						<%   else	%>
                        		<a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=acpt_no%>','asview_pop','scrollbars=yes,width=800,height=700')"><%=acpt_no%></a>
                        <% end if	%>
                                </td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("work_gubun")%></td>
								<td><%=rs("work_memo")%></td>
								<td class="right"><%=formatnumber(rs("overtime_amt"),0)%></td>
								<td><%=you_view%></td>
								<td><%=cancel_yn%></td>
								<td>
						<% if rs("end_yn") = "Y" then	%>
                                마감
                        	<%   else	%>
							<% if rs("mg_ce_id") = user_id or rs("reg_id") = user_id then	%>
							<%   if rs("acpt_no") = 0 then	%>
                                <a href="#" onClick="pop_Window('overtime_hanjin_add.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime__hanjinadd_popup','scrollbars=yes,width=900,height=350')">수정</a>
							<%     else	%>
							<%       if rs("work_date") > "2014-12-31" then	%>
                                <a href="#" onClick="pop_Window('overtime_as_mod_15.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>','overtime_as_mod_15_popup','scrollbars=yes,width=750,height=330')">수정</a>
                            <%	  		else	%>
                                <a href="#" onClick="pop_Window('overtime_add.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime_add_popup','scrollbars=yes,width=750,height=300')">수정</a>
							<%		  end if	%>
							<%   end if	%>	
							<%   else	%>
                                <a href="#" onClick="pop_Window('overtime_cancel.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime_cancel_popup','scrollbars=yes,width=750,height=300')">수정</a>
							<% end if	%>
						<% end if	%>
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr>
								<th colspan="2" class="first">합 계</th>
							  <th colspan="3">신청금액 :&nbsp;<%=formatnumber(cost_sum,0)%></th>
							  <th colspan="3">지급금액 :&nbsp;<%=formatnumber(end_sum,0)%></th>
							  <th colspan="4">취소금액 :&nbsp;<%=formatnumber(cancel_sum,0)%></th>
						    </tr>
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
                    <a href="overtime_excel.asp?work_month=<%=work_month%>&view_c=<%=view_c%>&mg_ce=<%=mg_ce%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td width="85%">
					<div class="btnRight">
				<% if cost_grade = "0" or (saupbu <> "KAL지원사업부" and saupbu <> "공항지원사업부") then	%>
                    <a href="#" onClick="pop_Window('overtime_as_add.asp','overtime_as_add_popup','scrollbars=yes,width=900,height=660')" class="btnType04">2014년까지 A/S연동 야특근등록</a>
                    <a href="#" onClick="pop_Window('overtime_as_add_15.asp','overtime_as_add_15_popup','scrollbars=yes,width=900,height=660')" class="btnType04">2015년 A/S연동 야특근등록</a>
				<% end if	%>
				<% if cost_grade = "0" or saupbu = "KAL지원사업부" or saupbu = "공항지원사업부" then	%>
                    <a href="#" onClick="pop_Window('overtime_hanjin_add.asp','overtime_hanjin_as_add_popup','scrollbars=yes,width=900,height=300')" class="btnType04"> 한진당직및스케줄등록</a>
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

