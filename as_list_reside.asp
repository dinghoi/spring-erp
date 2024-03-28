<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim as_process
Dim field_check
Dim field_view
Dim win_sw
dim company_tab(160)

win_sw = "close"
be_pg = "as_list_reside.asp"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	date_sw=Request("date_sw")
	process_sw=Request("process_sw")
	field_check=Request("field_check")
	field_view=Request("field_view")
  else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	date_sw=Request.form("date_sw")
	process_sw=Request.form("process_sw")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	company=Request.form("company")
End if

If to_date = "" or from_date = "" Then
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-15),1,10)
	field_check = "total"
	date_sw = "acpt"
	process_sw = "N"
End If
company = reside_company

If field_check = "total" Then
	field_view = ""
End If

if field_check = "acpt_no" then
'	if field_view > "9999999" or field_view < "0" then
'	end if
end if	
			
pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

' 조건별 조회.........
' 날짜별 조회(1)
base_sql = "select acpt_no,acpt_man,as_type,acpt_date,as_process,acpt_user,as_memo,company,dept,tel_ddd,tel_no1,tel_no2,sido,gugun,request_date,visit_date,mg_ce,asets_no, large_paper_no from as_acpt "

if date_sw = "acpt" then
'	date_sql = "where (CAST(acpt_date as date) >= '" + from_date  + "' and CAST(acpt_date as date) <= '" + to_date  + "') and (mg_group ='" + mg_group + "')"
	date_sql = "where (CAST(acpt_date as date) >= '" + from_date  + "' and CAST(acpt_date as date) <= '" + to_date  + "') "
  else
'	date_sql = "where (visit_date >= '" + from_date  + "' and visit_date <= '" + to_date  + "') and (mg_group ='" + mg_group + "')"
	date_sql = "where (visit_date >= '" + from_date  + "' and visit_date <= '" + to_date  + "') "
end if

if process_sw = "Y" then
	process_sql = " and ( as_process = '완료' or as_process = '대체' or as_process = '취소' ) "
  else
	process_sql = " and ( as_process = '접수' or as_process = '연기' or as_process = '입고' or as_process = '대체입고' ) "
end if

if field_check <> "total" then
	if field_check = "asets_no" then
		field_sql = " and ( " + field_check + " = '" + field_view + "' ) "
	  elseif field_check = "acpt_no" then
		field_sql = " and ( " + field_check + " = '"&field_view&"' ) "
	  else			
		field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
	end if
  else
  	field_sql = " "
end if
order_sql = " ORDER BY acpt_date DESC"

'sql = base_sql + date_sql + process_sql + field_sql + order_sql

if company = "전체" then
	com_sql = " "
  else
  	com_sql = " and (company = '" + company + "') "
end if

Sql = "SELECT count(*) FROM as_acpt " + date_sql + com_sql + process_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + date_sql + com_sql + process_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "A/S 총괄 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.field_check.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/as_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_list_reside.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
                                <input name="process_sw" type="radio" value="N"  <% if process_sw = "N" then %>checked<% end if %> style="width:25px">미처리
                                <input name="process_sw" type="radio" value="Y"  <% if process_sw = "Y" then %>checked<% end if %> style="width:25px">처리완료

                              	<input type="radio" name="date_sw" value="acpt" <% if date_sw = "acpt" then %>checked<% end if %> style="width:25px">접수일
                              	<input type="radio" name="date_sw" value="visit" <% if date_sw = "visit" then %>checked<% end if %> style="width:25px">완료일
								<label>
								&nbsp;<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <label>
								<strong>조건</strong>
                                <select name="field_check" id="field_check" style="width:80px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                    <option value="acpt_no" <% if field_check = "acpt_no" then %>selected<% end if %>>접수번호</option>
                                    <option value="as_type" <% if field_check = "as_type" then %>selected<% end if %>>처리유형</option>
                                    <option value="mg_ce" <% if field_check = "mg_ce" then %>selected<% end if %>>담당CE</option>
                                    <option value="acpt_man" <% if field_check = "acpt_man" then %>selected<% end if %>>접수자</option>
                                    <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>시도</option>
                                    <option value="gugun" <% if field_check = "gugun" then %>selected<% end if %>>구군</option>
                                    <option value="acpt_user" <% if field_check = "acpt_user" then %>selected<% end if %>>사용자</option>
                                    <option value="dept" <% if field_check = "dept" then %>selected<% end if %>>조직명</option>
                                    <option value="asets_no" <% if field_check = "asets_no" then %>selected<% end if %>>자산번호</option>
                                </select>
								</label>
                                <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:80px" id="field_view" >
								</label>
                                <label>
								<strong>회사 : </strong><%=company%>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="8%" >
							<col width="4%" >
							<col width="6%" >
							<col width="10%" >
							<col width="8%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="*" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">처리유형</th>
								<th scope="col">
							  <% 
								if tottal_record <> 0 then
							  		if (c_grade = "0") and (rs( "as_process")="접수" or rs("as_process")="입고" or rs("as_process")="연기") then 
							  %>
                                접수일자(변경)
                              <%   		else	%>
                                접수일자
                              <% 	end if	%>
							  <%  else	%>
                                접수일자
							  <% end if  %>
                                </th>
								<th scope="col">상태</th>
								<th scope="col">사용자</th>
								<th scope="col">
                                <% 
								if tottal_record <> 0 then
								  if c_grade < "3" and (rs("as_process")="접수" or rs("as_process")="입고" or rs("as_process")="연기") then 
								%>
								  조직명(결과등록)
									<% else %>
								  조직명
								  <% end if %>
								<% else %>
								  조직명
								<% end if %>
                                </th>
								<th scope="col">
								<% 
                                if tottal_record <> 0 then
                              	  if (c_grade < "3" or rs("acpt_man") = user_name ) and (rs("as_process")="접수" or rs("as_process")="입고" or rs("as_process")="연기") then 
                                %>
                              	   전화(수정)
                                    <% else %>
                                   전화번호
                                <% end if %>
                                <% else %>
                                   전화번호
                                <% end if %>
                                </th>
								<th scope="col">요청일자</th>
								<th scope="col">처리일자</th>
								<th scope="col">담당CE</th>
								<th scope="col">장애내용</th>
								<th scope="col">조회</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

							dim len_date, hangle, bit01, bit02, bit03
							acpt_date = rs("acpt_date")
							len_date = len(acpt_date)
							bit01 = left(acpt_date, 10)
						' 	bit01 = Replace(bit01,"-",".")
							bit03 = left(right(acpt_date, 5), 2)
							hangle = mid(acpt_date, 12, 2)
							if len_date = 22 then
								bit02 = mid(acpt_date, 15, 2)
							  else
								bit02 = "0"&mid(acpt_date, 15, 1)
							end If
						 
							if hangle = "오후" and bit02 <> 12 then 
								bit02 = bit02 + 12
							end if
							
							date_to_date = bit01 & " " &bit02 & ":" & bit03
							acpt_date = mid(date_to_date,3)
'							acpt_date = replace(acpt_date,"-","/")

							as_memo = replace(rs("as_memo"),chr(34),chr(39))
							view_memo = as_memo
							if len(as_memo) > 18 then
								view_memo = mid(as_memo,1,18) + ".."
							end if
							if rs("as_process") = "접수" or rs("as_process") = "연기" or rs("as_process") = "입고" then
								visit_date = "."
								if rs("as_process") = "입고" then
									sql_into = "select in_process, into_date from as_into where acpt_no = "&rs("acpt_no")&" and in_process = '수리완료'"
									Set rs_into=DbConn.Execute(sql_into)
									if rs_into.eof or rs_into.bof then
										visit_date = "."
									  else 
										visit_date = rs_into("into_date")
									end if
									rs_into.close()
								end if			
							  else
								visit_date = mid(rs("visit_date"),3)
							end if 
						%>
							<tr>
								<td class="first"><%=rs("as_type")%></td>
								<td>
						<% if (c_grade = "0") and (rs( "as_process")="접수" or rs("as_process")="입고" or rs("as_process")="연기") then %>
								<a href="#" onClick="pop_Window('acpt_date_mod.asp?acpt_no=<%=rs("acpt_no")%>','acpt_date_mod_pop','scrollbars=yes,width=600,height=250')"><%=acpt_date%></a>
                        <% else %>
                                    <%=acpt_date%>
                        <% end if %>		  
                                </td>
								<td>
						<% if (c_grade = "0") and (rs( "as_process")="완료") then %>
								<a href="#" onClick="pop_Window('as_process_cancel.asp?acpt_no=<%=rs("acpt_no")%>','as_process_cancel_pop','scrollbars=yes,width=600,height=250')"><%=rs("as_process")%></a>
                        <% else %>
								<%=rs("as_process")%>
                        <% end if 	%>
                                </td>
								<td><%=rs("acpt_user")%></td>
								<td>
								  <% if c_grade < "3" and (rs("as_process")="접수" or rs("as_process")="입고" or rs("as_process")="연기") then %>
									<% if rs("large_paper_no") = "" or isnull(rs("large_paper_no")) then  %>
                                    <a href="as_result_reg.asp?acpt_no=<%=rs("acpt_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&company=<%=company%>"><%=rs("dept")%></a>
									<%  else	%>
                            		<a href="#" onClick="pop_Window('large_result_reg.asp?acpt_no=<%=rs("acpt_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_c=<%=view_c%>&dong=<%=dong%>&view_sort=<%=view_sort%>','lage_result_reg_popup','scrollbars=yes,width=750,height=450')"><%=rs("dept")%></a>
                                  <% end if	%>  
                                  <% else %>
                                    <%=rs("dept")%>
                                  <% end if %>		  
                                </td>
								<td>
								  <% if (c_grade < "3" or rs("acpt_man") = user_name ) and (rs( "as_process")="접수" or rs("as_process")="입고" or rs("as_process")="연기") then %>
                                  <a href="#" onClick="pop_Window('as_mod_reg.asp?acpt_no=<%=rs("acpt_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&company=<%=company%>','as_mod_pop','scrollbars=yes,width=1000,height=450')"><%=rs("tel_ddd")%>)<%=rs("tel_no1")%>-<%=rs("tel_no2")%></a>
                                  <% else %>
                                    <%=rs("tel_ddd")%>)<%=rs("tel_no1")%>-<%=rs("tel_no2")%>
                                  <% end if %>		  
                                </td>
								<td><%=mid(rs("request_date"),3)%></td>
								<td><%=visit_date%></td>
								<td><%=rs("mg_ce")%></td>
							  	<td class="left">
							<% if rs("as_process") = "완료" or rs("as_process") = "취소" then	%>
                                <a href="#" onClick="pop_Window('as_memo_mod.asp?acpt_no=<%=rs("acpt_no")%>','as_memo_mod_pop','scrollbars=yes,width=600,height=300')"><%=as_memo%></a>
							<%   else	%>
								<%=as_memo%>
                            <% end if	%>
                               </td>
								<td><a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=rs("acpt_no")%>','asview_pop','scrollbars=yes,width=800,height=700')">조회</a></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
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
                    <a href="excel_down_condi.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&company=<%=company%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="as_list_reside.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="as_list_reside.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="as_list_reside.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="as_list_reside.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[다음]</a> <a href="as_list_reside.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
                    </td>
			      </tr>
				  </table>
				<input type="hidden" name="user_id">
				<input type="hidden" name="pass">
			</form>
		</div>				
	</div>        				
	</body>
</html>

