<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

be_pg = "insa_promotion_list.asp"

Page=Request("page")
to_date = request("to_date")
in_grade = request("in_grade")
in_company = request("in_company")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	to_date=Request.form("to_date")
    in_grade=Request.form("in_grade")
	in_company=Request.form("in_company")
  else
	to_date = request("to_date")
    in_grade = request("in_grade")
	in_company = request("in_company")
end if

if in_company = "" then
	in_company = "케이원정보통신"
	to_date = curr_date
	in_grade = "대리2급"
end if

if in_grade = "대리2급" then
	condi_sql = "emp_grade like '%사원%' and "
end if
if in_grade = "대리1급" then
	condi_sql = "emp_grade like '%대리2급%' and "
end if
if in_grade = "과장" then
	condi_sql = "(emp_grade like '%대리2급%') or (emp_grade like '%대리1급%') and "
end if
if in_grade = "차장" then
	'condi_sql = "emp_grade and '과장' and "
	condi_sql = "emp_grade like '%과장%' and "
end if
if in_grade = "부장" then
	condi_sql = "emp_grade like '%차장%' and "
end if

pgsize = 10 ' 화면 한 페이지
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

target_date = to_date

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


tottal_record = 0

Sql = "SELECT * FROM emp_master where "+condi_sql+"isNull(emp_end_date) or emp_end_date = '1900-01-01'"
Set RsCount = Dbconn.Execute (sql)

do until RsCount.eof
   if RsCount("emp_grade_date") = "1900-01-01" then
      emp_grade_date = ""
      else
      emp_grade_date = RsCount("emp_grade_date")
   end if

   if emp_grade_date <> "" then
      year_cnt = datediff("yyyy", RsCount("emp_grade_date"), target_date)
      mon_cnt = datediff("m", RsCount("emp_grade_date"), target_date)
      day_cnt = datediff("d", RsCount("emp_grade_date"), target_date)
      else
      year_cnt = datediff("yyyy", RsCount("emp_first_date"), target_date)
      mon_cnt = datediff("m", RsCount("emp_first_date"), target_date)
      day_cnt = datediff("d", RsCount("emp_first_date"), target_date)
   end if

   target_cnt = cint(mon_cnt)

'   tottal_record = tottal_record + 1

   if (in_grade = "대리2급" or in_grade = "대리1급") and target_cnt > 24 then
      tottal_record = tottal_record + 1
      else if in_grade = "과장" and RsCount("emp_grade") = "대리1급" and target_cnt > 36 then
              tottal_record = tottal_record + 1
			  else if in_grade = "과장" and RsCount("emp_grade") = "대리2급" and target_cnt > 48 then
              tottal_record = tottal_record + 1
		           end if
		   end if
   end if
   RsCount.movenext()
loop
RsCount.close()

'tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM emp_master where "+condi_sql+"isNull(emp_end_date) or emp_end_date = '1900-01-01' ORDER BY emp_first_date,emp_no DESC limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1

title_line = " 승진대상자 현황 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "6 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=to_date%>" );
			});
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}

			function form_chk(){
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_asses_promo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_promotion_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>대상자 검색</dt>
                        <dd>
                            <p>
								<strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
                                rs_org.Open Sql, Dbconn, 1
							  %>
								<select name="in_company" id="in_company" style="width:120px">
                                <option value="" <% if in_company = "" then %>selected<% end if %>>선택</option>
                			  <%
								do until rs_org.eof
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If in_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()
								loop
								rs_org.Close()
							  %>
            					</select>
                                <strong>승진기준일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker">
                                <strong>승진직급 : </strong>
                              <%
								Sql="select * from emp_etc_code where emp_etc_type = '02' order by emp_etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
							  %>
								<select name="in_grade" id="in_grade" style="width:70px">
                                <option value="" <% if in_grade = "" then %>selected<% end if %>>선택</option>
                			  <%
								do until rs_etc.eof
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If in_grade = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()
								loop
								rs_etc.Close()
							  %>
            					</select>
                                <span>&nbsp;※ 승진기준은 매년 1월 1일 기준입니다.</span>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
							<col width="6%" >
							<col width="12%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
                                <th scope="col">생년월일</th>
								<th scope="col">현직급</th>
								<th scope="col">직책</th>
								<th scope="col">소속</th>
								<th scope="col">최초<br>입사일</th>
                                <th scope="col">입사일</th>
                                <th scope="col">최종<br>승진일</th>
                                <th scope="col">대상년한</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

						if rs("emp_grade_date") = "1900-01-01" then
						   emp_grade_date = ""
						   else
						   emp_grade_date = rs("emp_grade_date")
						end if

						if emp_grade_date <> "" then
						   year_cnt = datediff("yyyy", rs("emp_grade_date"), target_date)
                           mon_cnt = datediff("m", rs("emp_grade_date"), target_date)
                           day_cnt = datediff("d", rs("emp_grade_date"), target_date)
						   else
						       year_cnt = datediff("yyyy", rs("emp_first_date"), target_date)
                               mon_cnt = datediff("m", rs("emp_first_date"), target_date)
                               day_cnt = datediff("d", rs("emp_first_date"), target_date)
						end if

						target_cnt = cint(mon_cnt)

					if (in_grade = "대리2급" or in_grade = "대리1급") and target_cnt > 24 then

	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=emp_grade_date%>&nbsp;</td>
                                <td><%=mon_cnt%>&nbsp;개월</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
							</tr>
						<%
						      else if in_grade = "과장" and Rs("emp_grade") = "대리1급" and target_cnt > 36 then
	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;<td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=emp_grade_date%>&nbsp;</td>
                                <td><%=mon_cnt%>&nbsp;개월</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
							</tr>
						<%
						      else if in_grade = "과장" and Rs("emp_grade") = "대리2급" and target_cnt > 48 then
	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;<td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=emp_grade_date%>&nbsp;</td>
                                <td><%=mon_cnt%>&nbsp;개월</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
							</tr>
						<%
							     end if
							end if
						end if
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
                    <a href="insa_excel_promotlist.asp?in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_promotion_list.asp?page=<%=first_page%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_promotion_list.asp?page=<%=intstart -1%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_promotion_list.asp?page=<%=i%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_promotion_list.asp?page=<%=intend+1%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_promotion_list.asp?page=<%=total_page%>&in_company=<%=in_company%>&in_grade=<%=in_grade%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>

