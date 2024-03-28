<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim win_sw

ck_sw=Request("ck_sw")
Page=Request("page")

be_pg = "/card_owner_mg.asp"

If ck_sw = "y" Then
	use_yn=Request("use_yn")
	owner_company=Request("owner_company")
	field_check=Request("field_check")
	field_view=Request("field_view")
  else
	use_yn=Request.form("use_yn")
	owner_company=Request.form("owner_company")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
end if

If use_yn = "" Then
	use_yn = "Y"
	owner_company = "전체"
	field_check = "total"
	field_view = ""
End If

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

' 조건별 조회.........
close_sql = " where use_yn = '" + use_yn + "' "

if owner_company = "전체" then
	owner_company_sql = " "
  else
	owner_company_sql = " and ( owner_company = '" + owner_company + "' ) "
end if
if field_check = "total" then
	condi_sql = ""
  else
	condi_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

base_sql = "select * from card_owner"
order_sql = " ORDER BY card_no ASC"

sql = "select count(*) from card_owner " + close_sql + owner_company_sql + condi_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + close_sql + owner_company_sql + condi_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1
'Response.write Sql

title_line = "카드 사용자 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
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
				if (document.frm.use_yn.value == "") {
					alert ("사용여부를 선택하세요");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/card_slip_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="<%=be_pg%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건검색</dt>
                        <dd>
                            <p>
                                <label>
                              	<input type="radio" name="use_yn" value="Y" <% if use_yn="Y" then %>checked<% end if %> style="width:30px">사용
                              	<input type="radio" name="use_yn" value="N" <% if use_yn ="N" then %>checked<% end if %> style="width:30px">미사용
								</label>
                                <label>
								<strong>소유회사</strong>
								<%
								Call SelectEmpOrgList("owner_company", "owner_company", "width:120px", owner_company)
								%>
								<strong>항목조건</strong>
                                <select name="field_check" id="field_check" style="width:150px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                    <option value="card_type" <% if field_check = "card_type" then %>selected<% end if %>>카드종류</option>
                                    <option value="card_no" <% if field_check = "card_no" then %>selected<% end if %>>카드번호</option>
                                    <option value="emp_name" <% if field_check = "emp_name" then %>selected<% end if %>>사용자</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:150px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
							<col width="*" >
							<col width="10%" >
							<col width="10%" >
							<col width="9%" >
							<col width="6%" >
							<col width="4%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="10%" >
							<col width="3%" >
							<col width="3%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">카드종류</th>
								<th rowspan="2" scope="col">카드번호</th>
								<th rowspan="2" scope="col">소유회사</th>
								<th rowspan="2" scope="col">사용부서</th>
								<th rowspan="2" scope="col">사용인</th>
								<th rowspan="2" scope="col">차량유지비</th>
								<th rowspan="2" scope="col">한도</th>
								<th rowspan="2" scope="col">유효기간</th>
								<th rowspan="2" scope="col">발급일</th>
								<th rowspan="2" scope="col">사용개시일</th>
								<th rowspan="2" scope="col">비고</th>
								<th rowspan="2" scope="col">손익</th>
								<th rowspan="2" scope="col">수정</th>
								<th rowspan="2" scope="col">사용자</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							if rs("car_vat_sw") = "Y" then
								car_vat_view = "과세"
							  elseif rs("car_vat_sw") = "N" then
								car_vat_view = "비과세"
							  else
							  	car_vat_view = "경우에 따라"
							end if

							sql = "select * from memb where user_id ='"&rs("emp_no")&"'"
							set rs_emp = dbconn.execute(sql)
							if (rs_emp.eof or rs_bof) or (rs("emp_no") < "" or isnull(rs("emp_no"))) then
								org_name = "미등록"
								emp_grade = ""
							  else
								org_name = rs_emp("org_name")
								emp_grade = rs_emp("user_grade")
							end if
							sql = "select count(*) as hist_cnt from card_owner_history where card_no = '" + rs("card_no") + "'"
							set rs_hist=dbconn.execute(sql)
							if cint(rs_hist("hist_cnt")) > 0 then
								hist_sw = "y"
							  else
								hist_sw = "n"
							end if
							rs_hist.close()
						    %>
							<tr>
								<td class="first"><%=rs("card_type")%></td>
								<td>
								<%=rs("card_no")%>&nbsp;
								<%  if hist_sw = "y" then	%>
                                	<a href="#" onClick="pop_Window('/card_owner_hist_view.asp?card_no=<%=rs("card_no")%>','card_owner_hist_view_popup','scrollbars=yes,width=750,height=500')"><img src="image/hist.gif" width="24" height="11" border="0"></a>
                                <%  end if %>
                                </td>
								<td><%=rs("owner_company")%></td>
								<td><%=org_name%></td>
								<td><%=rs("emp_name")%>&nbsp;<%=emp_grade%></td>
								<td><%=car_vat_view%></td>
								<td><%=rs("card_limit")%>&nbsp;</td>
								<td><%=rs("valid_thru")%>&nbsp;</td>
								<td><%=rs("create_date")%>&nbsp;</td>
								<td><%=rs("start_date")%>&nbsp;</td>
                                <td><%=rs("card_memo")%>&nbsp;</td>
                                <td><%=rs("pl_yn")%>&nbsp;</td>
								<td>
                                    <a href="#" onClick="pop_Window('/card_owner_add.asp?card_no=<%=rs("card_no")%>&u_type=<%="U"%>','card_owner_add_pop','scrollbars=yes,width=850,height=340')">수정</a>
                                </td>
								<td>
                                    <a href="#" onClick="pop_Window('/card_owner_change.asp?card_no=<%=rs("card_no")%>&u_type=<%="U"%>','card_owner_change_popup','scrollbars=yes,width=850,height=200')">변경</a>
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
				    <td width="25%">
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="<%=be_pg%>?page=<%=first_page%>&use_yn=<%=use_yn%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[처음]</a>
                        <% if intstart > 1 then %>
                            <a href="<%=be_pg%>?page=<%=intstart -1%>&use_yn=<%=use_yn%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[이전]</a>
                        <% end if %>
                        <% for i = intstart to intend %>
                            <% if i = int(page) then %>
                                <b>[<%=i%>]</b>
                            <% else %>
                                <a href="<%=be_pg%>?page=<%=i%>&use_yn=<%=use_yn%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <% end if %>
                        <% next %>
                        <% if 	intend < total_page then %>
                            <a href="<%=be_pg%>?page=<%=intend+1%>&use_yn=<%=use_yn%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[다음]</a> <a href="<%=be_pg%>?page=<%=total_page%>&use_yn=<%=use_yn%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[마지막]</a>
                            <%	else %>
                            [다음]&nbsp;[마지막]
                        <% end if %>
                    </div>
                    </td>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('/card_owner_add.asp','card_owner_add_popup','scrollbars=yes,width=850,height=310')" class="btnType04">신규카드등록</a>
					</div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>

