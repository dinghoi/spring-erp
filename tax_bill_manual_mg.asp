<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	Dim Rs
	Dim Repeat_Rows
	Dim from_date
	Dim to_date
	Dim win_sw

	win_sw = "close"

	ck_sw=Request("ck_sw")
	Page=Request("page")

	if ck_sw = "y" Then
		bill_month = request("bill_month")
		slip_gubun = request("slip_gubun")
	else
		bill_month = request.form("bill_month")
		slip_gubun = request.form("slip_gubun")
	end if

	if bill_month = "" then
		bill_month = mid(now(),1,4) + mid(now(),6,2)
		slip_gubun = "전체"
	end if

	from_date = mid(bill_month,1,4) + "-" + mid(bill_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))

	pgsize = 10 ' 화면 한 페이지

	If Page = "" Then
		Page = 1
		start_page = 1
	End If
	stpage = int((page - 1) * pgsize)

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	Set RsCount = Server.CreateObject("ADODB.Recordset")
	Set Rscost = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect

' 포지션별
	posi_sql = " and (emp_no = '"&user_id&"' or reg_id = '"&user_id&"') "

	if position = "팀원" then
		view_condi = "본인"
	end if

	if position = "파트장" then
		if org_name = "한화생명호남" then
			posi_sql = " and (org_name = '한화생명호남' or org_name = '한화생명전북') "
		  else
			posi_sql = " and org_name = '"&org_name&"'"
		end if
	end if

	if position = "팀장" then
		posi_sql = " and team = '"&team&"'"
	end if

	if position = "사업부장" or cost_grade = "2" then
		posi_sql = " and saupbu = '"&saupbu&"'"
	end if

	if position = "본부장" or cost_grade = "1" then
		posi_sql = " and bonbu = '"&bonbu&"'"
	end if

	view_grade = position

	if cost_grade = "0" then
		posi_sql = ""
	end if

	if slip_gubun = "전체" then
		gubun_sql = ""
	  else
	  	gubun_sql = " and slip_gubun = '"&slip_gubun&"' "
	end if

	base_sql = "select * from general_cost where (tax_bill_yn = 'Y') and (manual_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
	order_sql = " ORDER BY org_name, emp_name, slip_date ASC"

	sql = "select count(*) from general_cost where (tax_bill_yn = 'Y') and (manual_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + posi_sql + gubun_sql
	Set RsCount = Dbconn.Execute (sql)

	tottal_record = cint(RsCount(0)) 'Result.RecordCount

	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF

	sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from general_cost where (tax_bill_yn = 'Y') and (manual_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + posi_sql + gubun_sql
	Set rs_sum = Dbconn.Execute (sql)
	if isnull(rs_sum("price")) then
		sum_price = 0
		sum_cost = 0
		sum_cost_vat = 0
	  else
		sum_price = cdbl(rs_sum("price"))
		sum_cost = cdbl(rs_sum("cost"))
		sum_cost_vat = cdbl(rs_sum("cost_vat"))
	end if

	sql = base_sql + posi_sql + gubun_sql + order_sql + " limit "& stpage & "," &pgsize
	Rs.Open Sql, Dbconn, 1

	title_line = "수작업 매입 세금계산서 관리"
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
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.bill_month.value == "") {
					alert ("년월을 선택하세요");
					return false;
				}
				return true;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="tax_bill_manual_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조회조건</dt>
                        <dd>
                            <p>
								<label>
								<strong>계산서 발행년월 : </strong>
                                	<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
                                <label>
                                <strong>비용유형 : </strong>
                                <select name="slip_gubun" id="slip_gubun" style="width:120px">
                                  <option value='전체' <%If slip_gubun = "전체" then %>selected<% end if %>>전체</option>
                                  <%
                                    Sql="select * from type_code where etc_seq = '4' and etc_id = 'T' order by type_name asc"
                                    rs_etc.Open Sql, Dbconn, 1
                                    do until rs_etc.eof
                                    %>
                                  <option value='<%=rs_etc("type_name")%>' <%If slip_gubun = rs_etc("type_name") then %>selected<% end if %>><%=rs_etc("type_name")%></option>
                                  <%
                                        rs_etc.movenext()
                                    loop
                                    rs_etc.close()
                                    %>
                                  <option value='비용' <%If slip_gubun = "비용" then %>selected<% end if %>>비용</option>
                                </select>
                                </label>
            					<a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="5%" >
							<col width="7%" >
							<col width="11%" >
							<col width="12%" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
							<col width="4%" >
							<col width="7%" >
							<col width="12%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사용조직</th>
								<th scope="col">담당영업사업부</th>
								<th scope="col">담당자</th>
								<th scope="col">발행일자</th>
								<th scope="col">고객사</th>
								<th scope="col">외주업체</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">유형</th>
								<th scope="col">세부유형</th>
								<th scope="col">발행내역</th>
								<th scope="col">수정</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><%=formatnumber(tottal_record,0)%>&nbsp;건</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(sum_price,0)%></td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
								<td class="right"><%=formatnumber(sum_cost_vat,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
						do until rs.eof
							if rs("end_yn") = "Y" then
								end_yn = "마감"
								end_view = "N"
							  elseif rs("end_yn") = "I" then
								end_yn = "결재중"
								end_view = "N"
							  else
							  	end_yn = "진행"
							end if
							org_name = rs("emp_company") + "/" + rs("org_name")
							customer_no = mid(rs("customer_no"),1,3) + "-" + mid(rs("customer_no"),4,2) + "-" + mid(rs("customer_no"),6)
						%>
							<tr>
								<td class="first"><%=rs("org_name")%></td>
								<td><%=rs("mg_saupbu")%>&nbsp;</td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("customer")%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("slip_gubun")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
								<td>
							<% if rs("end_yn") = "C" or rs("end_yn") = "N" then %>
							<%   if (rs("reg_id") = user_id) or (rs("emp_no") = user_id) or cost_grade = "0" or position ="사업부장" or position = "본부장"  then	%>
                                <a href="#" onClick="pop_Window('tax_bill_manual_add.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','tax_bill_manual_add_pop','scrollbars=yes,width=1000,height=310')">수정</a>
							<%     else	%>
								불가
                            <%	 end if	%>
							<%  else	%>
								마감
                        	<% end if %>
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
				    <td width="24%">
					<div class="btnCenter">
                    <a href="tax_bill_manual_excel.asp?bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="tax_bill_manual_mg.asp?page=<%=first_page%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="tax_bill_manual_mg.asp?page=<%=intstart -1%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="tax_bill_manual_mg.asp?page=<%=i%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="tax_bill_manual_mg.asp?page=<%=intend+1%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[다음]</a>
                        <a href="tax_bill_manual_mg.asp?page=<%=total_page%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnRight">
					<a href="#" onClick="pop_Window('tax_bill_manual_add.asp','tax_bill_manual_add_pop','scrollbars=yes,width=1000,height=310')" class="btnType04">종이 세금계산서 등록</a>
					</div>
                    </td>
			      </tr>
				  </table>
				</form>
		</div>
	</div>
	</body>
</html>

