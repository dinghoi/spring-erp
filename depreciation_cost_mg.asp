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
		slip_month = request("slip_month")
		account = request("account")
	else
		slip_month = request.form("slip_month")
		account = request.form("account")
	end if

	if slip_month = "" then
		slip_month = mid(now(),1,4) + mid(now(),6,2)
		account = "전체"
	end if

	from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
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

	if cost_grade = "0" then
		posi_sql = ""
	end if

	if account = "전체" then
		gubun_sql = ""
	  else
	  	gubun_sql = " and account = '"&account&"' "
	end if

	base_sql = "select * from general_cost where (slip_gubun = '상각비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
	order_sql = " ORDER BY org_name, emp_name, slip_date ASC"

	sql = "select count(*) from general_cost where (slip_gubun = '상각비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + posi_sql + gubun_sql
	Set RsCount = Dbconn.Execute (sql)

	tottal_record = cint(RsCount(0)) 'Result.RecordCount

	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF

	sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from general_cost where (slip_gubun = '상각비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + posi_sql + gubun_sql
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

	title_line = "상각비 관리"
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
				if (document.frm.slip_month.value == "") {
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
				<form action="depreciation_cost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조회조건</dt>
                        <dd>
                            <p>
								<label>
								<strong>비용년월 : </strong>
                                	<input name="slip_month" type="text" value="<%=slip_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
                                <label>
                                <strong>상각비유형 : </strong>
                                <select name="account" id="account" style="width:120px">
                                  <option value='전체' <%If account = "전체" then %>selected<% end if %>>전체</option>
                                  <option value='대손상각비' <%If account = "대손상각비" then %>selected<% end if %>>대손상각비</option>
                                  <option value='고정자산' <%If account = "고정자산" then %>selected<% end if %>>고정자산</option>
                                  <option value='무형자산' <%If account = "무형자산" then %>selected<% end if %>>무형자산</option>
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
							<col width="15%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="*" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">비용회사</th>
								<th scope="col">비용일자</th>
								<th scope="col">담당자</th>
								<th scope="col">금액</th>
								<th scope="col">상각비유형</th>
								<th scope="col">상각비 세부내역</th>
								<th scope="col">수정</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><%=formatnumber(tottal_record,0)%>&nbsp;건</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
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
						%>
							<tr>
								<td class="first"><%=rs("emp_company")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("emp_name")%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
								<td>
							<% if rs("end_yn") = "C" or rs("end_yn") = "N" then %>
							<%   if (rs("reg_id") = user_id) or (rs("emp_no") = user_id) or cost_grade = "0" or position ="사업부장" or position = "본부장"  then	%>
                                <a href="#" onClick="pop_Window('depreciation_cost_add.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','depreciation_cost_add_pop','scrollbars=yes,width=800,height=200')">수정</a>
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
                    <a href="depreciation_cost_excel.asp?slip_month=<%=slip_month%>&account=<%=account%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="depreciation_cost_mg.asp?page=<%=first_page%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="depreciation_cost_mg.asp?page=<%=intstart -1%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="depreciation_cost_mg.asp?page=<%=i%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="depreciation_cost_mg.asp?page=<%=intend+1%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[다음]</a>
                        <a href="depreciation_cost_mg.asp?page=<%=total_page%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnRight">
					<a href="#" onClick="pop_Window('depreciation_cost_add.asp','depreciation_cost_add_pop','scrollbars=yes,width=800,height=200')" class="btnType04">상각비등록</a>
					</div>
                    </td>
			      </tr>
				  </table>
				</form>
		</div>
	</div>
	</body>
</html>

