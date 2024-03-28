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
        view_c = request("view_c")
        view_d = request("view_d")
		emp_name = request("emp_name")
	else
		bill_month = request.form("bill_month")
		slip_gubun = request.form("slip_gubun")
        view_c = request.form("view_c")
        view_d = request.form("view_d")
		emp_name = request.form("emp_name")
    end if

    if view_d = "" then
        view_d = "slip"
	end if

	if bill_month = "" then
		bill_month = mid(now(),1,4) + mid(now(),6,2)
		slip_gubun = "전체"
        view_c = "total"
        view_d = "slip"
		emp_name = ""
	end if

	if view_c = "total" then
		emp_name = ""
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

	if view_c = "total" then
		emp_sql = ""
	elseif view_c = "emp_name" then
	  	emp_sql = " and emp_name like '%"&emp_name&"%'"
	else
	  	emp_sql = " and customer like '%"&emp_name&"%'"
	end if

    base_sql = "select * from general_cost where (tax_bill_yn = 'Y') "
    if view_d = "slip" then
        base_sql = base_sql & " and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
        order_sql = " ORDER BY org_name, emp_name, slip_date ASC"
    end if
    if view_d = "reg" then
        base_sql = base_sql & " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59')"
        order_sql = " ORDER BY org_name, emp_name, reg_date ASC"
    end if

    sql = "select count(*) from general_cost where (tax_bill_yn = 'Y') "
    if view_d = "slip" then
        sql = sql & " and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')  "
    end if
    if view_d = "reg" then
        sql = sql &  " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59')"
    end if
    sql = sql + posi_sql + gubun_sql + emp_sql
	Set RsCount = Dbconn.Execute (sql)

	tottal_record = cint(RsCount(0)) 'Result.RecordCount

	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF

    sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from general_cost where (tax_bill_yn = 'Y') "
    if view_d = "slip" then
        sql = sql & "and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
    end if
    if view_d = "reg" then
        sql = sql &  " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59')"
    end if

    sql = sql +  posi_sql + gubun_sql + emp_sql

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

	sql = base_sql + posi_sql + gubun_sql + emp_sql + order_sql + " limit "& stpage & "," &pgsize
	Rs.Open Sql, Dbconn, 1

	title_line = "매입 세금계산서 관리"
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
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('emp_name_view').style.display = 'none';
				}
				if (eval("document.frm.view_c[1].checked") || eval("document.frm.view_c[2].checked")) {
					document.getElementById('emp_name_view').style.display = '';
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
				<form action="tax_bill_in_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조회조건</dt>
                        <dd>
                            <p>
								<label>
                                    <input type="radio" name="view_d" value="slip" <% if view_d = "slip" then %>checked<% end if %> style="width:25px">
                                    <strong>발생년월&nbsp;</strong>
                                    <input type="radio" name="view_d" value="reg" <% if view_d = "reg" then %>checked<% end if %> style="width:25px">
                                    <strong>발급년월&nbsp;</strong>

                                    : <input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
                                    (예201401)
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
								<label>
								<strong>조회범위 : </strong>
                              	<input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">전체
                                <input type="radio" name="view_c" value="emp_name" <% if view_c = "emp_name" then %>checked<% end if %> style="width:25px" onClick="condi_view()">개인별
                                <input type="radio" name="view_c" value="customer" <% if view_c = "customer" then %>checked<% end if %> style="width:25px" onClick="condi_view()">외주업체
								</label>
								<label>
                                	<input name="emp_name" type="text" value="<%=emp_name%>" style="width:100px; display:none" id="emp_name_view">
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
							<col width="7%" >
							<col width="4%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="11%" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
							<col width="4%" >
							<col width="7%" >
							<col width="11%" >
							<col width="3%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사용조직</th>
								<th scope="col">담당영업<br>사업부</th>
								<th scope="col">담당자</th>
								<th scope="col">발행일자</th>
								<th scope="col">발급일자</th>
								<th scope="col">고객사</th>
								<th scope="col">외주업체</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">유형</th>
								<th scope="col">세부유형</th>
								<th scope="col">발행내역</th>
								<th scope="col">손익</th>
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
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(sum_price,0)%></td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
								<td class="right"><%=formatnumber(sum_cost_vat,0)%></td>
								<td>&nbsp;</td>
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
                                <%
                                ' 5일 이후 지연 입력건 검출...
                                chk_slip_month = mid(rs("slip_date"),1,7)
                                chk_reg_month = mid(rs("reg_date"),1,7)
                                chk_reg_day = mid(rs("reg_date"),9,2)

                                if ((chk_slip_month < chk_reg_month) and (chk_reg_day > "05")) then
                                    bgcolor = "burlywood"
                                else
                                    bgcolor = "#f8f8f8"
                                end if
                                %>
                                <tr style="background-color: <%=bgcolor%>;">
                                    <td class="first"><%=rs("org_name")%></td>
                                    <td><%=rs("mg_saupbu")%>&nbsp;</td>
                                    <td><%=rs("emp_name")%></td>
                                    <td><%=rs("slip_date")%></td>
                                    <td><%=mid(rs("reg_date"),1,10)%></td>
                                    <td><%=rs("company")%></td>
                                    <td><%=rs("customer")%></td>
                                    <td class="right"><%=formatnumber(rs("price"),0)%></td>
                                    <td class="right"><%=formatnumber(rs("cost"),0)%></td>
                                    <td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
                                    <td><%=rs("slip_gubun")%></td>
                                    <td><%=rs("account")%></td>
                                    <td><%=rs("slip_memo")%></td>
                                    <td><%=rs("pl_yn")%></td>
                                    <td>
                                    <% if rs("end_yn") = "C" or rs("end_yn") = "N" then %>
                                        <% if (rs("reg_id") = user_id) or (rs("emp_no") = user_id) or cost_grade = "0" or position ="사업부장" or position = "본부장"  then	%>
                                            <a href="#" onClick="pop_Window('tax_bill_in_mod.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','tax_bill_in_mod_pop','scrollbars=yes,width=1000,height=300')">수정</a>
                                        <% else	%>
                                            불가
                                        <% end if %>
                                    <% else	%>
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
                    <a href="tax_bill_in_excel.asp?bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="tax_bill_in_mg.asp?page=<%=first_page%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[처음]</a>
                        <% if intstart > 1 then %>
                            <a href="tax_bill_in_mg.asp?page=<%=intstart -1%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[이전]</a>
                        <% end if %>
                        <% for i = intstart to intend %>
                            <% if i = int(page) then %>
                                <b>[<%=i%>]</b>
                            <% else %>
                                <a href="tax_bill_in_mg.asp?page=<%=i%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <% end if %>
                        <% next %>
                        <% if intend < total_page then %>
                            <a href="tax_bill_in_mg.asp?page=<%=intend+1%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[다음]</a>
                            <a href="tax_bill_in_mg.asp?page=<%=total_page%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[마지막]</a>
                        <% else %>
                            [다음]&nbsp;[마지막]
                        <% end if %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnCenter">
					</div>
                    </td>
			      </tr>
				  </table>
				</form>
		</div>
	</div>
	</body>
</html>

