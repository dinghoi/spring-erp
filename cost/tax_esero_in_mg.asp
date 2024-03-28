<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim page, bill_month, owner_company, field_check, field_view
Dim from_date, end_date, to_date, pgsize, start_page, stpage
Dim field_sql, owner_sql, rsCount, total_record, total_page
Dim rs_sum, sum_price, sum_cost, sum_cost_vat, rsTax, title_line
Dim rs_org, pg_url, be_pg
Dim arrTax

page = Request.QueryString("page")

bill_month = f_Request("bill_month")
owner_company = f_Request("owner_company")
field_check = f_Request("field_check")
field_view = f_Request("field_view")

title_line = "이세로 매입 세금계산서 관리"
be_pg = "/cost/tax_esero_in_mg.asp"

If bill_month = "" Then
	bill_month = Mid(Now(), 1, 4) & Mid(Now(), 6, 2)
	owner_company = "전체"
	field_check = "total"
	field_view = ""
End If

If field_check = "total" then
	field_view = ""
Else
	field_view = Trim(field_view)
End If

from_date = Mid(bill_month, 1, 4) & "-" & Mid(bill_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&bill_month="&bill_month&"&owner_company="&owner_company&"&field_check="&field_check&"&field_view="&field_view

Dim rsMail, mail_id, from_sql
'Dim base_sql, join_sql, order_sql

If field_check <> "total" Then
	'담당자 검색일 경우
	If field_check = "emp_name" Then
		objBuilder.Append "SELECT emtt.emp_email FROM emp_master AS emtt "
		objBuilder.Append "INNER JOIN memb AS memt ON emtt.emp_no = memt.user_id "
		objBuilder.Append "WHERE emtt.emp_name = '"&field_view&"' "
		objBuilder.Append "LIMIT 1 "

		Set rsMail = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		mail_id = rsMail("emp_email")

		rsMail.Close() : Set rsMail = Nothing

		field_sql = "AND SUBSTRING(receive_email, 1, INSTR(receive_email, '@') - 1) = '"&mail_id&"' "
	Else
		field_sql = "AND "&field_check&" LIKE '%"&field_view&"%' "
	End If
End If

If owner_company <> "전체" Then
	owner_sql = "AND owner_company = '"&owner_company&"' "
End If

from_sql = from_sql & "FROM tax_bill AS tabt "
from_sql = from_sql & "WHERE (tabt.bill_date >='"&from_date&"' AND tabt.bill_date <='"&to_date&"') "
from_sql = from_sql & "	AND tabt.end_yn = 'Y' AND tabt.cost_reg_yn = 'N' AND tabt.bill_id ='1' "

'Record Count
objBuilder.Append "SELECT SUM(price) AS 'price', SUM(cost) AS 'cost', SUM(cost_vat) AS 'cost_vat', COUNT(*) AS 'cnt' "
objBuilder.Append from_sql & owner_sql & field_sql

Set rsCount = DBConn.Execute (objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount("cnt"))

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

If IsNull(rsCount("price")) Then
	sum_price = 0
	sum_cost = 0
	sum_cost_vat = 0
Else
	sum_price = CDbl(rsCount("price"))
	sum_cost = CDbl(rsCount("cost"))
	sum_cost_vat = CDbl(rsCount("cost_vat"))
End If

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT r1.receive_email, r1.trade_no, r1.bill_date, r1.owner_company, r1.trade_name, "
objBuilder.Append "	r1.price, r1.cost, r1.cost_vat, r1.bill_collect, r1.tax_bill_memo, r1.approve_no, "
objBuilder.Append "	r1.trade_sw, emtt.emp_name, emtt.emp_bonbu, emtt.emp_saupbu "
objBuilder.Append "FROM ("
objBuilder.Append "SELECT tabt.receive_email, tabt.trade_no, tabt.bill_date, "
objBuilder.Append "	tabt.owner_company, tabt.trade_name, tabt.price, tabt.cost, tabt.cost_vat, "
objBuilder.Append "	tabt.bill_collect, tabt.tax_bill_memo, tabt.approve_no, "
objBuilder.Append "	(SELECT IF(IFNULL(trade_no, '') = '', 'N', 'Y') FROM trade WHERE trade_no = tabt.trade_no) AS 'trade_sw', "
objBuilder.Append "	IF(receive_email = '' OR IFNULL(receive_email, '') = '', NULL, "
objBuilder.Append "		(SELECT emtt.emp_no FROM emp_master AS emtt "
objBuilder.Append "		WHERE emtt.emp_email = SUBSTRING(tabt.receive_email, 1, INSTR(tabt.receive_email, '@') - 1) "
objBuilder.Append "			AND emtt.emp_pay_id <> '2' "
objBuilder.Append "		LIMIT 1) "
objBuilder.Append "	) AS 'emp_no' "
objBuilder.Append from_sql & owner_sql & field_sql
objBuilder.Append "ORDER BY tabt.bill_date, tabt.approve_no ASC "
objBuilder.Append "LIMIT "& stpage & "," &pgsize&") r1 "
objBuilder.Append "LEFT OUTER JOIN memb AS memt ON r1.emp_no = memt.user_id AND memt.grade < '5' "
objBuilder.Append "LEFT OUTER JOIN emp_master AS emtt ON r1.emp_no = emtt.emp_no "

Set rsTax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsTax.EOF Then
	arrTax = rsTax.getRows()
End If

rsTax.Close() : Set rsTax = Nothing
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
		<!--<script type="text/javascript" src="/java/js_window.js"></script>-->

		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.bill_month.value == ""){
					alert("년월을 선택하세요");
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
				<form action="<%=be_pg%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조회조건</dt>
                        <dd>
                            <p>
								<label>
								<strong>계산서 발행년월 : </strong>
                                	<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);" onkeypress="if(event.keyCode == '13'){frmcheck();}">
								</label>
                                <label>
									<strong>회사</strong>
									<select name="owner_company" id="owner_company" style="width:150px">
									  <option value="전체" <%If owner_company = "전체" Then %>selected<%End If %>>전체</option>
									  <%
										' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
										'sql = "SELECT org_name from emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') AND org_level = '회사' ORDER BY org_company ASC"

										objBuilder.Append "SELECT org_name FROM emp_org_mst "
										objBuilder.Append "WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
										objBuilder.Append "	AND org_level = '회사' ORDER BY org_company ASC "

										Set rs_org = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										Do Until rs_org.EOF
									  %>
										<option value='<%=rs_org("org_name")%>' <%If owner_company = rs_org("org_name") Then %>selected<%End If %>><%=rs_org("org_name")%></option>
									  <%
											rs_org.MoveNext()
										Loop
										rs_org.Close() : Set rs_org = Nothing
										DBConn.Close() : Set DBConn = Nothing
									  %>
									</select>
                                </label>
                                <label>
									<strong>세부조건</strong>
									<select name="field_check" id="field_check" style="width:100px">
										<option value="total" <%If field_check = "total" Then %>selected<%End If%>>전체</option>
										<option value="trade_name" <%If field_check = "trade_name" Then %>selected<%End If %>>상호명</option>
										<option value="tax_bill_memo" <%If field_check = "tax_bill_memo" Then %>selected<%End If %>>거래내역</option>
										<option value="receive_email" <%If field_check = "receive_email" Then %>selected<%End If %>>이메일</option>
										<!--담당자 추가[허정호_20220207]-->
										<option value="emp_name" <%If field_check = "emp_name" Then %>selected<%End If %>>담당자</option>
									</select>
								</label>
                                <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:150px" id="field_view" onkeypress="if(event.keyCode == '13'){frmcheck();}">
								</label>
            					<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색" /></a>

                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="10%" >
							<col width="7%" >
							<col width="11%" >
							<col width="8%" >
							<col width="6%" >
							<col width="13%" >
							<col width="7%" >
							<col width="6%" >
							<col width="3%" >
							<col width="*" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">발행일</th>
								<th scope="col">계산서소유회사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">상호명</th>
								<th scope="col">사업부</th>
								<th scope="col">담당자</th>
								<th scope="col">공급받는자이메일</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">청구</th>
								<th scope="col">거래내역</th>
								<th scope="col">등록구분</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><%=FormatNumber(total_record, 0)%>&nbsp;건</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right">&nbsp;</td>
								<td class="right"><%=FormatNumber(sum_cost, 0)%></td>
								<td class="right"><%=FormatNumber(sum_cost_vat, 0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<%
							Dim i, t_receive_email, t_trade_no, t_bill_date, t_owner_company, t_trade_name
							Dim t_price, t_cost, t_cost_vat, t_bill_collect, t_tax_bill_memo, t_approve_no
							Dim t_trade_sw, t_emp_name, t_emp_bonbu, t_emp_saupbu

							If IsArray(arrTax) Then
								For i = LBound(arrTax) To UBound(arrTax, 2)
									t_receive_email = f_toString(arrTax(0, i), "-")
									t_trade_no = arrTax(1, i)
									t_bill_date = arrTax(2, i)
									t_owner_company = arrTax(3, i)
									t_trade_name = arrTax(4, i)
									t_price = arrTax(5, i)
									t_cost = arrTax(6, i)
									t_cost_vat = arrTax(7, i)
									t_bill_collect = arrTax(8, i)
									t_tax_bill_memo = arrTax(9, i)
									t_approve_no = arrTax(10, i)
									t_trade_sw = arrTax(11, i)
									t_emp_name = f_toString(arrTax(12, i), "-")
									t_emp_bonbu = f_toString(arrTax(13, i), "-")
							%>
								<tr>
									<td class="first"><%=t_bill_date%></td>
									<td><%=t_owner_company%></td>
									<td><%=Mid(t_trade_no, 1, 3)%>-<%=Mid(t_trade_no, 4, 2)%>-<%=Right(t_trade_no, 5)%></td>
									<td><%=t_trade_name%></td>
									<td><%=t_emp_bonbu%>&nbsp;</td>
									<td><%=t_emp_name%></td>
									<td><%=t_receive_email%>&nbsp;</td>
									<td class="right"><%=FormatNumber(t_cost, 0)%></td>
									<td class="right"><%=FormatNumber(t_cost_vat, 0)%></td>
									<td><%=t_bill_collect%>&nbsp;</td>
									<td class="left"><%=t_tax_bill_memo%></td>
									<td>
							<%If t_trade_sw = "Y" Then %>
								<a href="#" onClick="pop_Window('/cost/tax_esero_in_detail_add.asp?approve_no=<%=t_approve_no%>&bill_month=<%=bill_month%>','tax_esero_in_detail_add_pop','scrollbars=yes,width=1000,height=280')">비용등록</a>
							<%Else%>
								<a href="#" onClick="pop_Window('/cost/tax_trade_add.asp?approve_no=<%=t_approve_no%>','tax_trade_add_pop','scrollbars=yes,width=800,height=450')">거래처등록</a>
							<%End If%>
									</td>
								</tr>
							<%
								Next
							End If
							%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
						<a href="/cost/excel/tax_esero_in_excel.asp?bill_month=<%=bill_month%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
					%>
                    </td>
				    <td width="30%">
					<div class="btnCenter">
					<%
					If account_grade = "0" Then
					%>
						<a href="/cost/excel/tax_esero_in_excel_detail.asp?bill_month=<%=bill_month%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">일괄업로드 엑셀</a>
						<a href="/cost/excel/cost_org_list_excel.asp" class="btnType04">조직코드 엑셀</a>
						<a href="/cost/excel/cost_trade_list_excel.asp" class="btnType04">거래처 엑셀</a>
					<%
					End If
					%>
					</div>
                    </td>
			      </tr>
				  </table>
				</form>
		</div>
	</div>
	</body>
</html>