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
Dim field_check, field_view, cost_year
Dim start_page, pgsize, stpage
Dim total_page
Dim where_sql, order_sql, field_sql
Dim sum_price, sum_cost, sum_cost_vat
Dim rsCount, total_record
Dim rs_sum, rs, title_line
Dim sales_saupbu, from_date, to_date, work_month
Dim page, be_pg, pg_url

'Dim ck_sw, sales_month, sales_yymm

be_pg = "/finance/sales_bill_mg.asp"

'ck_sw = Request("ck_sw")

'If ck_sw = "y" Then
'	sales_month = Request("sales_month")
'	sales_saupbu = Request("sales_saupbu")
'	field_check = Request("field_check")
'	field_view = Request("field_view")
'Else
'	sales_month = Request.Form("sales_month")
'	sales_saupbu = Request.Form("sales_saupbu")
'	field_check = Request.Form("field_check")
'	field_view = Request.Form("field_view")
'End If

page = f_Request("page")
sales_saupbu = f_Request("sales_saupbu")
field_check = f_Request("field_check")
field_view = f_Request("field_view")
from_date = f_Request("from_date")
to_date = f_Request("to_date")

'If sales_month = "" Then
'	sales_month = Mid(Now(), 1, 4) & Mid(Now(), 6, 2)
'	sales_saupbu = "전체"
'	field_check = "total"
'	field_view = ""
'End If

'If field_check = "total" Then
'	field_view = ""
'End If

'sales_yymm = Mid(sales_month, 1, 4) & "-" & Mid(sales_month, 5, 2)
'cost_year = Mid(sales_month, 1, 4)

If sales_saupbu = "" Or IsNull(sales_saupbu) Then
	sales_saupbu = "전체"
End If

If field_check = "" Or IsNull(field_check) Then
	field_check = "total"
End If

'조회 날짜 설정
work_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

If from_date = "" Then
    from_date = Mid(work_month, 1, 4) & "-" & Mid(work_month, 5, 2) & "-01"
End If

If to_date = "" Then
    to_date = CStr(DateAdd("d", -1, DateAdd("m", 1, DateValue(from_date))))
End If

cost_year = Mid(from_date, 1, 4)

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&from_date="&from_date&"&to_date="&to_date&"&sales_saupbu="&sales_saupbu&"&field_check="&field_check&"&field_view="&field_view

If field_check = "total" Then
field_view = ""
	field_sql = " "
Else
	field_sql = " AND ("&field_check&" LIKE '%"&field_view&"%') "
End If

'Select Case sales_saupbu
'	Case "전체"
'		where_sql = " "
'	Case "회사간거래", "기타사업부"
'		where_sql = "AND sst.saupbu = '"&sales_saupbu&"' "
'	Case Else
'		where_sql = "AND eomt.org_bonbu = '"&sales_saupbu&"' "
'End Select

If sales_saupbu = "전체" Then
	where_sql = ""
Else
	where_sql =" AND sst.saupbu = '"&sales_saupbu&"' "
End If

order_sql = " ORDER BY sales_date ASC "

'매출 건수 조회
objBuilder.Append "SELECT COUNT(*) FROM saupbu_sales AS sst "
'objBuilder.Append "INNER JOIN emp_master AS emtt ON sst.emp_no = emtt.emp_no "
'objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

'objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&sales_yymm&"' "
objBuilder.Append "WHERE sales_date BETWEEN '"&from_date&"' AND '"&to_date&"' "

objBuilder.Append field_sql & where_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(RsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

'매출 총계 조회
objBuilder.Append "SELECT SUM(sales_amt) AS price, SUM(cost_amt) AS cost, "
objBuilder.Append "SUM(vat_amt) AS cost_vat "
objBuilder.Append "FROM saupbu_sales AS sst "
'objBuilder.Append "INNER JOIN emp_master AS emtt ON sst.emp_no = emtt.emp_no "
'objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

'objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&sales_yymm&"' "
objBuilder.Append "WHERE sales_date BETWEEN '"&from_date&"' AND '"&to_date&"' "

objBuilder.Append field_sql & where_sql

Set rs_sum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs_sum("price")) Then
	sum_price = 0
	sum_cost = 0
	sum_cost_vat = 0
Else
	sum_price = CDbl(rs_sum("price"))
	sum_cost = CDbl(rs_sum("cost"))
	sum_cost_vat = CDbl(rs_sum("cost_vat"))
End If

rs_sum.Close() : Set rs_sum = Nothing

'매출 조회
objBuilder.Append "SELECT sst.sales_date, sst.sales_company, sst.saupbu, sst.company, sst.trade_no, "
objBuilder.Append "	sst.group_name, sst.sales_amt, sst.cost_amt, sst.vat_amt, sst.emp_name, sst.sales_memo, "
objBuilder.Append "	sst.approve_no,	sst.emp_no	"
'objBuilder.Append "	eomt.org_company, eomt.org_bonbu "
objBuilder.Append "FROM saupbu_sales AS sst "
'objBuilder.Append "INNER JOIN emp_master AS emtt ON sst.emp_no = emtt.emp_no "
'objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

'objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&sales_yymm&"' "
objBuilder.Append "WHERE sales_date BETWEEN '"&from_date&"' AND '"&to_date&"' "

objBuilder.Append field_sql & where_sql & order_sql & " "
objBuilder.Append "LIMIT "&stpage&","&pgsize

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "매출 업로드 내역 관리"
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
		<!--<script type="text/javascript" src="/java/js_window.js"></script>-->

		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}

			/*function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.sales_month.value == ""){
					alert ("매출년월을 선택하세요");
					return false;
				}
				return true;
			}*/

			$(function() {
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=from_date%>" );

				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck(){
				//var st_date = $("#datepicker1").datepicker({dateFormat: 'dd-mm-yy'});
				//var end_date = $("#datepicker2").datepicker({dateFormat : 'dd-mm-yy'});

				var fDate = $("#datepicker1").datepicker('getDate');
				var lDate = $("#datepicker2").datepicker('getDate');

				//console.log(fDate);
				//console.log(lDate);

				var diff = new Date(lDate - fDate);
				var days = diff/1000/60/60/24;

				if(fDate = ""){
					alert("검색 시작년월일이 없습니다.");
					return false;
				}

				if(lDate = ""){
					alert("검색 종료년월일이 없습니다.");
					return false;
				}

				//console.log(days);
				//return false;

				if(days < 0){
					alert("검색 시작년월일이 종료 년월일 보다 작을 수 없습니다.");
					return false;
				}

				document.frm.submit();
				return;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/account_cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/finance/sales_bill_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조회조건</dt>
                        <dd>
                            <p>
								<label>
									<!--<strong>매출년월 : </strong>
                                	<input name="sales_month" type="text" value="<%'=sales_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">-->

									&nbsp;&nbsp;<strong>시작일자&nbsp;</strong> :
									<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker1">
									&nbsp;~&nbsp;
									&nbsp;&nbsp;<strong>종료일자&nbsp;</strong> :
									<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker2">
								</label>
                                <label>
								<strong>영업본부</strong>
								<%
								Dim rs_org

								'objBuilder.Append "SELECT org_bonbu FROM emp_org_mst "
								'objBuilder.Append "WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
								'objBuilder.Append "	AND org_bonbu NOT IN (' ', '경영본부', '빅데이타연구소', '전략부문', '기술연구소', 'OA수행본부') "
								'objBuilder.Append "GROUP BY org_bonbu "
								'objBuilder.Append "ORDER BY FIELD(org_company, '케이원', '케이네트웍스', '케이시스템') ASC, org_bonbu ASC "

								objBuilder.Append "SELECT saupbu FROM sales_org "
								objBuilder.Append "WHERE sales_year >= '"&cost_year&"' "
								objBuilder.Append "	AND saupbu IN (SELECT saupbu FROM sales_org WHERE sales_year >= '"&cost_year&"' GROUP BY saupbu) "
								objBuilder.Append "GROUP BY saupbu "
								objBuilder.Append "ORDER BY sort_seq "

								Set rs_org = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
								%>
									<select name="sales_saupbu" id="sales_saupbu" style="width:150px">
										<option value="전체" <%If sales_saupbu = "전체" Then%>selected<%End If %>>전체</option>
							    <%
							    Do Until rs_org.EOF
							    %>
										<option value='<%=rs_org("saupbu")%>' <%If sales_saupbu = rs_org("saupbu") Then%>selected<%End If %>><%=rs_org("saupbu")%></option>
                                <%
                                        rs_org.MoveNext()
                                Loop
                                rs_org.Close() : Set rs_org = Nothing
                                %>

									</select>
                                </label>
                                <label>
								<strong>세부조건</strong>
									<select name="field_check" id="field_check" style="width:100px">
										<option value="total" <%If field_check = "total" Then %>selected<%End If %>>전체</option>
										<option value="sst.sales_company" <%If field_check ="sales_company" Then %>selected<%End If%>>매출회사</option>
										<option value="sst.company" <%If field_check = "company" Then %>selected<%End If%>>고객사</option>
										<option value="sst.trade_no" <%If field_check = "trade_no" Then %>selected<%End If %>>사업자번호</option>
										<option value="sst.emp_name" <%If field_check = "emp_name" Then %>selected<%End If %>>담당자</option>
										<option value="sst.sales_memo" <%If field_check = "sales_memo" Then %>selected<%End If %>>품목명</option>
									</select>
								</label>
                                <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:150px" id="field_view" >
								</label>
            					<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="9%" >
							<col width="8%" >
							<col width="12%" >
							<col width="8%" >
							<col width="9%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="5%" >
							<col width="*" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">매출일자</th>
								<th scope="col">매출회사</th>
								<th scope="col">영업본부</th>
								<th scope="col">고객사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">그룹</th>
								<th scope="col">합계금액</th>
								<th scope="col">공급가액</th>
								<th scope="col">세액</th>
								<th scope="col">담당자</th>
								<th scope="col">품목명</th>
								<th scope="col">변경</th>
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
								<td class="right"><%=FormatNumber(sum_price, 0)%></td>
								<td class="right"><%=FormatNumber(sum_cost, 0)%></td>
								<td class="right"><%=FormatNumber(sum_cost_vat, 0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<%
							Do Until rs.EOF
							%>
							<tr>
								<td class="first"><%=rs("sales_date")%></td>
								<td><%=rs("sales_company")%></td>
								<td><%=rs("saupbu")%></td>
								<td><%=rs("company")%></td>
								<td><%=Mid(rs("trade_no"), 1, 3)%>-<%=Mid(rs("trade_no"), 4, 2)%>-<%=Right(rs("trade_no"), 5)%></td>
								<td><%=rs("group_name")%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rs("sales_amt"),0)%></td>
								<td class="right"><%=FormatNumber(rs("cost_amt"),0)%></td>
								<td class="right"><%=FormatNumber(rs("vat_amt"),0)%></td>
								<td><%=rs("emp_name")%>&nbsp;</td>
								<td class="left"><%=rs("sales_memo")%></td>
								<td>
								<%
								If rs("saupbu") = "기타사업부" Then
									'기타사업부일 경우 재무이사, 시스템관리자만 노출 처리[허정호_20220127]
									If user_id = "100359" Or user_id = "102592"  Then
								%>
										<a href="#" onClick="pop_Window('/sales/sales_saupbu_mod.asp?approve_no=<%=rs("approve_no")%>','sales_saupbu_mod_pop','scrollbars=yes,width=800,height=250')">수정</a>
								<%
									End If
								Else
									If sales_grade <= "1" Then
								%>
										<a href="#" onClick="pop_Window('/sales/sales_saupbu_mod.asp?approve_no=<%=rs("approve_no")%>','sales_saupbu_mod_pop','scrollbars=yes,width=800,height=250')">수정</a>
								<%
									End If
								End If
								%>
                                </td>
							</tr>
						<%
							rs.MoveNext()
						Loop
						rs.Close() : Set rs = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="24%">
						<div class="btnCenter">
							<a href="/sales/sales_report_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sales_saupbu=<%=sales_saupbu%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">엑셀다운로드</a>
						</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)

					DBConn.Close() : Set DBConn = Nothing
					%>
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

