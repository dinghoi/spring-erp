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
Dim ck_sw, page, bill_id, bill_month, cost_reg_yn, end_yn
Dim from_date, end_date, to_date
Dim pgsize, start_page, stpage
Dim cost_reg_sql, end_sql
Dim rsCost, cost_record, rs_mi_cost, mi_record, rsCount
Dim total_record, total_page, rs_sum, sum_price, sum_cost, sum_cost_vat
Dim rs, title_line, be_pg, pg_url

'ck_sw = Request("ck_sw")
page = Request.QueryString("page")

'If ck_sw = "y" Then
'	bill_id = Request("bill_id")
'	bill_month = Request("bill_month")
'	cost_reg_yn = Request("cost_reg_yn")
'	end_yn = Request("end_yn")
'Else
'	bill_id = Request.Form("bill_id")
'	bill_month = Request.Form("bill_month")
'	cost_reg_yn = Request.Form("cost_reg_yn")
'	end_yn = Request.Form("end_yn")
'End If

bill_id = f_Request("bill_id")
bill_month = f_Request("bill_month")
cost_reg_yn = f_Request("cost_reg_yn")
end_yn = f_Request("end_yn")

be_pg = "/finance/tax_esero_mg.asp"

If bill_month = "" Then
	bill_month = Mid(Now(), 1, 4) & Mid(Now(), 6, 2)
	bill_id = "1"
	cost_reg_yn = "T"
	end_yn = "T"
End If

from_date = Mid(bill_month, 1, 4) & "-" & Mid(bill_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&bill_id="&bill_id&"&bill_month="&bill_month&"&cost_reg_yn="&cost_reg_yn&"&end_yn="&end_yn

'If cost_reg_yn = "T" Then
'	cost_reg_sql = " "
'Else
If cost_reg_yn <> "T" Then
	cost_reg_sql = "AND cost_reg_yn = '"&cost_reg_yn&"' "
End If

'If end_yn = "T" Then
'	end_sql = ""
'Else
If end_yn <> "T" Then
	end_sql = "AND end_yn = '"&end_yn&"' "
End If

' 비용 등록 여부 확인
objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM tax_bill "
objBuilder.Append "WHERE (bill_date >='"&from_date&"' AND bill_date <='"&to_date&"') "
objBuilder.Append "	AND cost_reg_yn = 'Y' "
objBuilder.Append "	AND bill_id = '"&bill_id&"' "

Set rsCost = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

cost_record = CInt(rsCost(0)) 'Result.RecordCount

rsCost.Close() : Set rsCost = Nothing
' 비용등록 확인 끝

' 비용 미등록 여부 확인
'sql = "select count(*) from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (cost_reg_yn = 'N') and (bill_id = '"&bill_id&"') "
objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM tax_bill "
objBuilder.Append "WHERE (bill_date >='"&from_date&"' AND bill_date <='"&to_date&"') "
objBuilder.Append "	AND cost_reg_yn = 'N' "
objBuilder.Append "	AND bill_id = '"&bill_id&"' "

Set rs_mi_cost = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

mi_record = CInt(rs_mi_cost(0)) 'Result.RecordCount

rs_mi_cost.Close() : Set rs_mi_cost = Nothing
' 비용 미등록 확인 끝

'sql = "select count(*) from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (bill_id = '"&bill_id&"') " + cost_reg_sql + end_sql
objBuilder.Append "SELECT SUM(price) AS price, SUM(cost) AS cost, SUM(cost_vat) AS cost_vat, COUNT(*) AS 'cnt' "
objBuilder.Append "FROM tax_bill "
objBuilder.Append "WHERE (bill_date >='"&from_date&"' AND bill_date <='"&to_date&"') "
objBuilder.Append "	AND bill_id = '"&bill_id&"' "
objBuilder.Append cost_reg_sql & end_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'total_record = CInt(RsCount(0)) 'Result.RecordCount
total_record = CInt(rsCount("cnt"))

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

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

objBuilder.Append "SELECT cost_reg_yn, send_email, bill_date, trade_no, "
objBuilder.Append "	CASE "
objBuilder.Append "		WHEN owner_company = '케이원정보통신' THEN '케이원' "
objBuilder.Append "		WHEN owner_company = '코리아디엔씨' THEN '케이시스템' "
objBuilder.Append "		ELSE owner_company "
objBuilder.Append "	END AS owner_company, "
objBuilder.Append "	trade_name, trade_owner, price, cost, cost_vat, bill_collect, "
objBuilder.Append "	tax_bill_memo, end_yn, approve_no "
objBuilder.Append "FROM tax_bill "
objBuilder.Append "WHERE (bill_date >='"&from_date&"' AND bill_date <='"&to_date&"') "
objBuilder.Append "	AND bill_id = '"&bill_id&"' "
objBuilder.Append cost_reg_sql & end_sql
objBuilder.Append "ORDER BY bill_date ASC "
objBuilder.Append "LIMIT "& stpage & "," &pgsize

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "이세로 세금계산서 관리"
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
				return "1 1";
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if(document.frm.bill_id.value == ""){
					alert ("계산서 유형을 선택하세요");
					return false;
				}

				if(document.frm.bill_month.value == ""){
					alert ("년월을 선택하세요");
					return false;
				}

				if(document.frm.cost_reg_yn.value == ""){
					alert ("비용등록 여부를 선택하세요");
					return false;
				}

				return true;
			}

			function upload_cancel(){
				a = confirm('업로드를 취소하겠습니까?');

				if(a==true){
					document.frm.action = "/finance/tax_bill_upload_cancel.asp";
               		document.frm.submit();
					return true;
				}

				return false;
			}

			function end_process(){
				a = confirm('마감 하시겠습니까?');

				if(a==true){
					document.frm.action = "/finance/tax_esero_end.asp";
               		document.frm.submit();
					return true;
				}

				return false;
			}

			function cancel_process(){
				a = confirm('취소 하시겠습니까?');

				if(a==true){
					document.frm.action = "/finance/tax_esero_end_cancel.asp";
               		document.frm.submit();
					return true;
				}

				return false;
			}

			function tax_esero_del(page, t_id, b_id, b_month, c_yn, e_yn){
				var param = '?page='+page+'&t_id='+t_id+'&b_id='+b_id+'&b_month='+b_month+'&c_yn='+c_yn+'&e_yn='+e_yn;

				if(confirm('삭제 처리하시겠습니까?')){

					location.href = 'tax_esero_del.asp'+param;
					return;
				}
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/tax_bill_menu.asp" -->
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
								<strong>계산서 유형 : </strong>
                              	<input type="radio" name="bill_id" value="1" <%If bill_id = "1" Then %>checked<%End If %> style="width:25px">매입
                                <input type="radio" name="bill_id" value="2" <%If bill_id = "2" Then %>checked<%End If %> style="width:25px">매출
								</label>
								<label>
								<strong>발행년월 : </strong>
                                	<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
								<label>
								<strong>비용등록여부 : </strong>
                              	<input type="radio" name="cost_reg_yn" value="T" <%If cost_reg_yn = "T" Then %>checked<%End If %> style="width:25px">전체
                                <input type="radio" name="cost_reg_yn" value="Y" <%If cost_reg_yn = "Y" Then %>checked<%End If %> style="width:25px">등록
                                <input type="radio" name="cost_reg_yn" value="N" <%If cost_reg_yn = "N" Then %>checked<%End If %> style="width:25px">미등록
								</label>
								<label>
								<strong>마감여부 : </strong>
                              	<input type="radio" name="end_yn" value="T" <%If end_yn = "T" Then %>checked<%End If %> style="width:25px">전체
                                <input type="radio" name="end_yn" value="Y" <%If end_yn = "Y" Then %>checked<%End If %> style="width:25px">Yes
                                <input type="radio" name="end_yn" value="N" <%If end_yn = "N" Then %>checked<%End If %> style="width:25px">No
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
							<col width="10%" >
							<col width="7%" >
							<col width="11%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="3%" >
							<col width="12%" >
							<col width="*" >
							<col width="3%" >
							<col width="4%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">발행일</th>
								<th scope="col">계산서소유회사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">상호</th>
								<th scope="col">대표자명</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">청구</th>
								<th scope="col">계산서이메일</th>
								<th scope="col">품목명</th>
								<th scope="col">마감</th>
								<th scope="col">비용</th>
								<th scope="col">&nbsp;</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>건수</strong></td>
								<td><%=FormatNumber(total_record, 0)%>&nbsp;건</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=FormatNumber(sum_price, 0)%></td>
								<td class="right"><%=FormatNumber(sum_cost, 0)%></td>
								<td class="right"><%=FormatNumber(sum_cost_vat, 0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
						Dim cost_reg_view, email_view

						Do Until rs.EOF
							If rs("cost_reg_yn") = "Y" Then
								cost_reg_view = "등록"
							Else
							  	cost_reg_view = "미등록"
							End If

							If bill_id = "1" Then
								email_view = rs("send_email")
							Else
							  	email_view = rs("receive_email")
							End If
						%>
							<tr>
								<td class="first"><%=rs("bill_date")%></td>
								<td><%=rs("owner_company")%></td>
								<td><%=Mid(rs("trade_no"), 1, 3)%>-<%=Mid(rs("trade_no"), 4, 2)%>-<%=Right(rs("trade_no"), 5)%></td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("trade_owner")%></td>
								<td class="right"><%=FormatNumber(rs("price"), 0)%></td>
								<td class="right"><%=FormatNumber(rs("cost"), 0)%></td>
								<td class="right"><%=FormatNumber(rs("cost_vat"), 0)%></td>
								<td><%=rs("bill_collect")%></td>
								<td><%=email_view%>&nbsp;</td>
								<td class="left"><%=rs("tax_bill_memo")%></td>
								<td><%=rs("end_yn")%></td>
								<td><%=cost_reg_view%></td>
								<td>
								<%
								If rs("cost_reg_yn") <> "Y" Then
								%>
									<a href="#" onClick="tax_esero_del('<%=page%>', '<%=rs("approve_no")%>', '<%=bill_id%>', '<%=bill_month%>', '<%=cost_reg_yn%>', '<%=end_yn%>');" style="cursor:pointer;">삭제</a>
								<%
								End If
								%>
								</td>
							</tr>
						<%
							rs.MoveNext()
						Loop

						rs.Close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="24%">
					<div class="btnCenter">
					<%If cost_record = 0 And total_record > 0 Then %>
						<a href="#" onClick="upload_cancel();" class="btnType04">업로드취소</a>
					<%End If %>
					</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
					%>
                    </td>
				    <td width="24%">
					<div class="btnCenter">
					<%If total_record > 0 And end_yn = "N" Then%>
						<a href="#" onClick="end_process();" class="btnType04">마감처리</a>
					<%End If%>

					<%If cost_record <> 0 And mi_record <> 0 Then%>
						<a href="#" onClick="end_process();" class="btnType04">부분마감처리</a>
					<%End If%>

					<%If cost_record = 0 And end_yn = "Y" Then%>
						<a href="#" onClick="cancel_process();" class="btnType04">마감처리취소</a>
					<%End If%>
					</div>
                    </td>
			      </tr>
				  </table>
				</form>
		</div>
	</div>
	</body>
</html>