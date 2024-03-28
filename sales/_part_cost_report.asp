<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'on Error resume next
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
Dim cost_month, sales_saupbu, before_date
Dim condi_sql, mm, cost_year
Dim rsComCost, tot_part_cost
Dim from_date, end_date, to_date
Dim rsAsTot, tot_part_cnt
Dim title_line

cost_month = Request.Form("cost_month")
sales_saupbu = Request.Form("sales_saupbu")

If sales_saupbu = "" Then
	sales_saupbu = "전체"
End If

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date),6,2)
	sales_saupbu = "전체"
End If

from_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

mm = Mid(cost_month, 5, 2)
cost_year = Mid(cost_month, 1, 4)

'부문공통비 전체 비용
'sql = "SELECT SUM(cost_amt_"&mm&") AS tot_cost FROM company_cost WHERE cost_year ='"&cost_year&"' AND cost_center = '부문공통비'"
'Set rs = DbConn.Execute(SQL)
objBuilder.Append "SELECT SUM(cost_amt_"& mm &") AS tot_cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_year ='"& cost_year &"' "
objBuilder.Append "AND cost_center = '부문공통비' "

Set rsComCost = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsComCost("tot_cost")) Then
	tot_part_cost = 0
Else
	tot_part_cost = CLng(rsComCost("tot_cost"))
End If

rsComCost.Close() : Set rsComCost = Nothing

If sales_saupbu = "전체" Then
	condi_sql = ""
Else
  	condi_sql = " AND trat.saupbu ='"& sales_saupbu &"' "
End If

'A/S 전체 카운트
objBuilder.Append "SELECT COUNT(*) AS tot_cnt "
objBuilder.Append "FROM as_acpt_end AS asat "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON asat.mg_ce_id = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&cost_month&"' "
objBuilder.Append "INNER JOIN trade AS trat ON asat.company = trat.trade_name "
objBuilder.Append "WHERE asat.as_type NOT IN ('원격처리', '야특근')"
objBuilder.Append "	AND asat.as_process <> '취소'"
objBuilder.Append "	AND asat.reside = '0'"
objBuilder.Append "	AND asat.reside_place = ''"
objBuilder.Append "	AND (CAST(asat.visit_date AS DATE) >= '"&from_date&"' AND CAST(asat.visit_date AS DATE) <= '"&to_date&"') "
objBuilder.Append "	AND emmt.cost_center = '부문공통비' "
objBuilder.Append condi_sql

Set rsAsTot = DBconn.Execute(objBuilder.ToString())
objBuilder.Clear()

tot_part_cnt = rsAsTot("tot_cnt")

rsAsTot.Close() : Set rsAsTot = Nothing

'A/S 장애처리 현황
'sql = "SELECT * FROM company_as WHERE (as_month = '"&cost_month&"')"&condi_sql&" ORDER BY as_company"
objBuilder.Append "SELECT company, bonbu, "
objBuilder.Append "	SUM(IF(as_type = '기타' OR as_type = '방문처리', cnt, 0)) AS 'fault', "
objBuilder.Append "	SUM(IF(as_type = '신규설치' OR as_type = '신규설치공사' OR as_type = '이전설치' "
objBuilder.Append "		OR as_type = '이전설치공사' OR as_type = '랜공사' OR as_type = '이전랜공사', cnt, 0)) AS 'setting', "
objBuilder.Append "	SUM(IF(as_type = '예방점검', cnt, 0)) AS 'testing', "
objBuilder.Append "	SUM(IF(as_type = '장비회수', cnt, 0)) AS 'collect', "
objBuilder.Append	tot_part_cost&" / "&tot_part_cnt&" * SUM(cnt) AS 'as_cost' "	'/*부문공통비 전체 비용 / as 전체 건수 * 사이트별 AS 건수*/
objBuilder.Append "FROM ( "
objBuilder.Append "	SELECT asat.company, trat.saupbu AS bonbu, as_type, COUNT(*) AS cnt, SUM(as_standard_money) AS std_cost "
objBuilder.Append "	FROM as_acpt_end AS asat "
objBuilder.Append "	INNER JOIN emp_master_month AS emmt ON asat.mg_ce_id = emmt.emp_no "
objBuilder.Append "		AND emmt.emp_month = '"&cost_month&"' "
objBuilder.Append "	INNER JOIN trade AS trat ON asat.company = trat.trade_name "
objBuilder.Append "	WHERE asat.as_type NOT IN ('원격처리', '야특근') "
objBuilder.Append "		AND asat.as_process <> '취소' "
objBuilder.Append "		AND asat.reside = '0' "
objBuilder.Append "		AND asat.reside_place = '' "
objBuilder.Append "		AND (CAST(asat.visit_date as date) >= '"&from_date&"' AND CAST(asat.visit_date as date) <= '"&to_date&"') "
objBuilder.Append "		AND emmt.cost_center = '부문공통비' "
objBuilder.Append condi_sql
objBuilder.Append "	GROUP BY asat.company, as_type "
objBuilder.Append ") r1 "
objBuilder.Append "GROUP BY company "

Response.write objBuilder.ToString()

Set rsComCost = Server.CreateObject("ADODB.RecordSet")
rsComCost.Open objBuilder.ToString(), Dbconn, 1
objBuilder.Clear()

title_line = "A/S 장애처리 현황"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<!-- <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" /> -->
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("발생년월을 입력하세요.");
					return false;
				}
				return true;
			}

			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<!--<h3 class="stit">원격처리는 5%, 원격외는 95% 비중으로 적용한 배부기준입니다. </h3>-->
				<form action="/sales/_part_cost_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>발생년월&nbsp;</strong>(예201401) :
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>

                                <!--<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>-->
								<img src="/image/but_ser.jpg" onclick="frmcheck();" style="cursor:pointer;" alt="검색">
                            </p>
						</dd>
					</dl>
				</fieldset>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
                    	<td>
      			<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">거래처명</th>
								<th scope="col">장애</th>
								<th scope="col">공사/설치</th>
								<th scope="col">예방점검</th>
								<th scope="col">장비회수</th>
								<th scope="col">관리본부</th>
								<th scope="col">부분공통비</th>
								<th scope="col"></th>
							</tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
							<col width="2%" >
						</colgroup>
						<tbody>
						<%
						Dim i, fault_sum, setting_sum, testing_sum, collect_sum, as_cost_sum

						fault_sum = 0
						setting_sum = 0
						testing_sum = 0
						collect_sum = 0
						as_cost_sum = 0
						i = 0

						Do Until rsComCost.EOF
							i = i + 1

							fault_sum = fault_sum + CLng(rsComCost("fault"))
							setting_sum = setting_sum + CLng(rsComCost("setting"))
							testing_sum = testing_sum + CLng(rsComCost("testing"))
							collect_sum = collect_sum + CLng(rsComCost("collect"))
							as_cost_sum = as_cost_sum + CDbl(rsComCost("as_cost"))
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rsComCost("company")%></td>
								<td class="right"><%=FormatNumber(rsComCost("fault"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("setting"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("testing"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("collect"), 0)%>&nbsp;</td>
								<td><%=rsComCost("bonbu")%></td>
								<td class="right"><%=FormatNumber(rsComCost("as_cost"), 0)%>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
							rsComCost.MoveNext()
						Loop

						rsComCost.Close() : Set rsComCost = Nothing

						DBConn.Close() : Set DBConn = Nothing
						%>
							<tr>
								<td colspan="2" bgcolor="#FFE8E8" class="first">총계</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(fault_sum, 0)%>&nbsp;건</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(setting_sum, 0)%>&nbsp;건</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(testing_sum, 0)%>&nbsp;건</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(collect_sum, 0)%>&nbsp;건</td>
								<td colspan="2" bgcolor="#FFE8E8" class="right"><%=FormatNumber(as_cost_sum, 0)%>&nbsp;</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
							</tr>
						</tbody>
						</table>
                        </DIV>
						</td>
                    </tr>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="/sales/part_cost_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">엑셀다운로드</a>
					</div>
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				  </table>
			</form>
				<br>
		</div>
	</div>
	</body>
</html>
