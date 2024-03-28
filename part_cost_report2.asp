<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/common/func.asp" --><!--사용자 정의 함수 : 허정호_20201202-->
<%
'on Error resume next

'=========================================================
'사용하지 않는 레코드 소멸 코드 추가[허정호_20201203]
Set Rs_acc = Nothing
Set rs_trade = Nothing
Set rs_reside = Nothing
Set Rs_type = Nothing
Set rs_emp = Nothing
Set rs_etc = Nothing
Set rs_memb = Nothing
Set Rs_ddd = Nothing
Set RsCount = Nothing
Set Rs_in = Nothing
Set rs_hol = Nothing
Set rs_next = Nothing
Set rs_pre = Nothing
'=========================================================

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder

Set objBuilder = New StringBuilder
'===================================================

Dim from_date
Dim to_date
Dim win_sw

cost_month = Request.Form("cost_month")
sales_saupbu = Request.Form("sales_saupbu")

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date), 1, 4) + Mid(CStr(before_date), 6, 2)
	sales_saupbu = "전체"
End If

If sales_saupbu = "전체" Then
	condi_sql = ""
Else
  	condi_sql = " AND saupbu ='"&sales_saupbu&"'"
End If

mm = Mid(cost_month, 5, 2)
cost_year = Mid(cost_month, 1, 4)

'sql = "SELECT SUM(cost_amt_"&mm&") AS tot_cost FROM company_cost WHERE cost_year ='"&cost_year&"' AND cost_center = '부문공통비'"
'Set rs = DbConn.Execute(SQL)
objBuilder.Append "SELECT SUM(cost_amt_"& mm &") AS tot_cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_year ='"& cost_year &"' "
objBuilder.Append "AND cost_center = '부문공통비' "

Set rs = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs("tot_cost")) Then
	tot_part_cost = 0
Else
	tot_part_cost = clng(rs("tot_cost"))
End If

rs.close()

'sql = "SELECT * FROM company_as WHERE (as_month = '"&cost_month&"')"&condi_sql&" ORDER BY as_company"

sql = "SELECT											" & chr(13) &_
	  "	company											" & chr(13) &_
	  "	, saupbu										" & chr(13) &_
	  "	, count(acpt_no) AS remote_cnt					" & chr(13) &_
	  "	, sum(as_standard_money) AS cost_amt			" & chr(13) &_
	  "	, (sum(as_standard_money)/(SELECT sum(as_standard_money) FROM AS_ACPT WHERE 1=1 AND DATE_FORMAT( acpt_date, '%Y%m') = '"&cost_month&"' AND as_process = '완료' AND length(trim(saupbu)) > 0))*100  AS charge_per	" & chr(13) &_
	  "FROM AS_ACPT										" & chr(13) &_
	  "WHERE 1=1										" & chr(13) &_
	  "AND DATE_FORMAT( acpt_date, '%Y%m') = '"&cost_month&"'	" & chr(13) &_
	  "AND as_process = '완료'							" & chr(13) &_
	  "AND length(trim(saupbu)) > 0 "&condi_sql&"		" & chr(13) &_
	  "GROUP BY company, saupbu							" & chr(13) &_
	  "ORDER BY company ASC								"
rs.Open sql, Dbconn, 1

title_line = "부문공동비 AS 배분 기준 현황(변경후)"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				<h3 class="stit"> </h3>
				<form action="part_cost_report2.asp" method="post" name="frm">
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
                                <label>
								<strong>사업부 &nbsp;:</strong>
							<%
                                'sql_org="select saupbu from company_as where (saupbu <> '') and (as_month = '"&cost_month&"') group by saupbu order by saupbu asc"
                                'sql_org="SELECT DISTINCT saupbu FROM AS_ACPT WHERE 1=1 AND DATE_FORMAT( acpt_date, '%Y%m') = '"&cost_month&"' AND as_process = '완료' AND length(trim(saupbu)) > 0  order by saupbu asc"
                                'rs_org.Open sql_org, Dbconn, 1
								objBuilder.Append "SELECT DISTINCT saupbu "
								objBuilder.Append "FROM AS_ACPT "
								objBuilder.Append "WHERE DATE_FORMAT( acpt_date, '%Y%m') = '"& cost_month &"' "
								objBuilder.Append "AND as_process = '완료' "
								objBuilder.Append "AND LENGTH(TRIM(saupbu)) > 0 "
								objBuilder.Append "ORDER BY saupbu ASC "
								rs_org.Open objBuilder.ToString(), Dbconn, 1
								objBuilder.Clear()
                            %>
                                <select name="sales_saupbu" id="sales_saupbu" style="width:150px">
                                    <option value="전체" <%If sales_saupbu = "전체" then %>selected<% end if %>>전체</option>
                                    <option value="" <%If sales_saupbu = "" then %>selected<% end if %>>미지정</option>
							<%
                                do until rs_org.eof
                            %>
          							<option value='<%=rs_org("saupbu")%>' <%If rs_org("saupbu") = sales_saupbu  then %>selected<% end if %>><%=rs_org("saupbu")%></option>
							<%
                                    rs_org.MoveNext()
                                Loop

                                rs_org.Close()
								Set rs_org = Nothing
                            %>
                                </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
                    	<td>
      			<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
						<table cellpadding="0" cellspacing="0" style="word-break:break-all" class="tableList">
						<colgroup>
							<col width="50" >
							<col width="400" >
							<col width="200" >
							<col width="200" >
							<col width="100" >
							<col width="200" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">회사</th>
								<th scope="col">사업부</th>
								<th scope="col">건수</th>
								<th scope="col">차지율(%)</th>
								<th scope="col">부분공통비</th>
							</tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" style="word-break:break-all" class="scrollList">
						<colgroup>
							<col width="51" >
							<col width="406" >
							<col width="202" >
							<col width="203" >
							<col width="101" >
							<col width="187" >
						</colgroup>
						<tbody>
						<%
						remote_sum = 0
						charge_per_sum = 0
						charge_cost_sum = 0
						i = 0

						Do Until rs.EOF
							i = i + 1
							remote_sum = CInt(rs("remote_cnt")) + remote_sum
							charge_per_sum = CDbl(rs("charge_per")) + charge_per_sum
							charge_cost_sum = CLng(rs("cost_amt")) + charge_cost_sum
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("saupbu")%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rs("remote_cnt"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rs("charge_per"), 3)%>&nbsp;%&nbsp;</td>
								<td class="right"><%=FormatNumber(rs("cost_amt"), 0)%>&nbsp;</td>

							</tr>
						<%
							rs.MoveNext()
						Loop

						rs.Close()
						Set rs = Nothing

						DBConn.Close()
						Set DBConn = Nothing
						%>
							<tr>
								<td colspan="2" bgcolor="#FFE8E8" class="first">총계</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(remote_sum, 0)%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(charge_per_sum, 3)%>&nbsp;%&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(charge_cost_sum, 0)%>&nbsp;</td>
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
                    <a href="part_cost_excel2.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">엑셀다운로드</a>
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
