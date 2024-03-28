<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim sum_amt(9)
Dim tot_amt(9)
Dim detail_tab(30)
Dim cost_amt(30,9)
Dim saupbu_tab(9)
Dim sales_amt(9)
Dim cost_tab

Dim cost_month, before_date, cost_year, cost_mm, c_month
Dim i
Dim rs_org, rsDeptSum
Dim title_line
Dim bi_saupbu

cost_tab = Array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

cost_month = Request.Form("cost_month")

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date), 6, 2)
End If

cost_year = Mid(cost_month, 1, 4)
cost_mm = Mid(cost_month, 5)
c_month = cost_year & "-" & cost_mm

For i = 0 To 8
	sum_amt(i) = 0
	tot_amt(i) = 0
	sales_amt(i) = 0
Next

i = 0

'sql = "select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
objBuilder.Append "SELECT saupbu "
objBuilder.Append "FROM sales_org "
objBuilder.Append "WHERE sales_year='" & cost_year & "' "
objBuilder.Append "ORDER BY sort_seq "

Set rs_org = Server.CreateObject("ADODB.RecordSet")
rs_org.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rs_org.EOF
	i = i + 1
	saupbu_tab(i) = rs_org(0)

	rs_org.MoveNext()
Loop

rs_org.Close()
Set rs_org = Nothing

i = i + 1
'saupbu_tab(i) = ""
'i = i + 1
'saupbu_tab(i) = "소계"

'sql = "select saupbu,sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&c_month&"' group by saupbu"
objBuilder.Append "SELECT saupbu, SUM(cost_amt) AS sales_amt "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(sales_date,1,7) = '"&c_month&"' "
objBuilder.Append "GROUP BY saupbu "

Set rsDeptSum = Server.CreateObject("ADODB.RecordSet")
rsDeptSum.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsDeptSum.EOF
	bi_saupbu = rsDeptSum("saupbu")

	If bi_saupbu = "기타사업부" Then
		bi_saupbu = ""
	End If

	For i = 1 To 7
		If saupbu_tab(i) = bi_saupbu Then
			sales_amt(i) = CCur(rsDeptSum("sales_amt"))
			sales_amt(8) = sales_amt(8) + CCur(rsDeptSum("sales_amt"))

			Exit For
		End If
	Next

	rsDeptSum.MoveNext()
Loop

rsDeptSum.Close()
Set rsDeptSum = Nothing

title_line = "사업부별 월별 손익 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
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

			function frmcheck(){
				if (chkfrm()) {
					document.frm.submit();
				}
			}

			function chkfrm(){
				if (document.frm.cost_month.value == "") {
					alert ("조회년월을 입력하세요.");
					return false;
				}
				return true;
			}

			function scrollAll(){
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
				<form action="saupbu_profit_loss_month.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>조회년월&nbsp;</strong>(예201401) :
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
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
							<col width="70px" >
							<col width="170px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">비용항목</th>
							  <th rowspan="2" scope="col">세부내역</th>
						<%For i = 1 To 6%>
							  <th scope="col"><%=saupbu_tab(i)%></th>
						<%Next%>
							  <th scope="col">기타사업부</th>
							  <th scope="col">소계</th>
                          </tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:470px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="70px" >
							<col width="170px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="120px" >
							<col width="*" >
						</colgroup>
						<tbody>
						<tr bgcolor="#FFFFCC">
							<td colspan="2" class="first" scope="col"><strong>매출</strong></td>
					<%For i = 1 To 8%>
                    		<td class="right" scope="col"><%=FormatNumber(sales_amt(i), 0)%></td>
 					<%Next%>
                         </tr>
					<%
					Dim jj, rec_cnt, j
					Dim rsCostAccount, rsDeptProfitLoss, rsCostSum

					For jj = 0 To 10
						rec_cnt = 0

						For i = 1 To 30
							detail_tab(i) = ""
							For j = 1 To 8
								cost_amt(i, j) = 0
								sum_amt(j) = 0
							Next
						Next

						If cost_tab(jj) = "인건비" Then
							'sql = "select cost_detail from saupbu_cost_account where cost_id ='"&cost_tab(jj)&"' order by view_seq"
							objBuilder.Append "SELECT cost_detail  "
							objBuilder.Append "FROM saupbu_cost_account "
							objBuilder.Append "WHERE cost_id ='"&cost_tab(jj)&"' "
							objBuilder.Append "ORDER BY view_seq "

							Set rsCostAccount = Server.CreateObject("ADODB.RecordSet")
							rsCostAccount.Open objBuilder.ToString(), DBConn, 1
							objBuilder.Clear()

							Do Until rsCostAccount.EOF
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rsCostAccount("cost_detail")

								rsCostAccount.MoveNext()
							Loop

							rsCostAccount.Close()
						  Else
							'sql = "select cost_detail from saupbu_profit_loss where (cost_year ='"&cost_year&"') and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
							objBuilder.Append "SELECT cost_detail "
							objBuilder.Append "FROM saupbu_profit_loss "
							objBuilder.Append "WHERE (cost_year ='"&cost_year&"') "
							objBuilder.Append "AND cost_id ='"&cost_tab(jj)&"' "
							objBuilder.Append "GROUP BY cost_detail "
							objBuilder.Append "ORDER BY cost_detail "

							Set rsDeptProfitLoss = Server.CreateObject("ADODB.RecordSet")
							rsDeptProfitLoss.Open objBuilder.ToString(), DBConn, 1
							objBuilder.Clear()

							Do Until rsDeptProfitLoss.EOF
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rsDeptProfitLoss("cost_detail")

								rsDeptProfitLoss.MoveNext()
							Loop

							rsDeptProfitLoss.Close()
						End If

						If rec_cnt <> 0 Then
							' 당월 금액 SUM
							'sql = "select saupbu,cost_detail,sum(cost_amt_"&cost_mm&") as cost from saupbu_profit_loss where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"' group by saupbu,cost_detail order by saupbu, cost_detail "
							objBuilder.Append "select saupbu,cost_detail,sum(cost_amt_"&cost_mm&") as cost "
							objBuilder.Append "from saupbu_profit_loss "
							objBuilder.Append "where cost_year ='"&cost_year&"' "
							objBuilder.Append "and cost_id ='"&cost_tab(jj)&"' "
							objBuilder.Append "group by saupbu,cost_detail "
							objBuilder.Append "order by saupbu, cost_detail "

							Set rsCostSum = Server.CreateObject("ADODB.RecordSet")
							rsCostSum.Open objBuilder.ToString(), DBConn, 1
							objBuilder.Clear()

							Do Until rsCostSum.eof
								For i = 1 To 30
									If rsCostSum("cost_detail") = detail_tab(i) Then
										For j = 1 To 7
											If saupbu_tab(j) = rsCostSum("saupbu") Then
												cost_amt(i,j) = cost_amt(i,j) + CDbl(rsCostSum("cost"))
												cost_amt(i,8) = cost_amt(i,8) + CDbl(rsCostSum("cost"))
												sum_amt(j) = sum_amt(j) + CDbl(rsCostSum("cost"))
												sum_amt(8) = sum_amt(8) + CDbl(rsCostSum("cost"))
												tot_amt(j) = tot_amt(j) + CDbl(rsCostSum("cost"))
												tot_amt(8) = tot_amt(8) + CDbl(rsCostSum("cost"))

												Exit For
											End If
										Next
									End If
								Next

								rsCostSum.MoveNext()
							Loop

							rsCostSum.Close()
						%>
							<tr>
							  	<td rowspan="<%=rec_cnt + 1%>" class="first">
						<%If jj = 2 Or jj = 3 Then %>
                        	  	<%=cost_tab(jj)%><br>(현금사용)
						<%Else	%>
                        	  	<%=cost_tab(jj)%>
                        <%End If%>
                              	</td>
								<td class="left"><%=detail_tab(1)%></td>
						<%For j = 1 To 8%>
								<td class="right"><%=FormatNumber(cost_amt(1, j), 0)%></td>
						<%Next	%>
						  </tr>
						<%For i = 2 To rec_cnt	%>
                        	<tr>
								<td class="left" style=" border-left:1px solid #e3e3e3;"><%=detail_tab(i)%></td>
						<%   For j = 1 To 8	%>
								<td class="right"><%=FormatNumber(cost_amt(i, j), 0)%></td>
						<%   Next	%>
							</tr>
						<%Next	%>
							<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
						<%For j = 1 To 8%>
								<td class="right" bgcolor="#EEFFFF"><%=FormatNumber(sum_amt(j), 0)%></td>
						<%Next	%>
						  </tr>
					<%
						End If
					Next

					Set rsCostAccount = Nothing
					Set rsDeptProfitLoss = Nothing
					Set rsCostSum = Nothing

					DBConn.Close()
					Set DBConn = Nothing
					%>
					<tr bgcolor="#FFFFCC">
							  <td colspan="2" class="first" scope="col"><strong>비용합계</strong></td>
						<%For j = 1 To 8%>
								<td class="right"><%=FormatNumber(tot_amt(j), 0)%></td>
						<%Next%>
                         </tr>
						<tr bgcolor="#FFDFDF">
							  <td colspan="2" bgcolor="#FFDFDF" class="first" scope="col"><strong>손익</strong></td>
						<%
						Dim cal_amt
						For j = 1 To 8
						 	cal_amt = sales_amt(j) - tot_amt(j)
						 %>
								<td class="right"><%=FormatNumber(cal_amt, 0)%></td>
						<%
						Next
						%>
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
                    <a href="saupbu_profit_loss_month_excel.asp?cost_month=<%=cost_month%>" class="btnType04">엑셀다운로드</a>
					</div>
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="profit_loss_detail_excel.asp?cost_month=<%=cost_month%>" class="btnType04">매입세금계산서다운로드</a>
					</div>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>
	</div>
	</body>
</html>

