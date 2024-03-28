<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim cost_month, sales_saupbu
Dim before_date
Dim condi_sql
Dim mm, cost_year

Dim costRs, rsAs
Dim sql
Dim tot_part_cost
Dim title_line

cost_month = Request.Form("cost_month")
sales_saupbu = Request.Form("sales_saupbu")

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date),1,4) + Mid(CStr(before_date),6,2)
	sales_saupbu = "전체"
End If

'검색 조건 쿼리
If sales_saupbu = "전체" Then
	condi_sql = ""
Else
  	condi_sql = " AND saupbu ='"&sales_saupbu&"'"
End If

mm = Mid(cost_month, 5, 2)
cost_year = Mid(cost_month, 1, 4)

objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS tot_cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "AND cost_center = '부문공통비' "

Set costRs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 전체 부문공통비
If IsNull(costRs("tot_cost")) Then
	tot_part_cost = 0
Else
	tot_part_cost = CLng(costRs("tot_cost"))
End If

costRs.Close()
Set costRs = Nothing

' 고객사별 AS 현황
sql = "  SELECT as_month                         "&chr(13)&_
      "       , as_company /* 고객사 */          "&chr(13)&_
      "       , saupbu     /* 사업부 */          "&chr(13)&_
      "       , as_cnt                           "&chr(13)&_
      "       , divide_amt_1                     "&chr(13)&_
      "       , divide_amt_2                     "&chr(13)&_
      "       , charge_per                       "&chr(13)&_
      "       , cost_amt   /* 부문공통비 */      "&chr(13)&_
      "       , reg_id                           "&chr(13)&_
      "       , reg_name                         "&chr(13)&_
      "       , reg_date                         "&chr(13)&_
      "    FROM company_asunit                   "&chr(13)&_
      "   WHERE as_month = '"&cost_month&"'      "&chr(13)&_
      "         "&condi_sql&"                    "&chr(13)&_
      "ORDER BY as_company                       "

Set rsAs = Server.CreateObject("ADODB.RecordSet")
rsAs.Open sql, Dbconn, 1

title_line = "부문공통비 AS 배부기준(표준단가)"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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

                <!-- <h3 class="stit">원격처리는 5%, 원격외는 95% 비중으로 적용한 배부기준입니다. </h3> -->
                <h3 class="stit">1. 1차배분금액은 AS건수 기준<br>
                2. 2차 배분금액은 과부족분에 대해 매출비중에 따라 배분한 금액<br>
                3. 부문공통비 적용 사업부 : SI1본부, SI2본부, NI본부, 공공본부</h3>

				<form action="part_cost_report_unit.asp" method="post" name="frm">
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
								Dim sql_org, rs_org

                                sql_org = "select saupbu "
								sql_org = sql_org & "from company_as "
								sql_org = sql_org & "where (saupbu <> '') and (as_month = '"&cost_month&"') "
								sql_org = sql_org & "group by saupbu "
								sql_org = sql_org & "order by saupbu asc "

								Set rs_org = Server.CreateObject("ADODB.RecordSet")
                                rs_org.Open sql_org, DBConn, 1
                                %>
                                <select name="sales_saupbu" id="sales_saupbu" style="width:150px">
                                    <option value="전체" <%If sales_saupbu = "전체" then %>selected<% end if %>>전체</option>
                                    <option value="" <%If sales_saupbu = "" then %>selected<% end if %>>미지정</option>
                                    <%
                                    do until rs_org.eof
                                        %>
                                        <option value='<%=rs_org("saupbu")%>' <%If rs_org("saupbu") = sales_saupbu  then %>selected<% end if %>><%=rs_org("saupbu")%></option>
                                        <%
                                        rs_org.movenext()
                                    loop
                                    rs_org.Close()
                                    %>
                                </select>
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
                                <col width="4%" >
                                <col width="*" >
                                <col width="15%" >
                                <col width="8%" >
                                <col width="10%" >
                                <col width="10%" >
                                <col width="14%" >
                                <col width="14%" >
                                <col width="3%" >
                            </colgroup>
                            <thead>
                                <tr>
                                    <th class="first" scope="col">순번</th>
                                    <th scope="col">회사</th>
                                    <th scope="col">사업부</th>
                                    <th scope="col">AS건수</th>
                                    <th scope="col">1차배부금액</th>
                                    <th scope="col">2차배부금액</th>
                                    <th scope="col">차지율(%)</th>
                                    <th scope="col">부문공통비</th>
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
							<col width="15%" >
							<col width="8%" >
                            <col width="10%" >
                            <col width="10%" >
							<col width="14%" >
							<col width="14%" >
							<col width="2%" >
						</colgroup>
						<tbody>
                            <%
							Dim as_sum, divide_amt_1_sum, divide_amt_2_sum, charge_per_sum, cost_amt_sum, i
							'Dim charge_cost

                            as_sum       = 0 ' AS건수 (ToBe)
                            divide_amt_1_sum  = 0
                            divide_amt_2_sum  = 0
                            charge_per_sum  = 0
                            cost_amt_sum = 0
                            i = 0

                            do until rsAs.eof
                                i = i + 1
                                'charge_cost     = int(rsAs("charge_per") * tot_part_cost)
                                as_sum          = CInt(rsAs("as_cnt"))+ as_sum ' AS건수 (ToBe)
								divide_amt_1_sum = divide_amt_1_sum + CLng(rsAs("divide_amt_1"))
								divide_amt_2_sum = divide_amt_2_sum + CLng(rsAs("divide_amt_2"))
                                charge_per_sum  = rsAs("charge_per")  + charge_per_sum
                                cost_amt_sum    = rsAs("cost_amt")    + cost_amt_sum
                                %>
                                <tr>
                                    <!-- 순번        --> <td class="first"><%=i%></td>
                                    <!-- 회사        --> <td><%=rsAs("as_company")%></td>
                                    <!-- 사업부      --> <td><%=rsAs("saupbu")%>&nbsp;</td>
                                    <!-- AS건수      --> <td class="right"><%=formatnumber(CInt(rsAs("as_cnt")),0)%>&nbsp;</td>
                                    <!-- 1차배부금액 --> <td class="right"><%=formatnumber(CLng(rsAs("divide_amt_1")),0)%>&nbsp;</td>
                                    <!-- 2차배부금액 --> <td class="right"><%=formatnumber(CLng(rsAs("divide_amt_2")),0)%>&nbsp;</td>
                                    <!-- 차지율(%)   --> <td class="right"><%=formatnumber(rsAs("charge_per"),5)%>&nbsp;%&nbsp;</td>
                                    <!-- 부문공통비  --> <td class="right"><%=formatnumber(rsAs("cost_amt"),0)%>&nbsp;</td>  <!-- (회사별사업부별)부문공통비 -->
                                    <td>&nbsp;</td>
                                </tr>
                                <%
                                rsAs.movenext()
                            Loop
                            %>
							<tr>
								<td colspan="2" bgcolor="#FFE8E8" class="first">총계</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(as_sum,0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" class="right"><%=formatnumber(divide_amt_1_sum,0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" class="right"><%=formatnumber(divide_amt_2_sum,0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" class="right"><%=formatnumber(charge_per_sum,5)%>&nbsp;%&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=formatnumber(cost_amt_sum,0)%>&nbsp;</td>
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
                            <a href="/part_cost_excel_unit.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">엑셀다운로드</a>
                            </div>
                        </td>
                        <td width="50%"></td>
                        <td width="25%"></td>
                    </tr>
                </table>
			    </form>
				<br>
		</div>
	</div>
	</body>
</html>
