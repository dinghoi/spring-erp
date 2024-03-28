<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

stin_order_no = request("stin_order_no")
stin_in_date = request("stin_in_date")
stin_order_seq = request("stin_order_seq")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_stin where (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"') and (stin_in_date = '"&stin_in_date&"')"
Set Rs_order = DbConn.Execute(SQL)
if not Rs_order.eof then
    	stin_in_date = Rs_order("stin_in_date")
		
		stin_order_no = Rs_order("stin_order_no")
		stin_order_seq = Rs_order("stin_order_seq")
		
		stin_id = Rs_order("stin_id")
		stin_buy_company = Rs_order("stin_buy_company")
		stin_buy_bonbu = Rs_order("stin_buy_bonbu")
		stin_buy_saupbu = Rs_order("stin_buy_saupbu")
		
		stin_goods_type = Rs_order("stin_goods_type")
		
	    'stin_trade_no = Rs_order("stin_trade_no")
		stin_trade_no = mid(Rs_order("stin_trade_no"),1,3) + "-" + mid(Rs_order("stin_trade_no"),4,2) + "-" + mid(Rs_order("stin_trade_no"),6)
        stin_trade_name = Rs_order("stin_trade_name")
        stin_trade_person = Rs_order("stin_trade_person")
		stin_trade_email = Rs_order("stin_trade_email")
		
        stin_stock_company = Rs_order("stin_stock_company")
        stin_stock_code = Rs_order("stin_stock_code")
        stin_stock_name = Rs_order("stin_stock_name")
		
        stin_price = Rs_order("stin_price")
        stin_cost = Rs_order("stin_cost")
        stin_cost_vat = Rs_order("stin_cost_vat")
		
		stin_company = Rs_order("stin_company")
        stin_emp_no = Rs_order("stin_emp_no")
        stin_emp_name = Rs_order("stin_emp_name")
		
		stin_att_file = Rs_order("stin_att_file")
		stin_memo = Rs_order("stin_memo")
		
		po_date = Rs_order("po_date")
		po_number = Rs_order("po_number")
		park_bl = Rs_order("park_bl")
		
		won_ex = Rs_order("won_ex")
		tong_cost = Rs_order("tong_cost")
		stock_cost = Rs_order("stock_cost")
		trans_cost = Rs_order("trans_cost")
		air_cost = Rs_order("air_cost")
		inland_cost = Rs_order("inland_cost")
		
   else
		stin_in_date = ""
		
		stin_order_no = ""
		stin_order_seq = ""
		
		stin_id = ""
		stin_buy_company = ""
		stin_buy_bonbu = ""
		stin_buy_saupbu = ""
		
		stin_goods_type = ""
		
	    stin_trade_no = ""
        stin_trade_name = ""
        stin_trade_person = ""
		stin_trade_email = ""
		
        stin_stock_company = ""
        stin_stock_code = ""
        stin_stock_name = ""
		
        stin_price = 0
        stin_cost = 0
        stin_cost_vat = 0
		
		stin_company = ""
        stin_emp_no = ""
        stin_emp_name = ""
		
		stin_att_file = ""
		stin_memo = ""

end if
Rs_order.close()

sql = "select * from met_stin_goods where (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"') and (stin_date = '"&stin_in_date&"') ORDER BY stin_goods_seq,stin_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "외자입고 상세조회"

view_att_file = stin_att_file
path = "/met_upload"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>자재관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}		
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}					
			function chkfrm() {
						
				{
				a=confirm('발주를 취소하겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = false; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 10; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 10; //오른쯕 여백 설정
                factory.printing.bottomMargin = 15; //바닦 여백 설정
        //		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
        //		factory.printing.printer = ""; //프린터 할 프린터 이름
        //		factory.printing.paperSize = "A4"; //용지선택
        //		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
        //		factory.printing.collate = true; //순서대로 출력하기
        //		factory.printing.copies = "1"; //인쇄할 매수
        //		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
        //		factory.printing.Printer(true); //출력하기
                factory.printing.Preview(); //윈도우를 통해서 출력
                factory.printing.Print(false); //윈도우를 통해서 출력
            }

		</script>

	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
				<form method="post" name="frm" action="met_buy_cancel.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="10%" >
							<col width="15%" >
							<col width="10%" >
							<col width="15%" >
							<col width="10%" >
							<col width="15%" >
							<col width="10%" >
							<col width="15%" >
						</colgroup>
						<tbody> 
							<tr>
                                <th>구매용도</th>
							    <td class="left"><%=stin_goods_type%>&nbsp;</td>
							    <th>구매그룹사</th>
							    <td class="left"><%=stin_buy_company%>&nbsp;</td>
							    <th>구매사업부</th>
							    <td class="left"><%=stin_buy_saupbu%>&nbsp;</td>
                                <th>입고담당</th>
							    <td class="left"><%=stin_emp_name%>(<%=stin_emp_no%>)&nbsp;</td>
 							</tr>
                            <tr>
                                <th>입고일</th>
							    <td class="left"><%=stin_in_date%>&nbsp;</td>
                                <th>입고번호</th>
                                <td class="left"><%=stin_order_no%>&nbsp;<%=stin_order_seq%></td>
                                <th>입고창고</th>
							    <td class="left"><%=stin_stock_name%>&nbsp;</td>
                                <th>구매거래처</th>
							    <td class="left"><%=stin_trade_name%>&nbsp;</td>
						    </tr>
                            <tr>
                                <th>PO 일자</th>
							    <td class="left"><%=po_date%>&nbsp;</td>
                                <th>PO_Number</th>
							    <td class="left"><%=po_number%>&nbsp;</td>
                                <th>Parking Number</th>
							    <td colspan="3" class="left"><%=park_bl%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th>비고</th>
							  <td colspan="7" class="left"><%=stin_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 입고 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="7%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="14%" >
							<col width="10%" >
							<col width="7%" >
                            <col width="7%" >
							<col width="8%" >
                            <col width="6%" >
							<col width="8%" >
                            <col width="6%" >
                            <col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">Part_No.</th>
								<th scope="col">입고수량</th>
                                <th scope="col">단가($)</th>
								<th scope="col">단가(원)</th>
                                <th scope="col">비용단가</th>
								<th scope="col">입고금액</th>
                                <th scope="col">환율</th>
                                <th scope="col">Serial_No</th>
							</tr>
						</thead>
						<tbody>     
						<%
							buy_cost_tot = 0
							i = 0
							do until rs.eof or rs.bof
							     i = i + 1
							
							     unit_wonga = rs("d_cost") * rs("w_won")
								 
								 buy_hap = rs("stin_qty") * rs("stin_unit_cost")
							     buy_cost_tot = buy_cost_tot + buy_hap
								 
							
						%>
							<tr>
								<td class="first"><%=i%></td>
                                <td><%=rs("stin_goods_type")%>&nbsp;</td>
								<td><%=rs("stin_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("stin_goods_code")%>&nbsp;</td>
                                <td><%=rs("stin_goods_name")%>&nbsp;</td>
                                <td><%=rs("part_number")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stin_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("d_cost"),2)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(unit_wonga,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("ex_cost"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(buy_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("w_won"),2)%>&nbsp;</td>
                                <td>&nbsp;
                            <%
                                    if rs("excel_att_file") <> "" then		
                            %>
                                        <a href="att_file_download.asp?path=<%=path%>&att_file=<%=rs("excel_att_file")%>">첨부</a>&nbsp;
                            <%      end if    %>
                                
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
						</tbody>
					</table>
                    <br>
                    <table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="*" >
						</colgroup>
						<tbody>
                        <% 
						    buy_vat_hap = int(buy_cost_tot * (10 / 100))
							buy_tot_price = buy_cost_tot + buy_vat_hap
						%>
							<tr>
							  <th>입고총액</th>
							  <td style="text-align:right"><%=formatnumber(buy_tot_price,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							  <th>입고금액</th>
							  <td style="text-align:right"><%=formatnumber(buy_cost_tot,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							  <th>부가세</th>
							  <td style="text-align:right"><%=formatnumber(buy_vat_hap,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						    </tr>
                            <tr>
							  <th>적용환율</th>
							  <td style="text-align:right"><%=formatnumber(won_ex,2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							  <th>통관수수료</th>
							  <td style="text-align:right"><%=formatnumber(tong_cost,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							  <th>창고료</th>
							  <td style="text-align:right"><%=formatnumber(stock_cost,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						    </tr>
                            <tr>
							  <th>운송료</th>
							  <td style="text-align:right"><%=formatnumber(trans_cost,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							  <th>항공운임료</th>
							  <td style="text-align:right"><%=formatnumber(air_cost,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							  <th>내륙운송료</th>
							  <td style="text-align:right"><%=formatnumber(inland_cost,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						    </tr>
                            <tr>
							  <th>첨부</th>
							  <td colspan="5" class="left">
                        <% 
                           If stin_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=stin_att_file%>"><%=stin_att_file%></a>
                        <%    Else %>
				                    &nbsp;
                        <% 
						    End If %>
                              </td>
						    </tr>
						</tbody>
					</table>
          	     <br>
     				<div class="noprint">
                        <div align=center>
                            <span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>
                            <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                        </div>
					</div>
					<br>               		
                    <input type="hidden" name="order_no" value="<%=stin_order_no%>">
					<input type="hidden" name="order_seq" value="<%=stin_order_seq%>">
					<input type="hidden" name="order_date" value="<%=stin_in_date%>">
	     </form>
    	</div>				
	  </div>     
	</body>
</html>

