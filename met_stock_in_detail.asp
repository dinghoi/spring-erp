<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

order_no = request("order_no")
stin_in_date = request("stin_in_date")
order_seq = request("order_seq")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_stin where (stin_order_no = '"&order_no&"') and (stin_order_seq = '"&order_seq&"') and (stin_in_date = '"&stin_in_date&"')"
Set Rs_order = DbConn.Execute(SQL)
if not Rs_order.eof then
    	order_in_date = Rs_order("stin_in_date")
		
		order_no = Rs_order("stin_order_no")
		order_seq = Rs_order("stin_order_seq")
		order_date = Rs_order("stin_order_date")
		order_buy_no = Rs_order("stin_buy_no")
		order_buy_seq = Rs_order("stin_buy_seq")
		order_buy_date = Rs_order("stin_buy_date")
		
		order_goods_type = Rs_order("stin_goods_type")
		
	    order_bill_collect = Rs_order("stin_bill_collect")
        order_collect_due_date = Rs_order("stin_collect_due_date")
	    order_trade_no = Rs_order("stin_trade_no")
        order_trade_name = Rs_order("stin_trade_name")
        order_trade_person = Rs_order("stin_trade_person")
		order_trade_email = Rs_order("stin_trade_email")
		
        order_stock_company = Rs_order("stin_stock_company")
        order_stock_code = Rs_order("stin_stock_code")
        order_stock_name = Rs_order("stin_stock_name")
		
        order_price = Rs_order("stin_price")
        order_cost = Rs_order("stin_cost")
        order_cost_vat = Rs_order("stin_cost_vat")
		
	    if order_collect_due_date = "0000-00-00" then
	          order_collect_due_date = ""
	    end if

   else
		order_buy_no = ""
		order_buy_seq = ""
		order_buy_date = ""
		order_goods_type = ""

	    order_bill_collect = ""
        order_collect_due_date = ""
	    order_trade_no = ""
        order_trade_name = ""
        order_trade_person = ""
		order_trade_email = ""

		order_in_date = ""
        order_stock_company = ""
        order_stock_code = ""
        order_stock_name = ""
        order_price = 0
        order_cost = 0
        order_cost_vat = 0

end if
Rs_order.close()

sql = "select * from met_stin_goods where (stin_order_no = '"&order_no&"') and (stin_order_seq = '"&order_seq&"') and (stin_date = '"&order_in_date&"') ORDER BY stin_goods_seq,stin_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "입고 상세조회"

buy_att_file = ""
view_att_file = buy_att_file
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
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 13; //오른쯕 여백 설정
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
//			function approve_request(slip_id,slip_no,slip_seq) 
			function approve_request() 
				{
				a=confirm('결재 요청하시겠습니까?')
				if (a==true) {
//					document.frm.action = "met_buy_approve_ok.asp?slip_id="+slip_id+'&slip_no='+slip_no+'&slip_seq='+slip_seq;
					document.frm.action = "met_buy_approve_ok.asp";
					document.frm.submit();
				}
				return false;
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
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="*" >
						</colgroup>
						<tbody> 
							<tr>
                                <th>구매번호</th>
							    <td class="left"><%=order_buy_no%>&nbsp;<%=order_buy_seq%></td>
							    <th>구매유형</th>
							    <td class="left"><%=order_goods_type%>&nbsp;</td>
							    <th>구매일자</th>
							    <td class="left"><%=order_buy_date%></td>
 							</tr>
                            <tr>
							    <th>발주일자</th>
							    <td class="left"><%=order_date%></td>
							    <th>발주번호</th>
							    <td colspan="3" class="left"><%=order_no%>&nbsp;<%=order_seq%></td>
						    </tr>
							<tr>
                                <th>발주처</th>
							    <td class="left"><%=order_trade_name%></td>
							    <th>사업자번호</th>
							    <td class="left"><%=order_trade_no%></td>
							    <th>담당자</th>
							    <td class="left"><%=order_trade_person%></td>
						    </tr>
                            <tr>
                                <th>이메일</th>
							    <td class="left"><%=order_trade_email%></td>
							    <th>대금<br>지급방법</th>
							    <td class="left"><%=order_bill_collect%></td>
							    <th>지급예정일</th>
							    <td class="left"><%=order_collect_due_date%></td>
						    </tr>
                            <tr>
                                <th>입고일</th>
							    <td class="left"><%=order_in_date%></td>
                                <th>입고창고</th>
							    <td colspan="3" class="left"><%=order_stock_name%>&nbsp;(<%=order_stock_company%>)</td>

						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 입고 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="14%" >
							<col width="16%" >

							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
                                <th scope="col">입고수량</th>
                                <th scope="col">입고단가</th>
								<th scope="col">입고금액</th>
							</tr>
						</thead>
						<tbody>     
						<%
							buy_cost_tot = 0
							i = 0
							do until rs.eof or rs.bof
							     i = i + 1
							
							     buy_hap = rs("stin_qty") * rs("stin_unit_cost")
							     buy_cost_tot = buy_cost_tot + buy_hap
							
						%>
							<tr>
								<td class="first"><%=i%></td>
                                <td><%=rs("stin_goods_type")%>&nbsp;</td>
								<td><%=rs("stin_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("stin_goods_code")%>&nbsp;</td>
                                <td><%=rs("stin_goods_name")%>&nbsp;</td>
                                <td><%=rs("stin_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stin_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stin_unit_cost"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(buy_hap,0)%>&nbsp;</td>
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
							  <td class="right"><%=formatnumber(buy_tot_price,0)%></td>
							  <th>입고금액</th>
							  <td class="right"><%=formatnumber(buy_cost_tot,0)%></td>
							  <th>부가세</th>
							  <td class="right"><%=formatnumber(buy_vat_hap,0)%></td>
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
                    <input type="hidden" name="order_no" value="<%=order_no%>">
					<input type="hidden" name="order_seq" value="<%=order_seq%>">
					<input type="hidden" name="order_date" value="<%=order_date%>">
	     </form>
    	</div>				
	  </div>     
	</body>
</html>

