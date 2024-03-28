<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

stock_goods_code = request("stock_goods_code")
stock_goods_type = request("stock_goods_type")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
Set rs = DbConn.Execute(SQL)
if not rs.eof then
    	goods_code = rs("goods_code")
		goods_grade = rs("goods_grade")
        goods_gubun = rs("goods_gubun")
	    goods_name = rs("goods_name")
	    goods_standard = rs("goods_standard")
	    goods_type = rs("goods_type")
		goods_model = rs("goods_model")
		goods_serial_no = rs("goods_serial_no")
   else
		goods_code = ""
		goods_grade = ""
        goods_gubun = ""
	    goods_name = ""
	    goods_standard = ""
	    goods_type = ""
		goods_model = ""
		goods_serial_no = ""
end if
rs.close()

sql = "select * from met_stock_gmaster where (stock_goods_code = '"&stock_goods_code&"') and (stock_goods_type = '"&stock_goods_type&"') ORDER BY stock_company,stock_code ASC"
Rs.Open Sql, Dbconn, 1

title_line = "품목별 창고별 재고현황"

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
				a=confirm('출고를 취소하겠습니까?')
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
                                <th>품목코드</th>
							    <td class="left"><%=goods_code%>&nbsp;</td>
							    <th>품목명</th>
							    <td class="left"><%=goods_name%>&nbsp;</td>
							    <th>상태</th>
							    <td class="left"><%=goods_grade%>&nbsp;</td>
 							</tr>
                            <tr>
							    <th>품목구분</th>
							    <td class="left"><%=goods_gubun%>&nbsp;</td>
							    <th>규격</th>
							    <td class="left"><%=goods_standard%>&nbsp;</td>
                                <th>모델</th>
							    <td class="left"><%=goods_model%>&nbsp;</td>
						    </tr>
                            <tr>
                                <th>Serial No.</th>
							    <td class="left" colspan="5"><%=goods_serial_no%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 창고별 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
                            <col width="*" >
							<col width="8%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col">회사</th>
                                <th scope="col">창고명</th>
                                <th scope="col">용도구분</th>
                                <th scope="col">전년이월</th>
                                <th scope="col">입고수량</th>
                                <th scope="col">출고수량</th>
								<th scope="col">현재고</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>     
						<%
							i = 0
							h_last_qty = 0
							h_in_qty = 0
							h_go_qty = 0
							h_JJ_qty = 0
							do until rs.eof or rs.bof
							     i = i + 1
							
							     if rs("stock_JJ_qty") > 0 then
								     h_last_qty = h_last_qty + rs("stock_last_qty")
									 h_in_qty = h_in_qty + rs("stock_in_qty")
									 h_go_qty = h_go_qty + rs("stock_go_qty")
									 h_JJ_qty = h_JJ_qty + rs("stock_JJ_qty")
						%>
							<tr>
                                <td><%=rs("stock_company")%>&nbsp;</td>
                                <td><%=rs("stock_name")%>&nbsp;</td>
                                <td><%=rs("stock_goods_type")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stock_last_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stock_JJ_qty"),0)%>&nbsp;</td>
								<td><%=rs("stock_level")%>&nbsp;</td>
							</tr>
						<%
								end if
								rs.movenext()
							loop
							rs.close()
						%>
                            <tr>
                                <td colspan="3" style="background:#ffe8e8;">총 계</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_last_qty,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_in_qty,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_go_qty,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_JJ_qty,0)%>&nbsp;</td>
								<td style="background:#ffe8e8;">&nbsp;</td>
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
                    <input type="hidden" name="stock_goods_code" value="<%=stock_goods_code%>">
					<input type="hidden" name="stock_goods_type" value="<%=stock_goods_type%>">
	     </form>
    	</div>				
	  </div>     
	</body>
</html>

