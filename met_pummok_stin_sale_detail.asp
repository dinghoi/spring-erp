<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

stock_goods_code = request("stock_goods_code")
stock_goods_type = request("stock_goods_type")
stock_code = request("stock_code")
stock_name = request("stock_name")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_mvin = Server.CreateObject("ADODB.Recordset")
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

sql = "select * from met_stin_goods where (stin_goods_code = '"&stock_goods_code&"') and (stin_goods_type = '"&stock_goods_type&"') and (stin_stock_code = '"&stock_code&"') ORDER BY stin_date DESC"

'sql = "select * from met_mv_in_goods where (in_goods_code = '"&stock_goods_code&"') and (in_goods_type = '"&stock_goods_type&"') and (mvin_in_stock = '"&stock_code&"') ORDER BY mvin_in_date DESC"
Rs.Open Sql, Dbconn, 1

title_line = goods_name + " 품목 < " + stock_name + " >창고 입고현황"

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
							    <th>창고</th>
							    <td class="left" colspan="3"><%=stock_name%>&nbsp;(<%=stock_code%>)</td>
                                <th>품목구분</th>
							    <td class="left"><%=goods_gubun%>&nbsp;</td>
						    </tr>
                            <tr>
                                <th>품목코드</th>
							    <td class="left"><%=goods_code%>&nbsp;</td>
							    <th>품목명</th>
							    <td class="left"><%=goods_name%>&nbsp;</td>
							    <th>상태</th>
							    <td class="left"><%=goods_grade%>&nbsp;</td>
 							</tr>
                            <tr>
							    <th>규격</th>
							    <td class="left"><%=goods_standard%>&nbsp;</td>
                                <th>모델</th>
							    <td class="left"><%=goods_model%>&nbsp;</td>
                                <th>Serial No.</th>
							    <td class="left"><%=goods_serial_no%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 일자별 입고 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="10%" >
                            <col width="15%" >
                            <col width="15%" >
                            <col width="10%" >
                            <col width="12%" >
                            <col width="10%" >
                            <col width="*" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col">입고일자</th>
                                <th scope="col">용도구분</th>
                                <th scope="col">입고구분</th>
                                <th scope="col">입고번호</th>
                                <th scope="col" class="right">입고수량</th>
                                <th scope="col" class="right">입고금액</th>
                                <th scope="col" class="right">출고수량</th>
                                <th scope="col">비고</th>
                                <th scope="col">출고등록</th>
							</tr>
						</thead>
						<tbody>     
						<%
							i = 0
							h_in_qty = 0
							h_in_amt = 0
							h_cg_qty = 0
                            do until Rs.eof or rs.bof
							   if Rs("stin_qty") > 0 then
									 h_in_qty = h_in_qty + Rs("stin_qty")
									 h_in_amt = h_in_amt + Rs("stin_amt")
									 h_cg_qty = h_cg_qty + Rs("cg_qty")

									 stin_id = Rs("stin_id") + "입고"
									 jj_qty = 0
									 jj_qty = Rs("stin_qty") - Rs("cg_qty")
						%>
							<tr>
                                <td><%=Rs("stin_date")%>&nbsp;</td>
                                <td><%=Rs("stin_goods_type")%>&nbsp;</td>
                                <td><%=stin_id%>&nbsp;</td>
                                <td><%=rs("stin_order_no")%>&nbsp;<%=rs("stin_order_seq")%></td>
                                <td class="right"><%=formatnumber(Rs("stin_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(Rs("stin_amt"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(Rs("cg_qty"),0)%>&nbsp;</td>
								<td>&nbsp;</td>
                                <td>
						<% if jj_qty > 0 then %>                                
                                <a href="#" onClick="pop_Window('met_chulgo_sale_add03.asp?stin_date=<%=Rs("stin_date")%>&stin_order_no=<%=Rs("stin_order_no")%>&stin_order_seq=<%=Rs("stin_order_seq")%>&stin_goods_seq=<%=Rs("stin_goods_seq")%>&stin_goods_type=<%=Rs("stin_goods_type")%>&stin_goods_code=<%=Rs("stin_goods_code")%>&stin_stock_code=<%=Rs("stin_stock_code")%>&u_type=<%=""%>','met_chulgo_sale_add03_pop','scrollbars=yes,width=1230,height=650')">영업</a>
						<%    else   %>     
                                 -  
						<% end if %>                                                               
                                </td>
							</tr>
						<%
								end if
								Rs.movenext()
							loop
							Rs.close()   
						%>							                    
                            <tr>
                                <td colspan="4" style="background:#ffe8e8;">총 계</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_in_qty,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_in_amt,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_cg_qty,0)%>&nbsp;</td>
								<td colspan="2" style="background:#ffe8e8;">&nbsp;</td>
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

