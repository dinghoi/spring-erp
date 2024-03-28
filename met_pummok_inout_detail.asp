<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim rs
Dim rs_numRows

stock_goods_code = request("stock_goods_code")
stock_code = request("stock_code")
stock_company = request("stock_company")
stock_name = request("stock_name")
stock_goods_type = request("stock_goods_type")


Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_jae = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect


sql = " delete from met_stock_inout where (stock_goods_code = '"&stock_goods_code&"') and (stock_code = '"&stock_code&"') and (stock_goods_type = '"&stock_goods_type&"')" 	
dbconn.execute(sql)

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

jjj = 0

'구매입고
sql = "select * from met_stin_goods where (stin_goods_code = '"&stock_goods_code&"') and (stin_goods_type = '"&stock_goods_type&"') and (stin_stock_code = '"&stock_code&"')"
Rs.Open Sql, Dbconn, 1

do until rs.eof
    id_seq = "1"

    jjj = jjj + 1
    inout_number = right(("00000" + cstr(jjj)),5)

    sql="insert into met_stock_inout (stock_code,stock_goods_type,stock_goods_code,stock_date,id_seq,inout_number,stock_company,stock_name,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_grade,stock_last_qty,stock_in_qty,stock_go_qty,stock_jj_qty,stock_id,inout_no,inout_seq) values ('"&rs("stin_stock_code")&"','"&rs("stin_goods_type")&"','"&rs("stin_goods_code")&"','"&rs("stin_date")&"','"&id_seq&"','"&inout_number&"','"&rs("stin_stock_company")&"','"&rs("stin_stock_name")&"','"&rs("stin_goods_gubun")&"','"&rs("stin_goods_name")&"','"&rs("stin_standard")&"','"&goods_grade&"',0,'"&rs("stin_qty")&"',0,0,'"&rs("stin_id")&"','"&rs("stin_order_no")&"','"&rs("stin_order_seq")&"')"
	
	dbconn.execute(sql)

	rs.movenext()
loop
rs.close()		

' 본사출고건 입고잡은것, 창고이동 입고건 ..... 입고현황에 나오게 할것	
sql = "select * from met_mv_in_goods where (in_goods_code = '"&stock_goods_code&"') and (in_goods_type = '"&stock_goods_type&"') and (mvin_in_stock = '"&stock_code&"')"
Rs.Open Sql, Dbconn, 1

do until Rs.eof or rs.bof
   if Rs("in_qty") > 0 then
		 mvin_in_date = rs("mvin_in_date")
	     yymmdd = mid(cstr(mvin_in_date),3,2) + mid(cstr(mvin_in_date),6,2)  + mid(cstr(mvin_in_date),9,2)
	     rele_no = yymmdd + rs("mvin_in_stock")
		 
		 stin_id = Rs("mvin_id") + "입고"
         id_seq = "2"
		 
		 jjj = jjj + 1
         inout_number = right(("00000" + cstr(jjj)),5)
		 
		 sql="insert into met_stock_inout (stock_code,stock_goods_type,stock_goods_code,stock_date,id_seq,inout_number,stock_company,stock_name,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_grade,stock_last_qty,stock_in_qty,stock_go_qty,stock_jj_qty,stock_id,inout_no,inout_seq) values ('"&rs("mvin_in_stock")&"','"&rs("in_goods_type")&"','"&rs("in_goods_code")&"','"&rs("mvin_in_date")&"','"&id_seq&"','"&inout_number&"','"&stock_company&"','"&stock_name&"','"&rs("in_goods_gubun")&"','"&rs("in_goods_name")&"','"&rs("in_standard")&"','"&goods_grade&"',0,'"&rs("in_qty")&"',0,0,'"&stin_id&"','"&rele_no&"','"&rs("mvin_in_seq")&"')"
	 
	     dbconn.execute(sql)
    end if
	rs.movenext()
loop
rs.close()	

'본사출고 / 고객사출고
sql = "select * from met_chulgo_goods where (cg_goods_code = '"&stock_goods_code&"') and (cg_goods_type = '"&stock_goods_type&"') and (chulgo_stock = '"&stock_code&"')"
Rs.Open Sql, Dbconn, 1
do until rs.eof

    chulgo_date = rs("chulgo_date")
	yymmdd = mid(cstr(chulgo_date),3,2) + mid(cstr(chulgo_date),6,2)  + mid(cstr(chulgo_date),9,2)
	rele_no = yymmdd + rs("chulgo_stock")
    id_seq = "3"

    jjj = jjj + 1
    inout_number = right(("00000" + cstr(jjj)),5)

    sql="insert into met_stock_inout (stock_code,stock_goods_type,stock_goods_code,stock_date,id_seq,inout_number,stock_company,stock_name,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_grade,stock_last_qty,stock_in_qty,stock_go_qty,stock_jj_qty,stock_id,inout_no,inout_seq,chulgo_return,out_service_no,trade_name,trade_dept,rele_company,rele_saupbu,rele_team,rele_stock_name) values ('"&rs("chulgo_stock")&"','"&rs("cg_goods_type")&"','"&rs("cg_goods_code")&"','"&rs("chulgo_date")&"','"&id_seq&"','"&inout_number&"','"&rs("chulgo_stock_company")&"','"&rs("chulgo_stock_name")&"','"&rs("cg_goods_gubun")&"','"&rs("cg_goods_name")&"','"&rs("cg_standard")&"','"&rs("cg_goods_grade")&"',0,0,'"&rs("cg_qty")&"',0,'"&rs("cg_type")&"','"&rele_no&"','"&rs("chulgo_seq")&"','"&rs("cg_return")&"','"&rs("rl_service_no")&"','"&rs("rl_trade_name")&"','"&rs("rl_trade_dept")&"','"&rs("rl_company")&"','"&rs("rl_saupbu")&"','"&rs("rl_team")&"','"&rs("rl_stock_name")&"')"
	
	dbconn.execute(sql)

	rs.movenext()
loop
rs.close()		

sql = "select * from met_stock_gmaster where (stock_goods_code = '"&stock_goods_code&"') and (stock_code = '"&stock_code&"') and (stock_goods_type = '"&stock_goods_type&"') ORDER BY stock_company,stock_code ASC"
Set Rs_jae = DbConn.Execute(SQL)
if not Rs_jae.eof then

   stock_level = Rs_jae("stock_level")
   goods_code = Rs_jae("stock_goods_code")
   goods_gubun = Rs_jae("stock_goods_gubun")
   goods_name = Rs_jae("stock_goods_name")
   goods_standard = Rs_jae("stock_goods_standard")
   goods_grade = Rs_jae("stock_goods_grade")
   stock_last_qty = Rs_jae("stock_last_qty")
   stock_JJ_qty = Rs_jae("stock_JJ_qty")
end if
Rs_jae.close()


sql = "select * from met_stock_inout where (stock_goods_code = '"&stock_goods_code&"') and (stock_code = '"&stock_code&"') and (stock_goods_type = '"&stock_goods_type&"') ORDER BY stock_date,id_seq,inout_no,inout_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "품목별 입.출고(현재고)현황"

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
				<form method="post" name="frm" action="met_pummok_inout_detail.asp">
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
                                <th>회사</th>
							    <td class="left"><%=stock_company%>&nbsp;</td>
							    <th>창고명</th>
							    <td class="left"><%=stock_name%>&nbsp;</td>
							    <th>창고구분</th>
							    <td class="left"><%=stock_level%>&nbsp;</td>
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
                <h3 class="stit" style="font-size:12px;">◈ 입 / 출고 내역(수량) ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
                            <col width="7%" >
                            <col width="9%" >
                            <col width="10%" >
                            <col width="9%" >
                            <col width="10%" >
                            <col width="9%" >
                            <col width="7%" >
                            <col width="5%" >
                            <col width="5%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col">일자</th>
                                <th scope="col">용도구분</th>
                                <th scope="col">구분</th>
                                <th scope="col">번호</th>
                                <th scope="col">요청사업부</th>
                                <th scope="col">입고창고</th>
                                
                                <th scope="col">고객사</th>
                                <th scope="col">지점</th>
                                <th scope="col">서비스No/<br>전표번호</th>
                                <th scope="col">전기<br>이월</th>
                                <th scope="col">입고</th>
                                <th scope="col">출고</th>
                                <th scope="col">현재고</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>     
						    <tr>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>전기이월</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td class="right"><%=formatnumber(stock_last_qty,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						
						<%
							i = 0
							h_last_qty = stock_last_qty
							h_in_qty = 0
							h_go_qty = 0
							h_jj_qty = stock_JJ_qty
							do until rs.eof or rs.bof
								h_in_qty = h_in_qty + rs("stock_in_qty")
								h_go_qty = h_go_qty + rs("stock_go_qty")
						%>
							<tr>
                                <td><%=rs("stock_date")%>&nbsp;</td>
                                <td><%=rs("stock_goods_type")%>&nbsp;</td>
                                <td><%=rs("stock_id")%>&nbsp;</td>
                                <td><%=rs("inout_no")%>&nbsp;<%=rs("inout_seq")%></td>
                                <td><%=rs("rele_saupbu")%>&nbsp;</td>
                                <td><%=rs("rele_stock_name")%>&nbsp;</td>
                                <td><%=rs("trade_name")%>&nbsp;</td>
                                <td><%=rs("trade_dept")%>&nbsp;</td>
                                <td><%=rs("out_service_no")%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
								<td><%=rs("chulgo_return")%>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
                            <tr>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>현재 재고</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=formatnumber(stock_JJ_qty,0)%>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="9" style="background:#ffe8e8;">총 계</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_last_qty,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_in_qty,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_go_qty,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_jj_qty,0)%>&nbsp;</td>
								<td style="background:#ffe8e8;">&nbsp;</td>
							</tr>
						</tbody>
					</table>
          	     <br>
     				<div class="btnleft">
                    <a href="met_pummok_inout_excel.asp?stock_company=<%=stock_company%>&stock_code=<%=stock_code%>&stock_name=<%=stock_name%>&stock_goods_code=<%=stock_goods_code%>&stock_goods_type=<%=stock_goods_type%>&goods_name=<%=goods_name%>" class="btnType04">엑셀다운로드</a>
					</div>    
                    
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

