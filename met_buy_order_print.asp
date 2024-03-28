<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

order_no = request("order_no")
order_seq = request("order_seq")
order_date = request("order_date")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_order where (order_no = '"&order_no&"') and (order_seq = '"&order_seq&"') and (order_date = '"&order_date&"')"
Set Rs_order = DbConn.Execute(SQL)
if not Rs_order.eof then
    	order_no = Rs_order("order_no")
		order_seq = Rs_order("order_seq")
		order_date = Rs_order("order_date")
		order_buy_no = Rs_order("order_buy_no")
		order_buy_seq = Rs_order("order_buy_seq")
		order_buy_date = Rs_order("order_buy_date")
		
		order_goods_type = Rs_order("order_goods_type")
		order_company = Rs_order("order_company")
	    order_bonbu = Rs_order("order_bonbu")
		order_saupbu = Rs_order("order_saupbu")
		order_team = Rs_order("order_team")
	    order_org_code = Rs_order("order_org_code")
	    order_org_name = Rs_order("order_org_name")
	    order_emp_no = Rs_order("order_emp_no")
	    order_emp_name = Rs_order("order_emp_name")
		
	    order_bill_collect = Rs_order("order_bill_collect")
        order_collect_due_date = Rs_order("order_collect_due_date")
	    order_trade_no = Rs_order("order_trade_no")
        order_trade_name = Rs_order("order_trade_name")
        order_trade_person = Rs_order("order_trade_person")
		order_trade_email = Rs_order("order_trade_email")
		
        buy_out_method = ""
        buy_out_request_date = ""
		
		order_in_date = Rs_order("order_in_date")
        order_stock_company = Rs_order("order_stock_company")
        order_stock_code = Rs_order("order_stock_code")
        order_stock_name = Rs_order("order_stock_name")
		
        order_price = Rs_order("order_price")
        order_cost = Rs_order("order_cost")
        order_cost_vat = Rs_order("order_cost_vat")
		
        order_memo = Rs_order("order_memo")
        if order_memo = "" or isnull(order_memo) then
	           order_memo = Rs_order("order_memo")
           else
	           order_memo = replace(order_memo,chr(10),"<br>")
        end if
        order_ing = Rs_order("order_ing")

	    if order_collect_due_date = "0000-00-00" then
	          order_collect_due_date = ""
	    end if
		if order_in_date = "0000-00-00" then
	      order_in_date = ""
	    end if
   else
		order_buy_no = ""
		order_buy_seq = ""
		order_buy_date = ""
		order_goods_type = ""
		order_company = ""
	    order_bonbu = ""
		order_saupbu = ""
		order_team = ""
	    order_org_code = ""
	    order_org_name = ""
	    order_emp_no = ""
	    order_emp_name = ""
	    order_bill_collect = ""
        order_collect_due_date = ""
	    order_trade_no = ""
        order_trade_name = ""
        order_trade_person = ""
		order_trade_email = ""
        buy_out_method = ""
        buy_out_request_date = ""
		order_in_date = ""
        order_stock_company = ""
        order_stock_code = ""
        order_stock_name = ""
        order_price = 0
        order_cost = 0
        order_cost_vat = 0
        order_memo = ""
        order_ing = ""
end if
Rs_order.close()


if order_company = "케이원정보통신" then
      company_name = "(주)" + "케이원정보통신"
	  owner_name = "김승일"
	  addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	  trade_no = "107-81-54150"
	  tel_no = "02) 853-5250"
	  e_mail = "js10547@k-won.co.kr"
   elseif order_company = "휴디스" then
              company_name = "(주)" + "휴디스"
			  owner_name = "김한종"
	          addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	          trade_no = "107-81-54150"
	          tel_no = "02) 853-5250"
	          e_mail = "js10547@k-won.co.kr"
		  elseif order_company = "케이네트웍스" then
                     company_name = "케이네트웍스" + "(주)"
					 owner_name = "이중원"
	                 addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                 trade_no = "107-81-54150"
	                 tel_no = "02) 853-5250"
	                 e_mail = "js10547@k-won.co.kr"
				 elseif order_company = "에스유에이치" then
                        company_name = "(주)" + "에스유에이치"	
						owner_name = "박미애"
	                    addr_name = "서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)"
	                    trade_no = "119-86-78709"
	                    tel_no = "02) 6116-8248"
	                    e_mail = "pshwork27@k-won.co.kr"
end if 

sql = "select * from met_order_goods where (og_order_no = '"&order_no&"') and (og_order_seq = '"&order_seq&"') and (og_order_date = '"&order_date&"') ORDER BY og_seq,og_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "발 주 서"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>자재관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
		  		 window.close () ;
			}
			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = true; //출력방향 설정: true - 가로, false - 세로
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
				
					document.frm.method = "post";
//					document.frm.enctype = "multipart/form-data";
					document.frm.action = "met_buy_order_prt_ok.asp";
					document.frm.submit();
            }
        </script>
        <style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 14px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style14C {font-size: 14px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
		.style14R {font-size: 14px; font-family: "굴림체", "굴림체", Seoul; text-align: right; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
-->
    </style>
	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="wrap">			
			<div id="container">
				<form action="met_buy_order_print.asp" method="post" name="frm">
				<div class="gView">
				<table width="1150" cellpadding="0" cellspacing="0">
				  <tr>
				    <td height="50px" class="style32BC"><strong><%=title_line%></strong></td>
			      </tr>
				  </table>
					<br>
				<table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
							<col width="20%" >
							<col width="30%" >
							<col width="4%" >
							<col width="16%" >
							<col width="30%" >
						</colgroup>
						<thead>
							<tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">발주일자(번호)</td>
                              <td align="center" class="style14C"><%=order_date%>&nbsp;(<%=order_no%>&nbsp;<%=order_seq%>)</td>
                              <th rowspan="6" align="center" class="style14C" style="background:#f8f8f8;">발<br>주<br>처</th>
                              <th align="center" class="style14C" style="background:#f8f8f8;">사업자등록번호</th>
                              <td align="center" class="style14C"><%=trade_no%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">거래처명</td>
                              <td align="center" class="style14C"><%=order_trade_name%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">상호</th>
                              <td align="center" class="style14C"><%=company_name%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">TEL No.</td>
                              <td align="center" class="style14C"><%=tel_no%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">주소</th>
                              <td class="left"><font style="font-size:14px"><%=addr_name%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">FAX No.</td>
                              <td align="center" class="style14C"><%=tel_no%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">발주담당자</th>
                              <td align="center" class="style14C"><%=order_emp_name%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">담당자</td>
                              <td align="center" class="style14C"><%=order_trade_person%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">TEL No.</th>
                              <td align="center" class="style14C"><%=tel_no%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">납기일</td>
                              <td align="center" class="style14C"><%=order_in_date%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">FAX No.</th>
                              <td align="center" class="style14C"><%=tel_no%></td>
						    </tr>
						</thead>
					</table>
                     <br>
                <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="*" >
                              <col width="20%" >
                              <col width="10%" >
                              <col width="10%" >
							  <col width="12%" >
							  <col width="14%" >
							  <col width="10%" >
                        </colgroup>
						 <thead>
                              <tr bgcolor="#f8f8f8">
                                <th class="first" height="30" align="center" scope="col" class="style14C">품 명</th>
                                <th scope="col" align="center" class="style14C">규 격</th>
                                <th scope="col" align="center" class="style14C">단위</th>
                                <th scope="col" align="center" class="style14C">수량</th>
                                <th scope="col" align="center" class="style14C">단 가</th>
                                <th scope="col" align="center" class="style14C">금 액</th>
                                <th scope="col" align="center" class="style14C">비고</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						do until rs.eof or rs.bof
                             
	           		 %>
							<tr>
                                <td height="30" align="center" class="style14C"><%=rs("og_goods_name")%>&nbsp;</td>
                                <td align="center" class="style14C"><%=rs("og_standard")%>&nbsp;</td>
                                <td align="center" class="style14C">&nbsp;</td>
                                <td align="right" class="style14R"><%=formatnumber(rs("og_qty"),0)%>&nbsp;</td>
                                <td align="right" class="style14R"><%=formatnumber(rs("og_unit_cost"),0)%>&nbsp;</td>
                                <td align="right" class="style14R" ><%=formatnumber(rs("og_amt"),0)%>&nbsp;</td>
                                <td align="center" class="style14C">&nbsp;</td>
							</tr>
					<%
							rs.movenext()
						loop
						rs.close()
					%>
                            <tr>
                                <td height="30" align="center" class="style14C" style="background:#f8f8f8;">공급가액</td>
                                <td align="right" class="style14R"><%=formatnumber(order_cost,0)%>&nbsp;</td>
                                <td colspan="2" align="center" class="style14C" style="background:#f8f8f8;">부가세액</td>
                                <td align="right" class="style14R"><%=formatnumber(order_cost_vat,0)%>&nbsp;</td>
                                <td align="center" class="style14C" style="background:#f8f8f8;">합계금액</td>
                                <td align="right" class="style14R"><%=formatnumber(order_price,0)%>&nbsp;</td>
							</tr>
						</tbody>
					</table> 
                    <br>
                    <h3 class="stit">1. 귀사의 일익 번창하심을 기원합니다.</h3>
                    <h3 class="stit">&nbsp;&nbsp;&nbsp;상기와 같이 발주하오니 납기일을 준수하여 입고 바랍니다.</h3> 
                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
							<col width="20%" >
							<col width="80%" >
						</colgroup>
						<thead>
							<tr>
                              <td height="30" align="center" class="style14C" style="background:#f8f8f8;">대금결재 조건</td>
                              <td class="left" ><font style="font-size:14px">&nbsp;<%=order_collect_due_date%>&nbsp;-&nbsp;<%=order_bill_collect%></td>
						    </tr>
                            <tr>
                              <td height="30" align="center" class="style14C" style="background:#f8f8f8;">납품 장소</td>
                              <td class="left" ><font style="font-size:14px">&nbsp;<%=order_stock_name%>&nbsp;-&nbsp;<%=addr_name%></td>
						    </tr>
                            <tr>
                              <td height="30" align="center" class="style14C" style="background:#f8f8f8;">특기 사항</td>
                              <td class="left"><font style="font-size:14px">&nbsp;<%=order_memo%></td>
						    </tr>
						</thead>
					</table>  
				<table width="1150" border="0" cellpadding="0" cellspacing="0" align="center" class="onlyprint">    
				  <tr>
				     <td colspan="2" height="100" align="center"><font style="font-size:16px"><strong>위 내용과 같이 발주 합니다.</td>
	              </tr>
	              <tr>
		             <td colspan="2" height="60" align="right" width="100%"><font style="font-size:14px"><%=mid(cstr(now()),1,4)%>년&nbsp;<%=mid(cstr(now()),6,2)%>월&nbsp;<%=mid(cstr(now()),9,2)%>일<br/><br/>
		서울시 금천구 가산디지털2로 18(대륭테크노타운 1차 6층)</td>
	             </tr>
	             <tr>  
	                <td height="60" align="right" width="95%"><font style="font-size:14px"><br><br>주식회사 케이원정보통신<br/>
		<font style="font-size:14px">대표이사 </font><font style="font-size:16px"><b>김승일</b></font></td>
                    <td height="60" align="right" valign="middle" width="5%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right></td>
	             </tr>                    
				</table>
                <br><br><br>
				<table width="1150" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<br>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				    <br>                 
                    </td>
			      </tr>
				</table>
                <input type="hidden" name="old_order_no" value="<%=order_no%>">
				<input type="hidden" name="old_order_seq" value="<%=order_seq%>">
                <input type="hidden" name="old_order_date" value="<%=order_date%>">
                
                <input type="hidden" name="order_buy_no" value="<%=order_buy_no%>">
				<input type="hidden" name="order_buy_seq" value="<%=order_buy_seq%>">
                <input type="hidden" name="order_buy_date" value="<%=order_buy_date%>">
			</form>
		</div>				
	</div>        				
	</body>
</html>

