<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

view_condi=Request("view_condi")
goods_type=Request("goods_type")
from_date=request("from_date")
to_date=request("to_date")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY stin_in_date DESC"  

if view_condi = "전체" then
   if goods_type = "전체" then
      where_sql = " WHERE (stin_id = '구매') and (stin_in_date >= '"+from_date+"' and stin_in_date <= '"+to_date+"')" 
	  else
	  where_sql = " WHERE (stin_id = '구매') and (stin_goods_type = '"&goods_type&"') and (stin_in_date >= '"+from_date+"' and stin_in_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "전체" then
      where_sql = " WHERE (stin_id = '구매') and (stin_stock_company = '"&view_condi&"') and (stin_in_date >= '"+from_date+"' and stin_in_date <= '"+to_date+"')"
	  else
	  where_sql = " WHERE (stin_id = '구매') and (stin_goods_type = '"&goods_type&"') and (stin_stock_company = '"&view_condi&"') and (stin_in_date >= '"+from_date+"' and stin_in_date <= '"+to_date+"')"
   end if
end if   

sql = "select * from met_stin " + where_sql + order_sql 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 입고 현황 - " + " (" + from_date + " ∼ " + to_date + ")"

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
                factory.printing.portrait = false; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 5; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 5; //오른쯕 여백 설정
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
				    <td height="50px" width="70%" class="style32BC"><strong><%=title_line%></strong></td>
			      </tr>
				  </table>
					<br>
                <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="6%" >
				              <col width="6%" >
                              <col width="4%" >
                              <col width="8%" >
                              <col width="8%" > 
                              <col width="*" >
                              
                              <col width="7%" >
                              <col width="8%" >
                              
                              <col width="8%" >
                              <col width="8%" >
                              <col width="8%" >
                              <col width="4%" >
                              <col width="4%" >
                              <col width="6%" >
                        </colgroup>
						<thead>
                              <tr bgcolor="#f8f8f8">
				                <th class="first" height="30" scope="col"><font style="font-size:14px">입고일자</th>
                                <th scope="col">용도구분</th>
				                <th scope="col">입고번호</th>
                                <th scope="col">입고구분</th>
                                <th scope="col">그룹사</th>
                                <th scope="col">사업부</th>
                                <th scope="col">입고창고</th>
                                
                                <th scope="col">입고금액</th>
                                <th scope="col">구매거래처</th>
                                
                                <th scope="col">품목구분</th>
                                <th scope="col">품목명</th>
                                <th scope="col">규격</th>
                                <th scope="col">수량</th>
                                <th scope="col">단가</th>
                                <th scope="col">입고액</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						do until rs.eof or rs.bof
                           stin_in_date = rs("stin_in_date")
						   
						   stin_order_no = rs("stin_order_no")
						   stin_order_seq = rs("stin_order_seq")
						   
						   k = 0
                           sql = "select * from met_stin_goods where (stin_date = '"&stin_in_date&"') and (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"')  ORDER BY stin_goods_seq,stin_goods_code ASC"
	                       Rs_buy.Open Sql, Dbconn, 1	
	                       while not Rs_buy.eof
		                     k = k + 1
							 if k = 1 then   
	           		 %>
                            <tr>
                                    <td height="30" align="left" bgcolor="#EEFFFF"><%=rs("stin_in_date")%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("stin_goods_type")%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("stin_order_no")%>-<%=rs("stin_order_seq")%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("stin_id")%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("stin_buy_company")%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("stin_buy_saupbu")%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("stin_stock_name")%>&nbsp;</td>
                                    <td bgcolor="#EEFFFF" class="right"><%=formatnumber(rs("stin_cost"),0)%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("stin_trade_name")%>&nbsp;</td>
                                    
                                    <td align="left" bgcolor="#EEFFFF"><%=Rs_buy("stin_goods_gubun")%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=Rs_buy("stin_goods_name")%>&nbsp;</td>
                                    <td align="left" bgcolor="#EEFFFF"><%=Rs_buy("stin_standard")%>&nbsp;</td>
                                    <td bgcolor="#EEFFFF" class="right"><%=formatnumber(Rs_buy("stin_qty"),0)%>&nbsp;</td>
                                    <td bgcolor="#EEFFFF" class="right"><%=formatnumber(Rs_buy("stin_unit_cost"),0)%>&nbsp;</td>
                                    <td bgcolor="#EEFFFF" class="right"><%=formatnumber(Rs_buy("stin_amt"),0)%>&nbsp;</td>
			                </tr>
            <%
			                    else
		    %>		
                                 <tr>
								    <td height="30" class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td align="right">&nbsp;</td>
                                    <td align="right">&nbsp;</td>
                                    <td align="right">&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    
								    <td align="left" ><%=Rs_buy("stin_goods_gubun")%>&nbsp;</td>
                                    <td align="left" ><%=Rs_buy("stin_goods_name")%>&nbsp;</td>
                                    <td align="left" ><%=Rs_buy("stin_standard")%>&nbsp;</td>
                                    <td class="right"><%=formatnumber(Rs_buy("stin_qty"),0)%>&nbsp;</td>
                                    <td class="right"><%=formatnumber(Rs_buy("stin_unit_cost"),0)%>&nbsp;</td>
                                    <td class="right"><%=formatnumber(Rs_buy("stin_amt"),0)%>&nbsp;</td>
						         </tr>            
            <%            							
							 end if
		                         Rs_buy.movenext()
	                       Wend
                           Rs_buy.close()
							  
						   rs.movenext()
						loop
						rs.close()
		    %>		
				</div>
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
			</form>
		</div>				
	</div>        				
	</body>
</html>

