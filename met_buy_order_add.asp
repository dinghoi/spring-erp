<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim code_tab(20)
dim goods_name(20)
dim goods_ttype(20)
dim goods_gubun(20)
dim goods_standard(20)
dim qty_tab(20)
dim unit_cost(20)
dim buy_amt(20)
dim bg_seq_tab(20)
dim oqty_tab(20)
dim b_qty(20)

dim pummok_tab(6,20)
dim amount_tab(4,20)

for i = 1 to 20
    code_tab(i) = ""
	goods_name(i) = ""
	goods_ttype(i) = ""
	goods_gubun(i) = ""
	goods_standard(i) = ""
	qty_tab(i) = 0
	unit_cost(i) = 0
	buy_amt(i) = 0
	bg_seq_tab(i) = ""
	oqty_tab(i) = 0
	b_qty(i) = 0
next

for i = 1 to 6
	for j = 1 to 20
		pummok_tab(i,j) = ""
	next
next
for i = 1 to 4
	for j = 1 to 20
		amount_tab(i,j) = 0
	next
next

u_type = request("u_type")

view_condi=Request("view_condi")
buy_no=Request("buy_no")
buy_date=Request("buy_date")
buy_seq = request("buy_seq")

curr_date = mid(cstr(now()),1,10)
order_date = curr_date
order_in_date = curr_date

order_stock_company = ""
order_stock_code = ""
order_stock_name = ""
order_memo = ""

mok_cnt = 0
pummok_cnt = 0

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_bg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

'구매품의 조회
   sql = "select * from met_buy where (buy_no = '"&buy_no&"') and (buy_date = '"&buy_date&"') and (buy_seq = '"&buy_seq&"')"
   Set Rs_buy = DbConn.Execute(SQL)
   if not Rs_buy.eof then
    	buy_no = Rs_buy("buy_no")
		buy_seq = Rs_buy("buy_seq")
		buy_date = Rs_buy("buy_date")
		buy_goods_type = Rs_buy("buy_goods_type")
		buy_company = Rs_buy("buy_company")
	    buy_bonbu = Rs_buy("buy_bonbu")
	    buy_saupbu = Rs_buy("buy_saupbu")
	    buy_team = Rs_buy("buy_team")
	    buy_org_code = Rs_buy("buy_org_code")
	    buy_org_name = Rs_buy("buy_org_name")
	    buy_emp_no = Rs_buy("buy_emp_no")
	    buy_emp_name = Rs_buy("buy_emp_name")
	    buy_bill_collect = Rs_buy("buy_bill_collect")
        buy_collect_due_date = Rs_buy("buy_collect_due_date")
'	    buy_trade_no = Rs_buy("buy_trade_no")
		buy_trade_no = mid(Rs_buy("buy_trade_no"),1,3) + "-" + mid(Rs_buy("buy_trade_no"),4,2) + "-" + right(Rs_buy("buy_trade_no"),5)		
        buy_trade_name = Rs_buy("buy_trade_name")
        buy_trade_person = Rs_buy("buy_trade_person")
		buy_trade_email = Rs_buy("buy_trade_email")
        buy_out_method = Rs_buy("buy_out_method")
        buy_out_request_date = Rs_buy("buy_out_request_date")
        buy_price = Rs_buy("buy_price")
        buy_cost = Rs_buy("buy_cost")
        buy_cost_vat = Rs_buy("buy_cost_vat")
        buy_memo = Rs_buy("buy_memo")
        buy_ing = Rs_buy("buy_ing")
		buy_sign_yn = Rs_buy("buy_sign_yn")
	    buy_sign_no = Rs_buy("buy_sign_no")
	    buy_sign_date = Rs_buy("buy_sign_date")
		buy_att_file = Rs_buy("buy_att_file")

	    if buy_out_request_date = "0000-00-00" then
	          buy_out_request_date = ""
	    end if
		if buy_collect_due_date = "0000-00-00" then
	          buy_collect_due_date = ""
	    end if
     else
		buy_goods_type = ""
		buy_company = ""
	    buy_bonbu = ""
	    buy_saupbu = ""
	    buy_team = ""
	    buy_org_code = ""
	    buy_org_name = ""
	    buy_emp_no = ""
	    buy_emp_name = ""
	    buy_bill_collect = ""
        buy_collect_due_date = ""
	    buy_trade_no = ""
        buy_trade_name = ""
        buy_trade_person = ""
		buy_trade_email = ""
        buy_out_method = ""
        buy_out_request_date = ""
        buy_price = 0
        buy_cost = 0
        buy_cost_vat = 0
        buy_memo = ""
        buy_ing = ""
		buy_sign_yn = ""
	    buy_sign_no = ""
	    buy_sign_date = ""
		buy_att_file = ""
   end if
   Rs_buy.close()
   
   i = 0
   buy_cost = 0
   
   sql = "select * from met_buy_goods where (bg_no = '"&buy_no&"') and (bg_date = '"&buy_date&"') and (buy_seq = '"&buy_seq&"') ORDER BY bg_seq,bg_goods_code ASC"
   Set Rs_good = DbConn.Execute(SQL)
   do until Rs_good.eof or Rs_good.bof
        i = i +1
	    bg_seq_tab(i) = Rs_good("bg_seq")
		goods_ttype(i) = Rs_good("bg_goods_type")
	    goods_gubun(i) = Rs_good("bg_goods_gubun")
		code_tab(i) = Rs_good("bg_goods_code")
		goods_name(i) = Rs_good("bg_goods_name")
		goods_standard(i) = Rs_good("bg_standard")
		qty_tab(i) = Rs_good("bg_qty")
		unit_cost(i) = Rs_good("bg_unit_cost")
		'buy_amt(i) = Rs_good("bg_buy_amt")
		oqty_tab(i) = Rs_good("bg_order_qty")
		
		b_qty(i) = qty_tab(i) - oqty_tab(i)
		buy_amt(i) = b_qty(i) * unit_cost(i)
		buy_cost = buy_cost + buy_amt(i)

        Rs_good.movenext()
   loop
   mok_cnt = i
   Rs_good.close()
   
   buy_cost_vat = int(buy_cost * ( 10 / 100 ))
   buy_price = buy_cost + buy_cost_vat
   
title_line = buy_goods_type + " 발주 등록 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>상품자재관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=buy_out_request_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=order_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=bill_due_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=bill_issue_date%>" );
			});	  
			$(function() {    $( "#datepicker4" ).datepicker();
												$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker4" ).datepicker("setDate", "<%=buy_collect_due_date%>" );
			});	  
			$(function() {    $( "#datepicker5" ).datepicker();
												$( "#datepicker5" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker5" ).datepicker("setDate", "<%=order_in_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm1()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
						
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function chkfrm1() {
//				if(document.frm.sales_company.value == "") {
//					alert('매출회사를 선택하세요');
//					frm.sales_company.focus();
//					return false;}
				if(document.frm.order_stock_name.value == "") {
					alert('입고창고를 선택하세요');
					frm.order_stock_name.focus();
					return false;}
				if(document.frm.trade_name.value == "") {
					alert('거래처를 선택하세요');
					frm.trade_name.focus();
					return false;}
				if(document.frm.trade_no.value == "") {
					alert('거래처를 선택하세요');
					frm.trade_no.focus();
					return false;}
				if(document.frm.trade_person.value == "") {
					alert('거래처 담당자를 입력하세요');
					frm.trade_person.focus();
					return false;}

				k = 0;
				for (j=0;j<4;j++) {
					if (eval("document.frm.bill_collect[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("대금지급방법을 선택하세요");
					return false;
				}	
						
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

		function NumCal(txtObj){
			var bqty_ary = new Array();
			var b_qty_ary = new Array();
			var qty_ary = new Array();
			var buy_ary = new Array();
			var buy_tot = new Array();

			for (j=1;j<21;j++) {
				bqty_ary[j] = eval("document.frm.bqty" + j + ".value").replace(/,/g,"");
				b_qty_ary[j] = eval("document.frm.b_qty" + j + ".value").replace(/,/g,"");
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
				buy_ary[j] = eval("document.frm.buy_cost" + j + ".value").replace(/,/g,"");
				
				acpt_qty = parseInt(qty_ary[j]);
				sign_qty = parseInt(b_qty_ary[j]);

	            if (acpt_qty > sign_qty) {
					alert ("품의수량보다 발주수량이 많습니다!!");
					return false;
				}
				
				buy_tot[j] = qty_ary[j] * buy_ary[j];
				
				buy_cal = qty_ary[j] * buy_ary[j];
				buy_cal = String(buy_cal);
				num_len = buy_cal.length;
				sil_len = num_len;
				if (buy_cal.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) buy_cal = buy_cal.substr(0,num_len -3) + "," + buy_cal.substr(num_len -3,3);
				if (sil_len > 6) buy_cal = buy_cal.substr(0,num_len -6) + "," + buy_cal.substr(num_len -6,3) + "," + buy_cal.substr(num_len -2,3);
				if (sil_len > 9) buy_cal = buy_cal.substr(0,num_len -9) + "," + buy_cal.substr(num_len -9,3) + "," + buy_cal.substr(num_len -5,3) + "," + buy_cal.substr(num_len -1,3);
				eval("document.frm.buy_tot" + j + ".value = buy_cal");
				
			}
			
			buy_tot_cost = 0;
			for (j=1;j<21;j++) {
				buy_tot_cost = buy_tot_cost + buy_tot[j];
			}
			
			tot_cal = buy_tot_cost;
			tot_cal = String(tot_cal);
			num_len = tot_cal.length;
			sil_len = num_len;
			if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
			if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
			if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
			eval("document.frm.buy_tot_cost.value = tot_cal");
			
			buy_vat = parseInt(buy_tot_cost * (10 / 100));
			buy_hap = buy_tot_cost + buy_vat;
			
			buy_vat = String(buy_vat);
			num_len = buy_vat.length;
			sil_len = num_len;
			if (buy_vat.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) buy_vat = buy_vat.substr(0,num_len -3) + "," + buy_vat.substr(num_len -3,3);
			if (sil_len > 6) buy_vat = buy_vat.substr(0,num_len -6) + "," + buy_vat.substr(num_len -6,3) + "," + buy_vat.substr(num_len -2,3);
			if (sil_len > 9) buy_vat = buy_vat.substr(0,num_len -9) + "," + buy_vat.substr(num_len -9,3) + "," + buy_vat.substr(num_len -5,3) + "," + buy_vat.substr(num_len -1,3);
			eval("document.frm.buy_tot_cost_vat.value = buy_vat");
			
			buy_hap = String(buy_hap);
			num_len = buy_hap.length;
			sil_len = num_len;
			if (buy_hap.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) buy_hap = buy_hap.substr(0,num_len -3) + "," + buy_hap.substr(num_len -3,3);
			if (sil_len > 6) buy_hap = buy_hap.substr(0,num_len -6) + "," + buy_hap.substr(num_len -6,3) + "," + buy_hap.substr(num_len -2,3);
			if (sil_len > 9) buy_hap = buy_hap.substr(0,num_len -9) + "," + buy_hap.substr(num_len -9,3) + "," + buy_hap.substr(num_len -5,3) + "," + buy_hap.substr(num_len -1,3);
			eval("document.frm.buy_tot_price.value = buy_hap");

			if (txtObj.value.length<1) {
				txtObj.value=txtObj.value.replace(/,/g,"");
				txtObj.value=txtObj.value.replace(/\D/g,"");
			}
			var num = txtObj.value;
			if (num == "--" ||  num == "." ) num = "";
			if (num != "" ) {
				temp=new String(num);
				if(temp.length<1) return "";
							
				// 음수처리
				if(temp.substr(0,1)=="-") minus="-";
					else minus="";
							
				// 소수점이하처리
				dpoint=temp.search(/\./);
						
				if(dpoint>0)
				{
				// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
				dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
				temp=temp.substr(0,dpoint);
				}else dpointVa="";
							
				// 숫자이외문자 삭제
				temp=temp.replace(/\D/g,"");
				zero=temp.search(/[1-9]/);
						
				if(zero==-1) return "";
				else if(zero!=0) temp=temp.substr(zero);
							
				if(temp.length<4) return minus+temp+dpointVa;
				buf="";
				while (true)
				{
				if(temp.length<3) { buf=temp+buf; break; }
					
				buf=","+temp.substr(temp.length-3)+buf;
				temp=temp.substr(0, temp.length-3);
				}
				if(buf.substr(0,1)==",") buf=buf.substr(1);
						
				//return minus+buf+dpointVa;
				txtObj.value = minus+buf+dpointVa;
			}else txtObj.value = "0";					
		}
		function pummok_list_view() {
				mok_cnt = parseInt(document.frm.mok_cnt.value);
				for (j=1;j<mok_cnt+1;j++) {
					eval("document.getElementById('pummok_list" + j + "')").style.display = '';
				}
				NumCal();
			}
		</script>

	</head>
	<body onload="pummok_list_view();">
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_buy_order_add_save.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="18%" >
						</colgroup>
						<tbody>
							<tr>
                                <th>구매번호</th>
                                <td class="left"><%=buy_no%>&nbsp;<%=buy_seq%></td>
                                <th>구매용도</th>
                                <td class="left"><%=buy_goods_type%></td>
                                <th>구매<br>품의일자</th>
                                <td class="left"><%=buy_date%></td>
                                <th>구매회사</th>
                                <td class="left"><%=buy_company%>
                                <input type="hidden" name="buy_no" value="<%=buy_no%>" ID="buy_no">
                                <input type="hidden" name="buy_seq" value="<%=buy_seq%>" ID="buy_seq">
                                <input type="hidden" name="buy_date" value="<%=buy_date%>" ID="buy_date">
                                <input type="hidden" name="buy_goods_type" value="<%=buy_goods_type%>" ID="buy_goods_type">
                                </td>
                            </tr>
                            <tr>
                                <th>사업부</th>
                                <td class="left"><%=buy_saupbu%></td>
                                <th>소속</th>
                                <td class="left"><%=buy_org_name%></td>
                                <th>구매담당</th>
                                <td colspan="3" class="left"><%=buy_emp_name%>&nbsp;(<%=buy_emp_no%>)
                                <input type="hidden" name="order_company" value="<%=buy_company%>" ID="buy_company">
                                <input type="hidden" name="order_bonbu" value="<%=buy_bonbu%>" ID="buy_bonbu">
                                <input type="hidden" name="order_saupbu" value="<%=buy_saupbu%>" ID="buy_saupbu">
                                <input type="hidden" name="order_team" value="<%=buy_team%>" ID="buy_team">
                                <input type="hidden" name="order_org_code" value="<%=buy_org_code%>" ID="buy_org_code">
                                <input type="hidden" name="order_org_name" value="<%=buy_org_name%>" ID="buy_org_name">
                                <input type="hidden" name="order_emp_no" value="<%=buy_emp_no%>" ID="buy_emp_no">
                                <input type="hidden" name="order_emp_name" value="<%=buy_emp_name%>" ID="buy_emp_name">
                                </td>
 							</tr>
						    <tr>
							  <th>발주일자</th>
							  <td class="left"><input name="order_date" type="text" value="<%=order_date%>" style="width:80px;text-align:center" id="datepicker1"></td>
							  <th>발주거래처</th>
							  <td class="left"><input name="trade_name" type="text" value="<%=buy_trade_name%>" readonly="true" style="width:120px">
						      <a href="#" class="btnType03" onClick="pop_Window('insa_trade_select.asp?gubun=<%="buy"%>','trade_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              </td>
							  <th>사업자번호</th>
							  <td class="left"><input name="trade_no" type="text" value="<%=buy_trade_no%>" readonly="true" style="width:150px"></td>
							  <th>거래처<br>담당자</th>
							  <td class="left"><input name="trade_person" type="text" value="<%=buy_trade_person%>" id="trade_person" style="width:120px; ime-mode:active" onKeyUp="checklength(this,20);"></td>
						    </tr>
                            <tr>
							  <th>이메일</th>
							  <td class="left"><input name="trade_email" type="text" value="<%=buy_trade_email%>" id="trade_email" style="width:180px;" onKeyUp="checklength(this,30);"></td>
                              <th>대금<br>지급방법</th>
							  <td colspan="3" class="left">
                              <input type="radio" name="bill_collect" value="현금" <% if buy_bill_collect = "현금" then %>checked<% end if %> style="width:40px" id="Radio3">현금
  							  <input type="radio" name="bill_collect" value="어음" <% if buy_bill_collect = "어음" then %>checked<% end if %> style="width:40px" id="Radio4">어음
                              <input type="radio" name="bill_collect" value="카드" <% if buy_bill_collect = "카드" then %>checked<% end if %> style="width:40px" id="Radio3">카드
  							  <input type="radio" name="bill_collect" value="외환" <% if buy_bill_collect = "외환" then %>checked<% end if %> style="width:40px" id="Radio4">외환
                              </td>
							  <th>지급예정일</th>
							  <td class="left"><input name="collect_due_date" type="text" value="<%=buy_collect_due_date%>" style="width:80px;text-align:center" id="datepicker4"></td>
						    </tr>
                            <tr>
							  <th>입고예정<br>창고</th>
							  <td colspan="5" class="left">
                              <input name="order_stock_company" type="text" value="<%=order_stock_company%>" readonly="true" style="width:120px">
                              -
                              <input name="order_stock_name" type="text" value="<%=order_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="order"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="order_stock_code" value="<%=order_stock_code%>" ID="order_stock_code">
                              <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="stock_bonbu">
                              <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="stock_bonbu">
                              <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="stock_team">
                              <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="stock_manager_code">
                              <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="stock_manager_name">
                              </td>
                              <th>입고예정일</th>
							  <td class="left"><input name="order_in_date" type="text" value="<%=order_in_date%>" style="width:80px;text-align:center" id="datepicker5"></td>
						    </tr>
                            <tr>
                                <th>구매의견</th>
                                <td colspan="7" class="left"><%=buy_memo%>&nbsp;
                                <input type="hidden" name="buy_memo" value="<%=buy_memo%>" ID="buy_memo">
                                </td>
                            </tr>
                            <tr>
							  <th>비고</th>
							  <td class="left" colspan="7" ><textarea name="order_memo" rows="3" id="textarea"><%=order_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
				</div>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 구매품의 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="14%" >
							<col width="12%" >
                            <col width="6%" >
                            <col width="6%" >
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
                                <th scope="col">구매수량</th>
                                <th scope="col">기발주량</th>
                                <th scope="col">구매단가</th>
								<th scope="col">발주수량</th>
								<th scope="col">발주금액</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
'							    if code_tab(i) = "" or isnull(code_tab(i)) then 
'			                           exit for
'		                           else
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><%=i%></td>
								<td><%=goods_ttype(i)%>
                                <input type="hidden" name="srv_type<%=i%>" value="<%=goods_ttype(i)%>" ID="Hidden1">
                                <input type="hidden" name="bg_seq<%=i%>" value="<%=bg_seq_tab(i)%>" ID="Hidden1">
                                </td>
                                <td><%=goods_gubun(i)%>
                                <input type="hidden" name="goods_gubun<%=i%>" value="<%=goods_gubun(i)%>" ID="Hidden1">
                                </td>
                                <td><%=code_tab(i)%>
                                <input type="hidden" name="goods_code<%=i%>" value="<%=code_tab(i)%>" ID="Hidden1">
                                </td>
								<td><%=goods_name(i)%>
                                <input type="hidden" name="goods_name<%=i%>" value="<%=goods_name(i)%>" ID="Hidden1">
								</td>
                                <td><%=goods_standard(i)%>
                                <input type="hidden" name="goods_standard<%=i%>" value="<%=goods_standard(i)%>" ID="Hidden1">
                                </td>
								<td align="right"><%=formatnumber(qty_tab(i),0)%>
                                <input type="hidden" name="bqty<%=i%>" value="<%=formatnumber(qty_tab(i),0)%>" ID="Hidden1">
                                </td>
                                <td align="right"><%=formatnumber(oqty_tab(i),0)%>
                                <input type="hidden" name="oqty<%=i%>" value="<%=formatnumber(oqty_tab(i),0)%>" ID="Hidden1">
                                </td>
                                <td align="right"><%=formatnumber(unit_cost(i),0)%>
                                <input type="hidden" name="buy_cost<%=i%>" value="<%=formatnumber(unit_cost(i),0)%>" ID="Hidden1">
                                <input type="hidden" name="b_qty<%=i%>" value="<%=formatnumber(b_qty(i),0)%>" ID="Hidden1">
                                </td>
              <% if  b_qty(i) = 0 then  %>
                                <td align="right"><%=formatnumber(b_qty(i),0)%>
                                </td>
              <%     else               %>               
                                <td><input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(b_qty(i),0)%>" onChange="NumCal(this);"></td>
              <% end if                 %>                     
								<td><input name="buy_tot<%=i%>" type="text" disabled id="buy_tot<%=i%>" style="width:80px;text-align:right" readonly="true" value="<%=formatnumber(buy_amt(i),0)%>"></td>
							</tr>
  			  <%
'						        end if
							next
			  %>
						</tbody>
					</table>                    
					<br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="21%" >
							<col width="13%" >
							<col width="21%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>발주총액</th>
							  <td class="left"><input name="buy_tot_price" type="text" id="buy_tot_price" style="width:150px;text-align:right" value="<%=formatnumber(buy_price,0)%>" readonly="true"></td>
							  <th>발주금액</th>
							  <td class="left"><input name="buy_tot_cost" type="text" id="buy_tot_cost" style="width:150px;text-align:right" value="<%=formatnumber(buy_cost,0)%>" readonly="true"></td>
							  <th>부가세</th>
							  <td class="left"><input name="buy_tot_cost_vat" type="text" id="buy_tot_cost_vat" style="width:150px;text-align:right" value="<%=formatnumber(buy_cost_vat,0)%>" readonly="true"></td>
						    </tr>
							<tr>
							  <th>구매품의 첨부</th>
							  <td colspan="5" class="left">
                        <% 
                           If buy_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=buy_att_file%>"><%=buy_att_file%></a>
                        <%    Else %>
				                    &nbsp;
                        <% 
						   End If %>
                              </td>
						    </tr>
						</tbody>
					</table>
				</div>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
            <% if u_type = "U" then	%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();"></span>
			<% end if	%>                          
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                
                <input type="hidden" name="old_order_no" value="<%=order_no%>">
                <input type="hidden" name="old_order_seq" value="<%=order_seq%>">
				<input type="hidden" name="old_order_date" value="<%=order_date%>">
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
				</form>
                </div>
			</div>
		</div>		
	</body>  
</html>
