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

slip_id=Request("slip_id")
slip_no=Request("slip_no")
slip_seq=Request("slip_seq")
sales_date = request("sales_date")

curr_date = mid(cstr(now()),1,10)
order_date = curr_date
order_in_date = curr_date

order_stock_company = ""
order_stock_code = ""
order_stock_name = ""
order_memo = ""
order_bill_collect = "현금"
order_collect_due_date = curr_date

if slip_id = "2" then
		slip_id_view = "수주전표"
end if
if slip_id = "1" then
		slip_id_view = "대기전표"
end if

mok_cnt = 0
pummok_cnt = 0

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_sale = Server.CreateObject("ADODB.Recordset")
Set Rs_bg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

'매출 조회
   sql = "select * from sales_slip where slip_id = '"&slip_id&"' and slip_no = '"&slip_no&"' and slip_seq = '"&slip_seq&"'"
   Set Rs_buy = DbConn.Execute(SQL)
   if not Rs_buy.eof then
    	sales_company = Rs_buy("sales_company")
		sales_bonbu = Rs_buy("sales_bonbu")
		sales_saupbu = Rs_buy("sales_saupbu")
		sales_team = Rs_buy("sales_team")
		sales_org_name = Rs_buy("sales_org_name")
	    emp_no = Rs_buy("emp_no")
	    emp_name = Rs_buy("emp_name")
	    emp_company = Rs_buy("emp_company")
	    bonbu = Rs_buy("bonbu")
	    saupbu = Rs_buy("saupbu")
	    team = Rs_buy("team")
	    org_name = Rs_buy("org_name")
'	    trade_no = Rs_buy("trade_no")
		trade_no = mid(Rs_buy("trade_no"),1,3) + "-" + mid(Rs_buy("trade_no"),4,2) + "-" + right(Rs_buy("trade_no"),5)
        trade_name = Rs_buy("trade_name")
        trade_person = Rs_buy("trade_person")
		trade_email = Rs_buy("trade_email")
        out_method = Rs_buy("out_method")
        out_request_date = Rs_buy("out_request_date")
		sales_date = Rs_buy("sales_date")
        bill_collect = Rs_buy("bill_collect")
		collect_due_date = Rs_buy("collect_due_date")
        slip_memo = Rs_buy("slip_memo")
        buy_price = Rs_buy("buy_price")
        buy_cost = Rs_buy("buy_cost")
        buy_cost_vat = Rs_buy("buy_cost_vat")
		sign_yn = Rs_buy("sign_yn")
	    sign_no = Rs_buy("sign_no")
	    sign_date = Rs_buy("sign_date")
		att_file = Rs_buy("att_file")

	    if out_request_date = "0000-00-00" then
	          out_request_date = ""
	    end if
		if collect_due_date = "0000-00-00" then
	          collect_due_date = ""
	    end if
   end if
   Rs_buy.close()
   
   i = 0
   buy_cost = 0
   
   sql = "select * from sales_slip_detail where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"' ORDER BY goods_seq ASC"
   Set Rs_good = DbConn.Execute(SQL)
   do until Rs_good.eof or Rs_good.bof
        i = i +1
	    bg_seq_tab(i) = Rs_good("goods_seq")
		goods_ttype(i) = Rs_good("srv_type")
'	    goods_gubun(i) = Rs_good("bg_goods_gubun")
		code_tab(i) = Rs_good("goods_code")
		goods_name(i) = Rs_good("pummok")
		goods_standard(i) = Rs_good("standard")
		qty_tab(i) = Rs_good("qty")
		unit_cost(i) = Rs_good("buy_cost")

		b_qty(i) = qty_tab(i) - oqty_tab(i)
		buy_amt(i) = b_qty(i) * unit_cost(i)
		buy_cost = buy_cost + buy_amt(i)

        Rs_good.movenext()
   loop
   mok_cnt = i
   Rs_good.close()
   
   buy_cost_vat = int(buy_cost * ( 10 / 100 ))
   buy_price = buy_cost + buy_cost_vat
   
title_line = slip_id_view + " 발주 등록 "

Sql = "SELECT * FROM emp_master where emp_no = '"&user_id&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	order_emp_no = rs_emp("emp_no")
		order_emp_name = rs_emp("emp_name")
		order_company = rs_emp("emp_company")
		order_bonbu = rs_emp("emp_bonbu")
		order_saupbu = rs_emp("emp_saupbu")
		order_team = rs_emp("emp_team")
		order_org_code = rs_emp("emp_org_code")
		order_org_name = rs_emp("emp_org_name")
   else
		order_emp_no = ""
		order_emp_name = ""
		order_company = ""
		order_bonbu = ""
		order_saupbu = ""
		order_team = ""
		order_org_code = ""
		order_org_name = ""
end if
rs_emp.close()

path_name = "/sales_file"

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
												$( "#datepicker" ).datepicker("setDate", "<%=order_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=order_collect_due_date%>" );
			});	 
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=order_in_date%>" );
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
                buy_tot[j] = qty_ary[j] * buy_ary[j];
				
				cost_u = parseInt(buy_ary[j]);
				cost_u = String(cost_u);
				num_len = cost_u.length;
				sil_len = num_len;
				if (cost_u.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) cost_u = cost_u.substr(0,num_len -3) + "," + cost_u.substr(num_len -3,3);
				if (sil_len > 6) cost_u = cost_u.substr(0,num_len -6) + "," + cost_u.substr(num_len -6,3) + "," + cost_u.substr(num_len -2,3);
				if (sil_len > 9) cost_u = cost_u.substr(0,num_len -9) + "," + cost_u.substr(num_len -9,3) + "," + cost_u.substr(num_len -5,3) + "," + cost_u.substr(num_len -1,3);
				eval("document.frm.buy_cost" + j + ".value = cost_u");
				
				
				acpt_qty = parseInt(qty_ary[j]);
				sign_qty = parseInt(b_qty_ary[j]);

				if (acpt_qty > sign_qty) {
					alert ("품의수량보다 발주수량이 많습니다!!");
					return false;
				}

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
				alert (buy_tot[j]);
				
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
                <form method="post" name="frm" action="met_sales_order_add_save.asp">
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
                                <th>전표구분</th>
                                <td class="left"><%=slip_id_view%></td>
                                <th>전표번호</th>
                                <td class="left"><%=slip_no%>&nbsp;<%=slip_seq%></td>
                                <th>매출일자</th>
                                <td class="left"><%=sales_date%></td>
                                <th>매출조직</th>
                                <td class="left"><%=sales_company%>&nbsp;<%=sales_saupbu%>
                                <input type="hidden" name="slip_id" value="<%=slip_id%>" ID="slip_id">
                                <input type="hidden" name="slip_no" value="<%=slip_no%>" ID="slip_no">
                                <input type="hidden" name="slip_seq" value="<%=slip_seq%>" ID="slip_seq">
                                <input type="hidden" name="sales_date" value="<%=sales_date%>" ID="sales_date">
                                <input type="hidden" name="slip_id_view" value="<%=slip_id_view%>" ID="slip_id_view">
                                </td>
                            </tr>
                            <tr>
                                <th>매출거래처</th>
                                <td class="left"><%=trade_name%>&nbsp;</td>
                                <th>사업자번호</th>
                                <td class="left"><%=trade_no%></td>
                                <th>매출처담당</th>
                                <td class="left"><%=trade_person%></td>
                                <th>영업담당</th>
                                <td class="left"><%=emp_name%>&nbsp;(<%=org_name%>)
                                <input type="hidden" name="sales_company" value="<%=sales_company%>" ID="sales_company">
                                <input type="hidden" name="sales_bonbu" value="<%=sales_bonbu%>" ID="sales_bonbu">
                                <input type="hidden" name="sales_saupbu" value="<%=sales_saupbu%>" ID="sales_saupbu">
                                <input type="hidden" name="sales_team" value="<%=sales_team%>" ID="sales_team">
                                <input type="hidden" name="sales_org_name" value="<%=sales_org_name%>" ID="sales_org_name">
                                <input type="hidden" name="emp_no" value="<%=emp_no%>" ID="emp_no">
                                <input type="hidden" name="emp_name" value="<%=emp_name%>" ID="emp_name">
                                <input type="hidden" name="org_name" value="<%=org_name%>" ID="org_name">
                                </td>
 							</tr>
						    <tr>
							    <th>발주일자</th>
							    <td class="left"><input name="order_date" type="text" value="<%=order_date%>" style="width:80px;text-align:center" id="datepicker"></td>
                                <th>발주담당자</th>
							    <td colspan="3" class="left"><%=order_emp_name%>(<%=order_emp_no%>)&nbsp;-&nbsp;<%=order_org_name%>
                                <input type="hidden" name="order_company" value="<%=order_company%>" ID="order_company">
                                <input type="hidden" name="order_bonbu" value="<%=order_bonbu%>" ID="order_bonbu">
                                <input type="hidden" name="order_saupbu" value="<%=order_saupbu%>" ID="order_saupbu">
                                <input type="hidden" name="order_team" value="<%=order_team%>" ID="order_team">
                                <input type="hidden" name="order_org_code" value="<%=order_org_code%>" ID="order_org_code">
                                <input type="hidden" name="order_org_name" value="<%=order_org_name%>" ID="order_org_name">
                                <input type="hidden" name="order_emp_no" value="<%=order_emp_no%>" ID="order_emp_no">
                                <input type="hidden" name="order_emp_name" value="<%=order_emp_name%>" ID="order_emp_name">
                                </td>
                                <th>출고요청일</th>
                                <td class="left"><%=out_request_date%>&nbsp;(<%=out_method%>)
                                <input type="hidden" name="out_request_date" value="<%=out_request_date%>" ID="out_request_date">
                                <input type="hidden" name="out_method" value="<%=out_method%>" ID="out_method">
                                </td>
                            <tr>
							    <th>발주거래처</th>
							    <td class="left"><input name="trade_name" type="text" value="<%=order_trade_name%>" readonly="true" style="width:120px">
						        <a href="#" class="btnType03" onClick="pop_Window('insa_trade_select.asp?gubun=<%="sale"%>','trade_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                                </td>
							    <th>사업자번호</th>
							    <td class="left"><input name="trade_no" type="text" value="<%=order_trade_no%>" readonly="true" style="width:150px"></td>
							    <th>발주처<br>담당자</th>
							    <td class="left"><input name="trade_person" type="text" value="<%=order_trade_person%>" id="trade_person" style="width:120px; ime-mode:active" onKeyUp="checklength(this,20);"></td>
                                <th>이메일</th>
							    <td class="left"><input name="trade_email" type="text" value="<%=order_trade_email%>" id="trade_email" style="width:180px;" onKeyUp="checklength(this,30);"></td>
						    </tr>
                            <tr>
                                <th>대금<br>지급방법</th>
							    <td colspan="5" class="left">
                                <input type="radio" name="bill_collect" value="현금" <% if order_bill_collect = "현금" then %>checked<% end if %> style="width:40px" id="Radio3">현금
  							    <input type="radio" name="bill_collect" value="어음" <% if order_bill_collect = "어음" then %>checked<% end if %> style="width:40px" id="Radio4">어음
                                <input type="radio" name="bill_collect" value="카드" <% if order_bill_collect = "카드" then %>checked<% end if %> style="width:40px" id="Radio3">카드
  							    <input type="radio" name="bill_collect" value="외환" <% if order_bill_collect = "외환" then %>checked<% end if %> style="width:40px" id="Radio4">외환
                                </td>
							    <th>지급예정일</th>
							    <td class="left"><input name="collect_due_date" type="text" value="<%=order_collect_due_date%>" style="width:80px;text-align:center" id="datepicker1"></td>
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
							    <td class="left"><input name="order_in_date" type="text" value="<%=order_in_date%>" style="width:80px;text-align:center" id="datepicker2"></td>
						    </tr>
                            <tr>
                                <th>영업의견</th>
                                <td colspan="7" class="left"><%=slip_memo%>&nbsp;
                                <input type="hidden" name="slip_memo" value="<%=slip_memo%>" ID="slip_memo">
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
                <h3 class="stit" style="font-size:12px;">◈ 품목 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="14%" >
							<col width="14%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
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
                                <input type="hidden" name="b_qty<%=i%>" value="<%=formatnumber(b_qty(i),0)%>" ID="Hidden1">
                                </td>
              <% if  b_qty(i) = 0 then  %>
                                <td align="right"><%=formatnumber(unit_cost(i),0)%></td>
                                <td align="right"><%=formatnumber(b_qty(i),0)%></td>
              <%     else               %>               
                                <td><input name="buy_cost<%=i%>" type="text" id="buy_cost<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(unit_cost(i),0)%>" onKeyUp="NumCal(this);"></td>
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
							  <th>매출품의 첨부</th>
							  <td colspan="5" class="left">
                        <% 
                           If buy_att_file <> "" Then 
                              path = "/sales_file/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=att_file%>"><%=att_file%></a>
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
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                
                <input type="hidden" name="order_id" value="<%=order_id%>">
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
