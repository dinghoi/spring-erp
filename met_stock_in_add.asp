<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim code_tab(20)
dim goods_name(20)
dim goods_type(20)
dim goods_gubun(20)
dim goods_standard(20)
dim qty_tab(20)
dim unit_cost(20)
dim buy_amt(20)
dim bg_seq_tab(20)

dim pummok_tab(5,20)
dim amount_tab(3,20)

for i = 1 to 20
    code_tab(i) = ""
	goods_name(i) = ""
	goods_type(i) = ""
	goods_gubun(i) = ""
	goods_standard(i) = ""
	qty_tab(i) = 0
	unit_cost(i) = 0
	buy_amt(i) = 0
	bg_seq_tab(i) = ""
next

for i = 1 to 5
	for j = 1 to 20
		pummok_tab(i,j) = ""
	next
next
for i = 1 to 3
	for j = 1 to 20
		amount_tab(i,j) = 0
	next
next

mok_cnt = 0
pummok_cnt = 0

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
user_org_name = request.cookies("nkpmg_user")("coo_org_name")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")
user_company = request.cookies("nkpmg_user")("coo_emp_company")

u_type = request("u_type")

order_no = request("order_no")
order_seq = request("order_seq")
order_date = request("order_date")

stin_in_date = request("stin_in_date")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_bg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

  
   Sql="select * from met_order where (order_no = '"&order_no&"') and (order_seq = '"&order_seq&"') and (order_date = '"&order_date&"')"
   Set Rs_order = DbConn.Execute(SQL)
   if not Rs_order.eof then
     	order_id = Rs_order("order_id")
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
'        order_trade_no = Rs_order("order_trade_no")
		order_trade_no = mid(Rs_order("order_trade_no"),1,3) + "-" + mid(Rs_order("order_trade_no"),4,2) + "-" + right(Rs_order("order_trade_no"),5)
        order_trade_name = Rs_order("order_trade_name")
        order_trade_person = Rs_order("order_trade_person")
		order_trade_email = Rs_order("order_trade_email")
        order_in_date = Rs_order("order_in_date")
        order_stock_company = Rs_order("order_stock_company")
        order_stock_code = Rs_order("order_stock_code")
        order_stock_name = Rs_order("order_stock_name")
        order_out_method = Rs_order("order_out_method")
        order_out_request_date = Rs_order("order_out_request_date")
        order_price = Rs_order("order_price")
        order_cost = Rs_order("order_cost")
        order_cost_vat = Rs_order("order_cost_vat")
        order_memo = Rs_order("order_memo")
        order_ing = Rs_order("order_ing")

    	if order_out_request_date = "0000-00-00" then
    	      order_out_request_date = ""
    	end if
		if order_collect_due_date = "0000-00-00" then
    	      order_collect_due_date = ""
    	end if
	
    	if order_in_date = "0000-00-00" then
    	      order_in_date = ""
    	end if
    end if
    Rs_order.close()

    i = 0
    Sql="select * from met_order_goods where (og_order_no = '"&order_no&"') and (og_order_seq = '"&order_seq&"') and (og_order_date = '"&order_date&"')"
    Set Rs_good = DbConn.Execute(SQL)
    do until Rs_good.eof or Rs_good.bof
        i = i +1
	    bg_seq_tab(i) = Rs_good("og_seq")
		goods_type(i) = Rs_good("og_goods_type")
	    goods_gubun(i) = Rs_good("og_goods_gubun")
		code_tab(i) = Rs_good("og_goods_code")
		goods_name(i) = Rs_good("og_goods_name")
		goods_standard(i) = Rs_good("og_standard")
		qty_tab(i) = int(Rs_good("og_qty"))
		unit_cost(i) = int(Rs_good("og_unit_cost"))
		buy_amt(i) = int(Rs_good("og_amt"))

        Rs_good.movenext()
   loop
   mok_cnt = i
   Rs_good.close()
   
title_line = " 입고 등록 "

if u_type = "U" then

	Sql="select * from met_stin where (stin_in_date = '"&stin_in_date&"') and (stin_order_no = '"&order_no&"') and (stin_order_seq = '"&order_seq&"')"
	Set rs=DbConn.Execute(Sql)

	stin_order_no = rs("stin_order_no")
	stin_order_seq = rs("stin_order_seq")
	stin_order_date = rs("stin_order_date")
	stin_order_date = rs("stin_order_date")
	stin_buy_no = rs("stin_buy_no")
	stin_buy_seq = rs("stin_buy_seq")
	stin_buy_date = rs("stin_buy_date")
	
	stin_company = rs("stin_company")
    stin_org_name = rs("stin_org_name")
	stin_emp_no = rs("stin_emp_no")
    stin_emp_name = rs("stin_emp_name")
	
    stin_bill_collect = rs("stin_bill_collect")
    stin_collect_due_date = rs("stin_collect_due_date")
    stin_trade_no = rs("stin_trade_no")
    stin_trade_name = rs("stin_trade_name")
    stin_trade_person = rs("stin_trade_person")
	stin_trade_email = rs("stin_trade_email")

    stin_stock_company = rs("stin_stock_company")
    stin_stock_code = rs("stin_stock_code")
    stin_stock_name = rs("stin_stock_name")
	
    stin_price = rs("stin_price")
    stin_cost = rs("stin_cost")
    stin_cost_vat = rs("stin_cost_vat")
	
    stin_id = rs("stin_id")

	if order_out_request_date = "0000-00-00" then
	      order_out_request_date = ""
	end if
	
	if order_in_date = "0000-00-00" then
	      order_in_date = ""
	end if
	
	rs.close()
	
	j = 0
	Sql="select * from met_stin_goods where (stin_date = '"&stin_in_date&"') and (stin_order_no = '"&order_no&"') and (stin_order_seq = '"&order_buy_no&"')"
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		j = j + 1
		pummok_tab(1,j) = rs("stin_goods_type")
		pummok_tab(2,j) = rs("stin_goods_gubun")
		pummok_tab(3,j) = rs("stin_goods_code")
		pummok_tab(4,j) = rs("stin_goods_name")
		pummok_tab(5,j) = rs("stin_standard")

		amount_tab(1,j) = rs("stin_qty")
		amount_tab(2,j) = rs("stin_unit_cost")
		amount_tab(3,j) = rs("stin_amt")
		rs.movenext()
	loop
	pummok_cnt = j
	mok_cnt = j
    
	title_line = " 입고 변경 "
	
end if

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
				return "1 1";
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
//				if(document.frm.sales_saupbu.value == "") {
//					alert('매출사업부를 선택하세요');
//					frm.sales_saupbu.focus();
//					return false;}
				if(document.frm.stin_stock_name.value == "") {
					alert('입고창고를 선택하세요');
					frm.stin_stock_name.focus();
					return false;}

					
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function pop_pummok() 
			{ 
				var code_ary = new Array();
				code_ary[0] = document.frm.goods_code1.value
				code_ary[1] = document.frm.goods_code2.value
				code_ary[2] = document.frm.goods_code3.value
				code_ary[3] = document.frm.goods_code4.value
				code_ary[4] = document.frm.goods_code5.value
				code_ary[5] = document.frm.goods_code6.value
				code_ary[6] = document.frm.goods_code7.value
				code_ary[7] = document.frm.goods_code8.value
				code_ary[8] = document.frm.goods_code9.value
				code_ary[9] = document.frm.goods_code10.value
				var popupW = 600;
				var popupH = 400;
				var left = Math.ceil((window.screen.width - popupW)/2);
				var top = Math.ceil((window.screen.height - popupH)/2);
				window.open('met_goods_select.asp?code_ary='+code_ary+'', '팝업공지', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
			} 
			function pop_pummok_del() 
			{ 
				var code_ary = new Array();
				var del_ary = new Array();
				code_ary[0] = document.frm.goods_code1.value
				code_ary[1] = document.frm.goods_code2.value
				code_ary[2] = document.frm.goods_code3.value
				code_ary[3] = document.frm.goods_code4.value
				code_ary[4] = document.frm.goods_code5.value
				code_ary[5] = document.frm.goods_code6.value
				code_ary[6] = document.frm.goods_code7.value
				code_ary[7] = document.frm.goods_code8.value
				code_ary[8] = document.frm.goods_code9.value
				code_ary[9] = document.frm.goods_code10.value

				if (document.frm.del_check1.checked == true) {
					del_ary[0] = 'Y'; } 					
				if (document.frm.del_check1.checked == false) {
					del_ary[0] = 'N'; } 
				if (document.frm.del_check2.checked == true) {
					del_ary[1] = 'Y'; } 					
				if (document.frm.del_check2.checked == false) {
					del_ary[1] = 'N'; } 
				if (document.frm.del_check3.checked == true) {
					del_ary[2] = 'Y'; } 					
				if (document.frm.del_check3.checked == false) {
					del_ary[2] = 'N'; } 
				if (document.frm.del_check4.checked == true) {
					del_ary[3] = 'Y'; } 					
				if (document.frm.del_check4.checked == false) {
					del_ary[3] = 'N'; } 
				if (document.frm.del_check5.checked == true) {
					del_ary[4] = 'Y'; } 					
				if (document.frm.del_check5.checked == false) {
					del_ary[4] = 'N'; } 
				if (document.frm.del_check6.checked == true) {
					del_ary[5] = 'Y'; } 					
				if (document.frm.del_check6.checked == false) {
					del_ary[5] = 'N'; } 
				if (document.frm.del_check7.checked == true) {
					del_ary[6] = 'Y'; } 					
				if (document.frm.del_check7.checked == false) {
					del_ary[6] = 'N'; } 
				if (document.frm.del_check8.checked == true) {
					del_ary[7] = 'Y'; } 					
				if (document.frm.del_check8.checked == false) {
					del_ary[7] = 'N'; } 
				if (document.frm.del_check9.checked == true) {
					del_ary[8] = 'Y'; } 					
				if (document.frm.del_check9.checked == false) {
					del_ary[8] = 'N'; } 
				if (document.frm.del_check10.checked == true) {
					del_ary[9] = 'Y'; } 					
				if (document.frm.del_check10.checked == false) {
					del_ary[9] = 'N'; } 
				var popupW = 600;
				var popupH = 400;
				var left = Math.ceil((window.screen.width - popupW)/2);
				var top = Math.ceil((window.screen.height - popupH)/2);
				window.open('met_goods_del_ok.asp?code_ary='+code_ary+'&del_ary='+del_ary+'', '팝업공지', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
			} 

		function NumCal(txtObj){
			var oqty_ary = new Array();
			var qty_ary = new Array();
			var buy_ary = new Array();
			var buy_tot = new Array();
			mok_cnt = parseInt(document.frm.mok_cnt.value);

			for (j=1;j<mok_cnt+1;j++) {
				oqty_ary[j] = eval("document.frm.oqty" + j + ".value").replace(/,/g,"");
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
				buy_ary[j] = eval("document.frm.buy_cost" + j + ".value").replace(/,/g,"");
				buy_tot[j] = qty_ary[j] * buy_ary[j];
				
				acpt_qty = parseInt(qty_ary[j]);
				order_qty = parseInt(oqty_ary[j]);
				
//				if (qty_ary[j] != oqty_ary[j]) {
				if (acpt_qty != order_qty) {	
					alert ("발주수량과 입고수량이같지않습니다!!");
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
			for (j=1;j<mok_cnt+1;j++) {
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

		
			if (txtObj.value.length<5) {
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
		</script>

	</head>
	<body>
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_stock_in_add_save.asp">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
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
                                <td class="left"><%=order_buy_no%>&nbsp;<%=order_buy_seq%></td>
                                <th>구매용도</th>
                                <td class="left"><%=order_goods_type%></td>
                                <th>구매<br>품의일자</th>
                                <td class="left"><%=order_buy_date%></td>
                                <th>구매회사</th>
                                <td class="left"><%=order_company%></td>
                                <input type="hidden" name="order_id" value="<%=order_id%>" ID="order_id">
                                <input type="hidden" name="buy_no" value="<%=order_buy_no%>" ID="buy_no">
                                <input type="hidden" name="buy_seq" value="<%=order_buy_seq%>" ID="buy_seq">
                                <input type="hidden" name="buy_date" value="<%=order_buy_date%>" ID="buy_date">
                                <input type="hidden" name="buy_goods_type" value="<%=order_goods_type%>" ID="buy_goods_type">
                                </td>
 							</tr>
                            <tr>
							    <th>사업부</th>
                                <td class="left"><%=order_saupbu%></td>
                                <th>소속</th>
                                <td class="left"><%=order_org_name%></td>
                                <th>구매담당</th>
                                <td colspan="3" class="left"><%=order_emp_name%>&nbsp;(<%=order_emp_no%>)
                                <input type="hidden" name="order_company" value="<%=order_company%>" ID="buy_company">
                                <input type="hidden" name="order_bonbu" value="<%=order_bonbu%>" ID="buy_bonbu">
                                <input type="hidden" name="order_saupbu" value="<%=order_saupbu%>" ID="buy_saupbu">
                                <input type="hidden" name="order_team" value="<%=order_team%>" ID="buy_team">
                                <input type="hidden" name="order_org_code" value="<%=order_org_code%>" ID="buy_org_code">
                                <input type="hidden" name="order_org_name" value="<%=order_org_name%>" ID="buy_org_name">
                                <input type="hidden" name="order_emp_no" value="<%=order_emp_no%>" ID="buy_emp_no">
                                <input type="hidden" name="order_emp_name" value="<%=order_emp_name%>" ID="buy_emp_name">
                                </td>
						    </tr>
                            <tr>
							  <th>발주일자<br>번호</th>
							  <td class="left"><%=order_date%>&nbsp;(<%=order_no%>&nbsp;<%=order_seq%>)
                              <input type="hidden" name="order_date" value="<%=order_date%>" ID="order_date">
                              <input type="hidden" name="order_no" value="<%=order_no%>" ID="order_no">
                              <input type="hidden" name="order_seq" value="<%=order_seq%>" ID="order_seq">
                              </td>
							  <th>발주거래처</th>
							  <td class="left"><%=order_trade_name%>
                              <input type="hidden" name="trade_name" value="<%=order_trade_name%>" ID="trade_name">
                              </td>
							  <th>사업자번호</th>
							  <td class="left"><%=order_trade_no%>
                              <input type="hidden" name="trade_no" value="<%=order_trade_no%>" ID="trade_no">
                              </td>
							  <th>거래처<br>담당자</th>
							  <td class="left"><%=order_trade_person%>
                              <input type="hidden" name="trade_person" value="<%=order_trade_person%>" ID="trade_person">
                              </td>
						    </tr>
                            <tr>
							  <th>이메일</th>
							  <td class="left"><%=order_trade_email%>
                              <input type="hidden" name="trade_email" value="<%=order_trade_email%>" ID="trade_email">
                              </td>
                              <th>대금<br>지급방법</th>
							  <td colspan="3" class="left"><%=order_bill_collect%>
                              <input type="hidden" name="bill_collect" value="<%=order_bill_collect%>" ID="bill_collect">
                              </td>
							  <th>지급예정일</th>
							  <td class="left"><%=order_collect_due_date%>
                              <input type="hidden" name="collect_due_date" value="<%=order_collect_due_date%>" ID="collect_due_date">
                              </td>
						    </tr>
                            <tr>
							  <th>입고일자</th>
							  <td class="left"><input name="stin_in_date" type="text" value="<%=order_in_date%>" style="width:80px;text-align:center" id="datepicker5"></td>
                              <th>입고창고</th>
							  <td colspan="3" class="left">
                              <input name="stin_stock_company" type="text" value="<%=order_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="stin_stock_name" type="text" value="<%=order_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="stin"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="stin_stock_code" value="<%=order_stock_code%>" ID="Hidden1">
                              <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="Hidden1">
                              </td>
                              <th>입고담당</th>
							  <td class="left"><%=user_name%>&nbsp;(<%=user_org_name%>)
                              <input type="hidden" name="emp_no" value="<%=user_id%>" ID="Hidden1">
                              <input type="hidden" name="emp_name" value="<%=user_name%>" ID="Hidden1">
                              <input type="hidden" name="emp_company" value="<%=user_company%>" ID="Hidden1">
                              <input type="hidden" name="emp_org_name" value="<%=user_org_name%>" ID="Hidden1">
                              </td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 발주 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="14%" >
							<col width="12%" >
                            <col width="6%" >
							<col width="11%" >
							<col width="9%" >
							<col width="11%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
                                <th scope="col">발주수량</th>
                                <th scope="col">발주단가</th>
								<th scope="col">입고수량</th>
								<th scope="col">입고금액</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
'							    if code_tab(i) <> "" then
								if code_tab(i) = "" or isnull(code_tab(i)) then 
			                           exit for 
								   else
						%>
			  				<tr id="pummok_list<%=i%>" style="display:">
								<td class="first"><%=i%></td>
								<td><%=goods_type(i)%>
                                <input type="hidden" name="srv_type<%=i%>" value="<%=goods_type(i)%>" ID="Hidden1">
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
								<td><%=formatnumber(qty_tab(i),0)%>
                                <input type="hidden" name="oqty<%=i%>" value="<%=qty_tab(i)%>" ID="Hidden1">
                                </td>
                                <td><%=formatnumber(unit_cost(i),0)%>
                                <input type="hidden" name="buy_cost<%=i%>" value="<%=formatnumber(unit_cost(i),0)%>" ID="Hidden1">
                                </td>
                                <td>
                                <input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" readonly="true" value="<%=formatnumber(qty_tab(i),0)%>">
                                </td>
								<td>
                                <input name="buy_tot<%=i%>" type="text" id="buy_tot<%=i%>" style="width:80px;text-align:right" readonly="true" value="<%=formatnumber(buy_amt(i),0)%>">
                                </td>
							</tr>
						<%
						        end if
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
							  <td class="left"><input name="buy_tot_price" type="text" id="buy_tot_price" style="width:150px;text-align:right" value="<%=formatnumber(order_price,0)%>" readonly="true"></td>
							  <th>발주금액</th>
							  <td class="left"><input name="buy_tot_cost" type="text" id="buy_tot_cost" style="width:150px;text-align:right" value="<%=formatnumber(order_cost,0)%>" readonly="true"></td>
							  <th>부가세</th>
							  <td class="left"><input name="buy_tot_cost_vat" type="text" id="buy_tot_cost_vat" style="width:150px;text-align:right" value="<%=formatnumber(order_cost_vat,0)%>" readonly="true"></td>
						    </tr>
						</tbody>
					</table>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
        
                <input type="hidden" name="old_order_id" value="<%=order_id%>">
                <input type="hidden" name="old_order_date" value="<%=order_date%>">
				<input type="hidden" name="old_order_buy_no" value="<%=order_buy_no%>">
				<input type="hidden" name="old_order_goods_type" value="<%=order_goods_type%>">
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
				</form>
		</div>
	  </div>        				
	</body>  
</html>
