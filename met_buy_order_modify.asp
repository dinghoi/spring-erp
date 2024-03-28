<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim pummok_tab(6,20)
dim amount_tab(4,20)

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

order_no=Request("order_no")
order_seq = request("order_seq")
order_date=Request("order_date")

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

if u_type = "U" then

	Sql="select * from met_order where (order_no = '"&order_no&"') and (order_seq = '"&order_seq&"') and (order_date = '"&order_date&"')"
	Set rs=DbConn.Execute(Sql)

	order_no = rs("order_no")
	order_seq = rs("order_seq")
	order_date = rs("order_date")
	order_buy_no = rs("order_buy_no")
	order_buy_seq = rs("order_buy_seq")
	order_buy_date = rs("order_buy_date")

	order_goods_type = rs("order_goods_type")
	order_company = rs("order_company")
    order_bonbu = rs("order_bonbu")
	order_saupbu = rs("order_saupbu")
	order_team = rs("order_team")
    order_org_code = rs("order_org_code")
    order_org_name = rs("order_org_name")
	order_emp_no = rs("order_emp_no")
    order_emp_name = rs("order_emp_name")
    order_bill_collect = rs("order_bill_collect")
    order_collect_due_date = rs("order_collect_due_date")
'	response.write(order_collect_due_date)
	
    order_trade_no = rs("order_trade_no")
    order_trade_name = rs("order_trade_name")
    order_trade_person = rs("order_trade_person")
	order_trade_email = rs("order_trade_email")
    order_in_date = rs("order_in_date")
    order_stock_company = rs("order_stock_company")
    order_stock_code = rs("order_stock_code")
    order_stock_name = rs("order_stock_name")
    order_out_method = ""
    order_out_request_date = ""
    order_price = rs("order_price")
    order_cost = rs("order_cost")
    order_cost_vat = rs("order_cost_vat")
    order_memo = rs("order_memo")
    order_ing = rs("order_ing")

	if order_out_request_date = "0000-00-00" then
	      order_out_request_date = ""
	end if
	
	if order_collect_due_date = "0000-00-00" then
	      order_collect_due_date = ""
	end if
	
	if order_in_date = "0000-00-00" then
	      order_in_date = ""
	end if
	
	rs.close()
    
	j = 0
	Sql="select * from met_order_goods where (og_order_no = '"&order_no&"') and (og_order_seq = '"&order_seq&"') and (og_order_date = '"&order_date&"')"
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		j = j + 1
		pummok_tab(1,j) = rs("og_goods_type")
		pummok_tab(2,j) = rs("og_goods_gubun")
		pummok_tab(3,j) = rs("og_goods_code")
		pummok_tab(4,j) = rs("og_goods_name")
		pummok_tab(5,j) = rs("og_standard")
		'pummok_tab(6,j) = rs("og_buy_seq")
		amount_tab(1,j) = rs("og_bg_qty")
		amount_tab(2,j) = rs("og_qty")
		amount_tab(3,j) = rs("og_unit_cost")
		amount_tab(4,j) = rs("og_amt")
		rs.movenext()
	loop
	pummok_cnt = j
	
	title_line = order_goods_type + " 발주 변경 "
	
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
												$( "#datepicker4" ).datepicker("setDate", "<%=order_collect_due_date%>" );
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
			var bqty_ary = new Array();
			var qty_ary = new Array();
			var buy_ary = new Array();
			var buy_tot = new Array();

			for (j=1;j<21;j++) {
				bqty_ary[j] = eval("document.frm.bqty" + j + ".value").replace(/,/g,"");
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
				buy_ary[j] = eval("document.frm.buy_cost" + j + ".value").replace(/,/g,"");
				
				acpt_qty = parseInt(qty_ary[j]);
				sign_qty = parseInt(bqty_ary[j]);
				
			
				if (acpt_qty > sign_qty) {
					alert ("의뢰수량보다 발주수량이 많습니다!!");
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
		function pummok_list_view() {
				pummok_cnt = parseInt(document.frm.pummok_cnt.value);
				for (j=1;j<pummok_cnt+1;j++) {
					eval("document.getElementById('pummok_list" + j + "')").style.display = '';
				}
				NumCal();
			}
		function delcheck() 
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.method = "post";
//					document.frm.enctype = "multipart/form-data";
					document.frm.action = "met_buy_order_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
			}
		</script>

	</head>
	<body onload="pummok_list_view();">
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_buy_order_modify_save.asp">
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
							  <th>발주일자</th>
							  <td class="left"><input name="order_date" type="text" value="<%=order_date%>" style="width:80px;text-align:center" id="datepicker1"></td>
							  <th>발주거래처</th>
							  <td class="left"><input name="trade_name" type="text" value="<%=order_trade_name%>" readonly="true" style="width:120px">
						      <a href="#" class="btnType03" onClick="pop_Window('insa_trade_select.asp?gubun=<%="buy"%>','trade_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              </td>
							  <th>사업자번호</th>
							  <td class="left"><input name="trade_no" type="text" value="<%=order_trade_no%>" readonly="true" style="width:150px"></td>
							  <th>거래처<br>담당자</th>
							  <td class="left"><input name="trade_person" type="text" value="<%=order_trade_person%>" id="trade_person" style="width:120px; ime-mode:active" onKeyUp="checklength(this,20);"></td>
						    </tr>
                            <tr>
							  <th>이메일</th>
							  <td class="left"><input name="trade_email" type="text" value="<%=order_trade_email%>" id="trade_email" style="width:180px;" onKeyUp="checklength(this,30);"></td>
                              <th>대금<br>지급방법</th>
							  <td colspan="3" class="left">
                              <input type="radio" name="bill_collect" value="현금" <% if order_bill_collect = "현금" then %>checked<% end if %> style="width:40px" id="Radio3">현금
  							  <input type="radio" name="bill_collect" value="어음" <% if order_bill_collect = "어음" then %>checked<% end if %> style="width:40px" id="Radio4">어음
                              <input type="radio" name="bill_collect" value="카드" <% if order_bill_collect = "카드" then %>checked<% end if %> style="width:40px" id="Radio3">카드
  							  <input type="radio" name="bill_collect" value="외환" <% if order_bill_collect = "외환" then %>checked<% end if %> style="width:40px" id="Radio4">외환
                              </td>
							  <th>지급예정일</th>
							  <td class="left"><input name="collect_due_date" type="text" value="<%=order_collect_due_date%>" style="width:80px;text-align:center" id="datepicker4"></td>
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
							  <th>비고</th>
							  <td class="left" colspan="7" ><textarea name="order_memo" rows="3" id="textarea"><%=order_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 구매요청 세부 내용 ◈</h3>
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
							<col width="11%" >
							<col width="11%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">선택</th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
                                <th scope="col">구매품의</th>
								<th scope="col">발주수량</th>
								<th scope="col">발주단가</th>
								<th scope="col">발주금액</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><%=i%></td>
								<td><%=pummok_tab(1,i)%>
                                <input type="hidden" name="srv_type<%=i%>" value="<%=pummok_tab(1,i)%>" ID="Hidden1">
                                <input type="hidden" name="bg_seq<%=i%>" value="<%=pummok_tab(6,i)%>" ID="Hidden1">
                                </td>
                                <td><%=pummok_tab(2,i)%>
                                <input type="hidden" name="goods_gubun<%=i%>" value="<%=pummok_tab(2,i)%>" ID="Hidden1">
                                </td>
                                <td><%=pummok_tab(3,i)%>
                                <input type="hidden" name="goods_code<%=i%>" value="<%=pummok_tab(3,i)%>" ID="Hidden1">
                                </td>
								<td><%=pummok_tab(4,i)%>
                                <input type="hidden" name="goods_name<%=i%>" value="<%=pummok_tab(4,i)%>" ID="Hidden1">
								</td>
                                <td><%=pummok_tab(5,i)%>
                                <input type="hidden" name="goods_standard<%=i%>" value="<%=pummok_tab(5,i)%>" ID="Hidden1">
                                </td>
								<td><%=formatnumber(amount_tab(1,i),0)%>
                                <input type="hidden" name="bqty<%=i%>" value="<%=formatnumber(amount_tab(1,i),0)%>" ID="Hidden1">
                                </td>
                                <td>
                                <input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(amount_tab(2,i),0)%>" onChange="NumCal(this);">
                                <input type="hidden" name="old_qty<%=i%>" value="<%=formatnumber(amount_tab(2,i),0)%>" ID="Hidden1">
                                </td>
								<td><%=formatnumber(amount_tab(3,i),0)%>
                                <input type="hidden" name="buy_cost<%=i%>" value="<%=formatnumber(amount_tab(3,i),0)%>" ID="Hidden1">
                                </td>
								<td>
                                <input name="buy_tot<%=i%>" type="text" id="buy_tot<%=i%>" style="width:80px;text-align:right" readonly="true" value="<%=formatnumber(amount_tab(4,i),0)%>">
                                </td>
							</tr>
						<%
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
							<tr>
							  <th>구매요청첨부</th>
							  <td colspan="5" class="left">
                        <% 
                           If buy_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=buy_att_file%>"><img src="image/att_file.gif" border="0"></a>
                        <%    Else %>
				                    &nbsp;
                        <% 
						   End If %>
                              </td>
						    </tr>
						</tbody>
					</table>
<br>
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
                
                <input type="hidden" name="old_buy_no" value="<%=order_buy_no%>">
                <input type="hidden" name="old_buy_seq" value="<%=order_buy_seq%>">
				<input type="hidden" name="old_buy_date" value="<%=order_buy_date%>">
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
				</form>
		   </div>	
        </div>				
	</body>  
</html>
