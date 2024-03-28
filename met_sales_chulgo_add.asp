<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim code_tab(20)
dim goods_name(20)
dim goods_type(20)
dim goods_gubun(20)
dim goods_standard(20)
dim goods_grade(20)
dim qty_tab(20)
dim seq_tab(20)
dim chul_qty_tab(20)
dim c_chk_tab(20)
dim c_qty_tab(20)
dim b_qty(20)

dim pummok_tab(6,20)
dim amount_tab(4,20)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
user_org_name = request.cookies("nkpmg_user")("coo_org_name")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")
user_company = request.cookies("nkpmg_user")("coo_emp_company")

u_type = request("u_type") 

slip_id=Request("slip_id")
slip_no=Request("slip_no")
slip_seq=Request("slip_seq")
sales_date = request("sales_date")

curr_date = mid(cstr(now()),1,10)
chulgo_date = curr_date

for i = 1 to 20
    seq_tab(i) = ""
	code_tab(i) = ""
	goods_name(i) = ""
	goods_type(i) = ""
	goods_gubun(i) = ""
	goods_standard(i) = ""
	goods_grade(i) = ""
	qty_tab(i) = 0
	chul_qty_tab(i) = 0
	c_chk_tab(i) = ""
	c_qty_tab(i) = 0
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

mok_cnt = 0
pummok_cnt = 0
' response.write(reg_date)

if slip_id = "2" then
		slip_id_view = "수주전표"
end if
if slip_id = "1" then
		slip_id_view = "대기전표"
end if

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_sale = Server.CreateObject("ADODB.Recordset")
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
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
		
		chulgo_date = out_request_date
   end if
   Rs_buy.close()

   i = 0
   sql = "select * from sales_slip_detail where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"' ORDER BY goods_seq ASC"
   Set Rs_good = DbConn.Execute(SQL)
   do until Rs_good.eof or Rs_good.bof
        i = i +1
	    seq_tab(i) = Rs_good("goods_seq")
		goods_type(i) = Rs_good("srv_type")
'	    goods_gubun(i) = Rs_good("rl_goods_gubun")
		code_tab(i) = Rs_good("goods_code")
		goods_name(i) = Rs_good("pummok")
		goods_standard(i) = Rs_good("standard")
		goods_grade(i) = "A급"
		qty_tab(i) = Rs_good("qty")
'		c_qty_tab(i) = Rs_good("cg_qty")
		b_qty(i) = qty_tab(i) - c_qty_tab(i)
		if qty_tab(i) = c_qty_tab(i) then
		        c_chk_tab(i) = "1"
		   else 
		        c_chk_tab(i) = "0"
		end if
		mok_cnt = i
        Rs_good.movenext()
   loop
   Rs_good.close()

Sql = "SELECT * FROM emp_master where emp_no = '"&user_id&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	chulgo_emp_no = rs_emp("emp_no")
		chulgo_emp_name = rs_emp("emp_name")
		chulgo_company = rs_emp("emp_company")
		chulgo_bonbu = rs_emp("emp_bonbu")
		chulgo_saupbu = rs_emp("emp_saupbu")
		chulgo_team = rs_emp("emp_team")
		chulgo_org_code = rs_emp("emp_org_code")
		chulgo_org_name = rs_emp("emp_org_name")
   else
		chulgo_emp_no = ""
		chulgo_emp_name = ""
		chulgo_company = ""
		chulgo_bonbu = ""
		chulgo_saupbu = ""
		chulgo_team = ""
		chulgo_org_code = ""
		chulgo_org_name = ""
end if
rs_emp.close()

title_line =  " 매출(대기/수주전표) 출고등록 "

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
												$( "#datepicker" ).datepicker("setDate", "<%=chulgo_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=chulgo_date%>" );
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
												$( "#datepicker5" ).datepicker("setDate", "<%=collect_date%>" );
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
				if(document.frm.chulgo_date.value == "") {
					alert('실출고일를 입력하세요');
					frm.chulgo_date.focus();
					return false;}
                
//				k = 0;
//				for (j=1;j<21;j++) {
//					if (eval("document.frm.qty[" + j + "].value") < eval("document.frm.chul_qty[" + j + "].value")) {
//						k = k + 1
//					}
//				}
//				if (k != 0) {
//					alert ("의뢰수량보다 출고수량이 더 많습니다");
//					frm.chul_qty1.focus();
//					return false;
//				}	
					
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
			var chul_qty = new Array();
			mok_cnt = parseInt(document.frm.mok_cnt.value);
            
			for (j=1;j<21;j++) {
				bqty_ary[j] = eval("document.frm.b_qty" + j + ".value").replace(/,/g,"");
				chul_qty[j] = eval("document.frm.chul_qty" + j + ".value").replace(/,/g,"");
				
				acpt_qty = parseInt(chul_qty[j]);
				sign_qty = parseInt(bqty_ary[j]);
				
//				if (chul_qty[j] > bqty_ary[j]) {
				if (acpt_qty > sign_qty) {
					alert ("의뢰수량보다 출고수량이 많습니다!!");
					return false;
				}
		        
			}

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
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_sales_chulgo_add_save.asp">
				<div class="gView">
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
                                <input type="hidden" name="trade_name" value="<%=trade_name%>" ID="trade_name">
                                </td>
                                <th>출고요청일</th>
                                <td class="left"><%=out_request_date%>&nbsp;(<%=out_method%>)
                                <input type="hidden" name="out_request_date" value="<%=out_request_date%>" ID="out_request_date">
                                <input type="hidden" name="out_method" value="<%=out_method%>" ID="out_method">
                                </td>
 							</tr>
                            <tr>
                                <th>실출고일</th>
							    <td class="left"><input name="chulgo_date" type="text" value="<%=chulgo_date%>" style="width:80px;text-align:center" id="datepicker"></td>
                                <th>출고창고</th>
							    <td colspan="3" class="left">
                                <input name="chulgo_stock_company" type="text" value="<%=chulgo_stock_company%>" readonly="true" style="width:120px">
                              -
                                <input name="chulgo_stock_name" type="text" value="<%=chulgo_stock_name%>" readonly="true" style="width:120px">
                              
						        <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="sale"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                                <input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>" ID="chulgo_stock">
                                <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="stock_bonbu">
                                <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="stock_bonbu">
                                <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="stock_team">
                                <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="stock_manager_code">
                                <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="stock_manager_name">
                                </td>
							    <th>출고담당</th>
							    <td class="left"><%=chulgo_emp_name%>(<%=chulgo_emp_no%>)
                                <input type="hidden" name="chulgo_emp_no" value="<%=chulgo_emp_no%>" ID="chulgo_emp_no">
                                <input type="hidden" name="chulgo_emp_name" value="<%=chulgo_emp_name%>" ID="chulgo_emp_name">
                                <input type="hidden" name="chulgo_company" value="<%=chulgo_company%>" ID="chulgo_company">
                                <input type="hidden" name="chulgo_bonbu" value="<%=chulgo_bonbu%>" ID="chulgo_bonbu">
                                <input type="hidden" name="chulgo_saupbu" value="<%=chulgo_saupbu%>" ID="chulgo_saupbu">
                                <input type="hidden" name="chulgo_team" value="<%=chulgo_team%>" ID="chulgo_team">
                                <input type="hidden" name="chulgo_org_name" value="<%=chulgo_org_name%>" ID="chulgo_org_name">
                              </td>
						    </tr>
                            <tr>
                                <th>영업의견</th>
                                <td colspan="7" class="left"><%=slip_memo%>&nbsp;
                                <input type="hidden" name="slip_memo" value="<%=slip_memo%>" ID="slip_memo">
                                </td>
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
                            <col width="8%" >
                            <col width="*" >
                            <col width="12%" >
							<col width="16%" >
							<col width="16%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
                                <th scope="col">용도구분</th>
                                <th scope="col">상태</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
								<th scope="col">의뢰수량</th>
                                <th scope="col">기출고</th>
                                <th scope="col">출고수량</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
							    if code_tab(i) = "" or isnull(code_tab(i)) then 
			                           exit for
		                           else
						%>
			  				<tr id="pummok_list<%=i%>" style="display:">   
								<td class="first"><%=i%></td>
                                <td><%=goods_type(i)%>
                                <input type="hidden" name="srv_type<%=i%>" value="<%=goods_type(i)%>" id="srv_type<%=i%>">
                                <input type="hidden" name="c_chk<%=i%>" value="<%=c_chk_tab(i)%>" id="c_chk<%=i%>">
                                <input type="hidden" name="bg_seq<%=i%>" value="<%=seq_tab(i)%>" ID="Hidden1">
                                </td>
                                <td><%=goods_grade(i)%>
                                <input type="hidden" name="goods_grade<%=i%>" value="<%=goods_grade(i)%>" id="goods_grade<%=i%>">
                                </td>
                                <td><%=goods_gubun(i)%>
                                <input type="hidden" name="goods_gubun<%=i%>" value="<%=goods_gubun(i)%>" id="goods_gubun<%=i%>">
                                </td>
                                <td><%=code_tab(i)%>
                                <input type="hidden" name="goods_code<%=i%>" value="<%=code_tab(i)%>" id="goods_code<%=i%>">
                                </td>
                                <td><%=goods_name(i)%>
                                <input type="hidden" name="goods_name<%=i%>" value="<%=goods_name(i)%>" id="goods_name<%=i%>">
								</td>
                                <td><%=goods_standard(i)%>
                                <input type="hidden" name="goods_standard<%=i%>" value="<%=goods_standard(i)%>" id="goods_standard<%=i%>">
                                </td>
								<td align="right"><%=formatnumber(qty_tab(i),0)%>
                                <input type="hidden" name="qty<%=i%>" value="<%=formatnumber(qty_tab(i),0)%>" ID="Hidden1">
                                </td>
                                <td align="right"><%=formatnumber(c_qty_tab(i),0)%>
                                <input type="hidden" name="c_qty<%=i%>" value="<%=formatnumber(c_qty_tab(i),0)%>" ID="Hidden1">
                                <input type="hidden" name="b_qty<%=i%>" value="<%=formatnumber(b_qty(i),0)%>" ID="Hidden1">
                                </td>
              <% if  b_qty(i) = 0 then  %>
                                <td align="right"><%=formatnumber(b_qty(i),0)%>
                                <input type="hidden" name="chul_qty<%=i%>" value="<%=formatnumber(b_qty(i),0)%>" ID="Hidden1">
                                </td>
              <%     else               %>               
                                <td><input name="chul_qty<%=i%>" type="text" id="chul_qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(b_qty(i),0)%>" onChange="NumCal(this);"></td>
              <% end if                 %>                                     
							</tr>
						<%     
						        end if
							next
						%>
						</tbody>
					</table>                    
					<br>
				</div>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                
                <input type="hidden" name="old_chulgo_date" value="<%=chulgo_date%>">
                <input type="hidden" name="old_chulgo_stock" value="<%=chulgo_stock%>">
				<input type="hidden" name="old_chulgo_seq" value="<%=chulgo_seq%>">
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
				</form>
		</div>				
	</body>
</html>
