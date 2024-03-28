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

dim service_no(20)
dim trade_name(20)
dim trade_dept(20)
dim r_bigo(20)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
user_org_name = request.cookies("nkpmg_user")("coo_org_name")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")
user_company = request.cookies("nkpmg_user")("coo_emp_company")

u_type = request("u_type") 

rele_no = request("rele_no")
rele_seq = request("rele_seq")
rele_date = request("rele_date")

curr_date = mid(cstr(now()),1,10)
order_date = curr_date
order_in_date = curr_date

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
	
	service_no(i) = ""
	trade_name(i) = ""
	trade_dept(i) = ""
	r_bigo(i) = ""
next

mok_cnt = 0
pummok_cnt = 0
' response.write(reg_date)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

'출고의뢰 조회
sql = "select * from met_chulgo_reg where (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"') and (rele_date = '"&rele_date&"')"
Set Rs_reg = DbConn.Execute(SQL)
if not Rs_reg.eof then
    	rele_no = Rs_reg("rele_no")
        rele_seq = Rs_reg("rele_seq")
	    rele_date = Rs_reg("rele_date")
        rele_id = Rs_reg("rele_id")
        rele_goods_type = Rs_reg("rele_goods_type")
		
		rele_stock = Rs_reg("rele_stock")
        rele_stock_company = Rs_reg("rele_stock_company")
        rele_stock_name = Rs_reg("rele_stock_name")
		
        rele_emp_no = Rs_reg("rele_emp_no")
        rele_emp_name = Rs_reg("rele_emp_name")
        rele_company = Rs_reg("rele_company")
        rele_bonbu = Rs_reg("rele_bonbu")
        rele_saupbu = Rs_reg("rele_saupbu")
        rele_team = Rs_reg("rele_team")
        rele_org_name = Rs_reg("rele_org_name")
        rele_trade_name = Rs_reg("rele_trade_name")
	    rele_trade_dept = Rs_reg("rele_trade_dept")
	    rele_delivery = Rs_reg("rele_delivery")
        service_acpt = Rs_reg("service_no")
        chulgo_ing = Rs_reg("chulgo_ing")
        chulgo_date = Rs_reg("chulgo_date")
        chulgo_stock = Rs_reg("chulgo_stock")
        chulgo_stock_name = Rs_reg("chulgo_stock_name")
	    chulgo_stock_company = Rs_reg("chulgo_stock_company")
	    rele_att_file = Rs_reg("rele_att_file")
	    rele_memo = Rs_reg("rele_memo")
        rele_sign_yn = Rs_reg("rele_sign_yn")
	    rele_sign_no = Rs_reg("rele_sign_no")
	    rele_sign_date = Rs_reg("rele_sign_date")
	    if chulgo_date = "0000-00-00" then
	          chulgo_date = ""
	    end if
   else
		rele_no = ""
        rele_seq = ""
	    rele_date = ""
        rele_id = ""
        rele_goods_type = ""
		
		rele_stock = ""
        rele_stock_company = ""
        rele_stock_name = ""
		
        rele_emp_no = ""
        rele_emp_name = ""
        rele_company = ""
        rele_bonbu = ""
        rele_saupbu = ""
        rele_team = ""
        rele_org_name = ""
        rele_trade_name = ""
	    rele_trade_dept = ""
	    rele_delivery = ""
        service_acpt = ""
        chulgo_ing = ""
        chulgo_date = ""
        chulgo_stock = ""
        chulgo_stock_name = ""
	    chulgo_stock_company = ""
	    rele_att_file = ""
	    rele_memo = ""
        rele_sign_yn = ""
	    rele_sign_no = ""
	    rele_sign_date = ""
end if
Rs_reg.close()

i = 0
sql = "select * from met_chulgo_reg_goods where (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"') and (rele_date = '"&rele_date&"')  ORDER BY rl_goods_seq,rl_goods_code ASC"
Set Rs_good = DbConn.Execute(SQL)
do until Rs_good.eof or Rs_good.bof
        i = i +1
	    seq_tab(i) = Rs_good("rl_goods_seq")
		goods_type(i) = Rs_good("rl_goods_type")
	    goods_gubun(i) = Rs_good("rl_goods_gubun")
		code_tab(i) = Rs_good("rl_goods_code")
		goods_name(i) = Rs_good("rl_goods_name")
		goods_standard(i) = Rs_good("rl_standard")
		goods_grade(i) = Rs_good("rl_goods_grade")
		qty_tab(i) = Rs_good("rl_qty")
		c_qty_tab(i) = Rs_good("cg_qty")
		b_qty(i) = qty_tab(i) - c_qty_tab(i)
		if qty_tab(i) = c_qty_tab(i) then
		        c_chk_tab(i) = "1"
		   else 
		        c_chk_tab(i) = "0"
		end if
		
		service_no(i) = Rs_good("rl_service_no")
		trade_name(i) = Rs_good("rl_trade_name")
		trade_dept(i) = Rs_good("rl_trade_dept")
		r_bigo(i) = Rs_good("rl_bigo")

        Rs_good.movenext()
loop

mok_cnt = i
pummok_cnt = i

Rs_good.close()

Sql = "SELECT * FROM emp_master where emp_no = '"&user_id&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_no = rs_emp("emp_no")
		emp_name = rs_emp("emp_name")
		emp_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
		emp_team = rs_emp("emp_team")
		emp_org_code = rs_emp("emp_org_code")
		emp_org_name = rs_emp("emp_org_name")
   else
		emp_name = ""
		emp_company = ""
		emp_bonbu = ""
		emp_saupbu = ""
		emp_team = ""
		emp_org_code = ""
		emp_org_name = ""
end if
rs_emp.close()

chulgo_emp_no = emp_no
chulgo_emp_name = emp_name
chulgo_company = emp_company
chulgo_bonbu = emp_bonbu
chulgo_saupbu = emp_saupbu
chulgo_team = emp_team
chulgo_org_name = emp_org_name

title_line =  " 본사출고의뢰 출고등록 "

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
		function pummok_list_view() {
				pummok_cnt = parseInt(document.frm.pummok_cnt.value);
				for (j=1;j<pummok_cnt+1;j++) {
					eval("document.getElementById('pummok_list" + j + "')").style.display = '';
				}
				NumCal();
			}
		</script>

	</head>
	<body onload="pummok_list_view();">
		    <div id="container">				
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_chulgo_cust_add_save.asp">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="*" >
							<col width="12%" >
							<col width="20%" >
							<col width="12%" >
							<col width="20%" >
						</colgroup>
						<tbody>
							<tr>
							  <th>의뢰회사</th>
							  <td class="left"><%=rele_company%>&nbsp;</td>
							  <th>의뢰자소속</th>
							  <td class="left"><%=rele_saupbu%>&nbsp;(<%=rele_org_name%>)</td>
							  <th>신청자</th>
							  <td class="left"><%=rele_emp_name%>(<%=rele_emp_no%>)
                                <input type="hidden" name="rele_company" value="<%=rele_company%>" ID="rele_company">
                                <input type="hidden" name="rele_saupbu" value="<%=rele_saupbu%>" ID="rele_saupbu">
                                <input type="hidden" name="rele_org_name" value="<%=rele_org_name%>" ID="rele_org_name">
                                <input type="hidden" name="rele_emp_no" value="<%=rele_emp_no%>" ID="rele_emp_no">
                                <input type="hidden" name="rele_emp_name" value="<%=rele_emp_name%>" ID="rele_emp_name">
                              </td>
						    </tr>
                            <tr>
							    <th>출고의뢰번호</th>
							    <td class="left"><%=rele_no%>&nbsp;<%=rele_seq%></td>
							    <th>의뢰일자</th>
							    <td class="left"><%=rele_date%></td>
							    <th>의뢰창고</th>
							    <td class="left"><%=rele_stock_name%>
                                  <input type="hidden" name="rele_no" value="<%=rele_no%>" ID="rele_no">
                                  <input type="hidden" name="rele_seq" value="<%=rele_seq%>" ID="rele_seq">
                                  <input type="hidden" name="rele_date" value="<%=rele_date%>" ID="rele_date">
                                  <input type="hidden" name="rele_goods_type" value="<%=rele_goods_type%>" ID="rele_goods_type">
                                  <input type="hidden" name="rele_stock" value="<%=rele_stock%>" ID="rele_stock">
                                  <input type="hidden" name="rele_stock_company" value="<%=rele_stock_company%>" ID="rele_stock_company">
                                  <input type="hidden" name="rele_stock_name" value="<%=rele_stock_name%>" ID="rele_stock_name">
                                </td>
						    </tr>
							<tr>
                              <th>출고요청일</th>
							  <td colspan="5" class="left"><%=chulgo_date%>&nbsp;
                                <input type="hidden" name="service_no" value="<%=service_acpt%>" ID="service_no">
                                <input type="hidden" name="rele_trade_name" value="<%=rele_trade_name%>" ID="rele_trade_name">
                                <input type="hidden" name="rele_trade_dept" value="<%=rele_trade_dept%>" ID="rele_trade_dept">
                                <input type="hidden" name="rele_chulgo_date" value="<%=chulgo_date%>" ID="rele_chulgo_date">
                              </td>
						    </tr>
                            <tr>
							  <th>실출고일</th>
							  <td class="left"><input name="chulgo_date" type="text" value="<%=chulgo_date%>" style="width:80px;text-align:center" id="datepicker"></td>
							  <th>출고창고</th>
							  <td class="left"><%=chulgo_stock_name%>&nbsp;(<%=chulgo_stock_company%>)
                                <input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>" ID="Hidden1">
                                <input type="hidden" name="chulgo_stock_name" value="<%=chulgo_stock_name%>" ID="Hidden1">
                                <input type="hidden" name="chulgo_stock_company" value="<%=chulgo_stock_company%>" ID="Hidden1">
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
                                <th>구매의견</th>
                                <td colspan="5" class="left"><%=rele_memo%>&nbsp;
                                <input type="hidden" name="rele_memo" value="<%=rele_memo%>" ID="rele_memo">
                                </td>
                            </tr>
						</tbody>
					</table>
				</div>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 출고의뢰 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="8%" >
                            <col width="8%" >
							<col width="12%" >
                            
                            <col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="*" >
                            
							<col width="6%" >
							<col width="6%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
                                <th scope="col">용도구분</th>
                                <th scope="col">상태</th>
                                <th scope="col">품목구분</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
                                
                                <th scope="col">서비스No</th>
                                <th scope="col">고객사</th>
                                <th scope="col">점소명</th>
                                <th scope="col">비고(사유)</th>
                                
								<th scope="col">의뢰수량</th>
                                <th scope="col">기출고</th>
                                <th scope="col">출고수량</th>
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
                                <td><%=goods_name(i)%>
                                <input type="hidden" name="goods_code<%=i%>" value="<%=code_tab(i)%>" id="goods_code<%=i%>">
                                <input type="hidden" name="goods_name<%=i%>" value="<%=goods_name(i)%>" id="goods_name<%=i%>">
								</td>
                                <td><%=goods_standard(i)%>
                                <input type="hidden" name="goods_standard<%=i%>" value="<%=goods_standard(i)%>" id="goods_standard<%=i%>">
                                </td>
                                
                                <td><%=service_no(i)%>
                                <input type="hidden" name="service_no<%=i%>" value="<%=service_no(i)%>" id="service_no<%=i%>">
                                </td>
                                 <td><%=trade_name(i)%>
                                <input type="hidden" name="trade_name<%=i%>" value="<%=trade_name(i)%>" id="trade_name<%=i%>">
                                </td>
                                 <td><%=trade_dept(i)%>
                                <input type="hidden" name="trade_dept<%=i%>" value="<%=trade_dept(i)%>" id="trade_dept<%=i%>">
                                </td>
                                 <td class="left"><%=r_bigo(i)%>
                                <input type="hidden" name="r_bigo<%=i%>" value="<%=r_bigo(i)%>" id="r_bigo<%=i%>">
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
'						        end if
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
	      </div>
		</div>
	</body>
</html>
