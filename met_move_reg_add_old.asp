<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim pummok_tab(13,20)
dim amount_tab(1,20)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")

u_type = request("u_type")

view_condi=Request("view_condi")
goods_type=Request("goods_type")
rele_id=Request("rele_id")

rele_date = ""
rele_stock = ""
rele_stock_company = ""
rele_stock_name = ""
chulgo_type = "출고요청"
chulgo_date = ""
chulgo_stock = ""
chulgo_stock_name = ""
chulgo_stock_company = ""

pummok_cnt = 0

path_name = "/met_upload"

for i = 1 to 13
	for j = 1 to 20
		pummok_tab(i,j) = ""
	next
next
for i = 1 to 1
	for j = 1 to 20
		amount_tab(i,j) = 0
	next
next
' response.write(reg_date)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_bg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

rele_emp_no = "100482"
rele_emp_name = "오세삼"
'rele_emp_no = user_id
'rele_emp_name = user_name

Sql = "SELECT * FROM emp_master where emp_no = '"&rele_emp_no&"'"
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

rele_saupbu = emp_saupbu
rele_stock_name = emp_org_name

title_line = goods_type + " 창고이동 출고의뢰 등록 "

if u_type = "U" then

	Sql="select * from met_mv_reg where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')  and (rele_id = '"&rele_id&"')"
	Set rs=DbConn.Execute(Sql)

	rele_date = rs("rele_date")
    rele_stock = rs("rele_stock")
	rele_seq = rs("rele_seq")
    rele_stock_company = rs("rele_stock_company")
    rele_stock_name = rs("rele_stock_name")
    rele_emp_no = rs("rele_emp_no")
    rele_emp_name = rs("rele_emp_name")
    rele_company = rs("rele_company")
    rele_bonbu = rs("rele_bonbu")
    rele_saupbu = rs("rele_saupbu")
    rele_team = rs("rele_team")
    rele_org_name = rs("rele_org_name")
    rele_trade_name = rs("rele_trade_name")
    service_no = rs("service_no")
    chulgo_type = rs("chulgo_type")
    chulgo_date = rs("chulgo_date")
    chulgo_stock = rs("chulgo_stock")
    chulgo_stock_name = rs("chulgo_stock_name")
	chulgo_stock_company = rs("chulgo_stock_company")
	rele_addr = rs("rele_addr")
	rele_memo = rs("rele_memo")
	in_stock_date = rs("in_stock_date")
	
	reg_att_file = rs("reg_att_file")

	if chulgo_date = "1900-01-01" then
	      chulgo_date = ""
	end if
	
	rs.close()
	
	j = 0
	Sql="select * from met_mv_reg_goods where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_stock_seq = '"&rele_seq&"')"
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		j = j + 1
		pummok_tab(1,j) = rs("bg_goods_type")
		pummok_tab(2,j) = rs("bg_goods_gubun")
		pummok_tab(3,j) = rs("bg_goods_code")
		pummok_tab(4,j) = rs("bg_goods_name")
		pummok_tab(5,j) = rs("bg_standard")
		pummok_tab(6,j) = rs("bg_seq")
		pummok_tab(7,j) = rs("rele_reside_compamy")
		pummok_tab(8,j) = rs("rele_custom_part")
		pummok_tab(9,j) = rs("rele_ce_no")
		pummok_tab(10,j) = rs("rele_ce_name")
		pummok_tab(11,j) = rs("rele_acpt_no")
		pummok_tab(12,j) = rs("rele_acpt_date")
		pummok_tab(13,j) = rs("rele_issue")
		amount_tab(1,j) = rs("rele_qty")
		rs.movenext()
	loop
	pummok_cnt = j
    
	title_line = goods_type + " 창고이동 출고의뢰 변경 "
	
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
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=rele_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=chulgo_rele_date%>" );
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
				if (chkfrm()) {
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
				if(document.frm.rele_trade_name.value == "") {
					alert('고객사를 선택하세요');
					frm.rele_trade_name.focus();
					return false;}
				if(document.frm.chulgo_stock_name.value == "") {
					alert('출고창고를 선택하세요');
					frm.chulgo_stock_name.focus();
					return false;}
				if(document.frm.chulgo_date.value == "") {
					alert('출고요청일를 입력하세요');
					frm.chulgo_date.focus();
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
				window.open('met_goods_select.asp?code_ary='+code_ary+'', '팝업공지', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
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
			var qty_ary = new Array();

			for (j=1;j<21;j++) {
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
		
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
                <form method="post" name="frm" action="met_move_reg_add_save.asp">
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
							  <th>신청창고(부서)</th>
							  <td colspan="3" class="left">
                              <input name="rele_saupbu" type="text" value="<%=rele_saupbu%>" readonly="true" style="width:120px">
                              <input name="rele_stock_name" type="text" value="<%=rele_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="mvreg"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="rele_stock" value="<%=chulgo_stock%>" ID="Hidden1">
                              <input type="hidden" name="rele_bonbu" value="<%=rele_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_company" value="<%=rele_stock_company%>" ID="Hidden1">
                              <input type="hidden" name="rele_team" value="<%=rele_team%>" ID="Hidden1">
                              </td>
							  <th>신청담당자</th>
							  <td class="left">
							  <input name="rele_emp_name" type="text" value="<%=rele_emp_name%>" readonly="true" style="width:70px">
                              <input name="rele_emp_no" type="text" value="<%=rele_emp_no%>" readonly="true" style="width:40px">
							  </td>
						    </tr>
							<tr>
							  <th>신청일자</th>
							  <td class="left"><input name="rele_date" type="text" value="<%=rele_date%>" style="width:80px;text-align:center" id="datepicker"></td>
                              <th>출고요청일</th>
							  <td colspan="3" class="left"><input name="chulgo_rele_date" type="text" value="<%=chulgo_rele_date%>" style="width:80px;text-align:center" id="datepicker1"></td>
						    </tr>
                            <tr>
                              <th>받을주소</th>
							  <td colspan="5" class="left"><input name="rele_addr" type="text" id="rele_addr" style="width:700px; ime-mode:active"  value="<%=rele_addr%>"></td>
                            </tr>  
							<tr>
							  <th>출고처창고</th>
							  <td colspan="3" class="left">
                              <input name="chulgo_stock_company" type="text" value="<%=chulgo_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="chulgo_stock_name" type="text" value="<%=chulgo_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="mvin"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>" ID="Hidden1">
                              <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="Hidden1">
                              </td>
                              <th>출고처 창고장</th>
							  <td class="left">
							  <input name="stock_manager_name" type="text" value="<%=stock_manager_name%>" readonly="true" style="width:70px">
                              <input name="stock_manager_code" type="text" value="<%=stock_manager_code%>" readonly="true" style="width:40px">
							  </td>
                            </tr>
                            <tr>
							  <th>비고</th>
							  <td colspan="5" class="left"><input name="rele_memo" type="text" value="<%=rele_memo%>" style="width:700px; ime-mode:active"></td>
						    </tr>
						</tbody>
					</table>
				</div>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 창고이동 출고의뢰 품목 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="12%" >
                            <col width="*" >
                            <col width="16%" >
							<col width="16%" >
							<col width="16%" >
							<col width="14%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first; left" colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;"><a href="#" onClick="pop_pummok_del()" class="btnType03">선택삭제</a>&nbsp;<a href="#" onClick="pop_pummok()" class="btnType03">품목선택</a></th>
							</tr>
							<tr>
								<th class="first" scope="col"><input type="checkbox" name="tot_check" id="tot_check"></th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
								<th scope="col">수량</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><input type="checkbox" name="del_check<%=i%>" id="del_check<%=i%>" value="Y"></td>
								<td>
                                <input name="srv_type<%=i%>" type="text" id="srv_type<%=i%>" style="width:70px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_gubun<%=i%>" type="text" id="goods_gubun<%=i%>" style="width:120px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_code<%=i%>" type="text" id="goods_code<%=i%>" style="width:90px" readonly="true">
                                </td>
								<td><input name="goods_name<%=i%>" type="text" id="goods_name<%=i%>" style="width:120px" readonly="true"></td>
								<td><input name="goods_standard<%=i%>" type="text" id="goods_standard<%=i%>" style="width:120px"></td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(qty,0)%>" onKeyUp="NumCal(this);"></td>
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
							  <th>첨부</th>
                              <td colspan="5" class="left">
                              <a href="download.asp?path=<%=path_name%>&att_file=<%=reg_att_file%>"><%=reg_att_file%></a>
                              <input name="att_file" type="file" id="att_file" size="100">
                              </td>
						    </tr>
						</tbody>
					</table>                                       
					<br>
				</div>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="u_goods_type" value="<%=goods_type%>" ID="Hidden1">
                <input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="emp_name" value="<%=emp_name%>" ID="Hidden1">
                <input type="hidden" name="emp_company" value="<%=emp_company%>" ID="Hidden1">
                <input type="hidden" name="emp_bonbu" value="<%=emp_bonbu%>" ID="Hidden1">
                <input type="hidden" name="emp_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="emp_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="emp_org_code" value="<%=emp_org_code%>" ID="Hidden1">
                <input type="hidden" name="emp_org_name" value="<%=emp_org_name%>" ID="Hidden1">
                
                <input type="hidden" name="old_att_file" value="<%=reg_att_file%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
				</form>
		</div>				
	</body>
</html>
