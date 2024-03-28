<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
stock_goods_code = request("stock_goods_code")
stock_goods_type = request("stock_goods_type")
stock_code = request("stock_code")
stock_name = request("stock_name")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_jae = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

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

if u_type = "U" then

	sql = "select * from met_stock_gmaster where (stock_code = '"&stock_code&"') and (stock_goods_type = '"&stock_goods_type&"') and (stock_goods_code = '"&stock_goods_code&"')"
	set rs = dbconn.execute(sql)

    stock_goods_type = rs("stock_goods_type")
	stock_last_qty = rs("stock_last_qty")
	stock_last_amt = rs("stock_last_amt")
	stock_in_qty = rs("stock_in_qty")
	stock_in_amt = rs("stock_in_amt")
	stock_go_qty = rs("stock_go_qty")
	stock_go_amt = rs("stock_go_amt")
	stock_JJ_qty = rs("stock_JJ_qty")
	stock_jj_amt = rs("stock_jj_amt")
	
	rs.close()

	title_line = "전기이월 수량/금액 변경"
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=stin_in_date%>" );
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
//				if(document.frm.car_no.value =="" ) {
//					alert('차량번호를 입력하세요');
//					frm.car_no.focus();
//					return false;}
			
				{
				a=confirm('변경하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function num_chk(txtObj){
				last_qty = eval("document.frm.stock_last_qty.value").replace(/,/g,"");		
				last_amt = eval("document.frm.stock_last_amt.value").replace(/,/g,"");		
				in_qty = eval("document.frm.stock_in_qty.value").replace(/,/g,"");		
				go_qty = eval("document.frm.stock_go_qty.value").replace(/,/g,"");		
				
				qty_cal = parseInt(last_qty) + parseInt(in_qty) - parseInt(go_qty);
				
				qty_cal = String(qty_cal);
				num_len = qty_cal.length;
				sil_len = num_len;
				if (qty_cal.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) qty_cal = qty_cal.substr(0,num_len -3) + "," + qty_cal.substr(num_len -3,3);
				if (sil_len > 6) qty_cal = qty_cal.substr(0,num_len -6) + "," + qty_cal.substr(num_len -6,3) + "," + qty_cal.substr(num_len -2,3);
				if (sil_len > 9) qty_cal = qty_cal.substr(0,num_len -9) + "," + qty_cal.substr(num_len -9,3) + "," + qty_cal.substr(num_len -5,3) + "," + qty_cal.substr(num_len -1,3);
				
				eval("document.frm.stock_jj_qty.value = qty_cal");

				if (txtObj.value.length >= 2) {
					if (txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
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
				<form action="met_stock_last_mod_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
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
							    <th>창고</th>
							    <td class="left" colspan="3"><%=stock_name%>&nbsp;(<%=stock_code%>)</td>
                                <th>품목구분</th>
							    <td class="left"><%=goods_gubun%>&nbsp;</td>
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
							    <th>규격</th>
							    <td class="left"><%=goods_standard%>&nbsp;</td>
                                <th>모델</th>
							    <td class="left"><%=goods_model%>&nbsp;</td>
                                <th>Serial No.</th>
							    <td class="left"><%=goods_serial_no%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 입출고 수량 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="15%" >
                            <col width="*" >
                            <col width="15%" >
                            <col width="15%" >
                            <col width="15%" >
						</colgroup>
						<thead>
							<tr>
                                <th scope="col">용도구분</th>
                                <th scope="col">전기이월 수량</th>
                                <th scope="col">전기이월 금액</th>
                                <th scope="col">입고수량</th>
                                <th scope="col">출고수량</th>
                                <th scope="col">현재고</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td><%=stock_goods_type%></td>
								<td>
                                <input name="stock_last_qty" type="text" id="stock_last_qty" value="<%=formatnumber(stock_last_qty,0)%>"  style="width:80px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <td>
                                <input name="stock_last_amt" type="text" id="stock_last_amt" value="<%=formatnumber(stock_last_amt,0)%>"  style="width:120px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <td>
                                <input name="stock_in_qty" type="text" disabled id="stock_in_qty" value="<%=formatnumber(stock_in_qty,0)%>"  style="width:80px;text-align:right" readonly="true">
                                </td>
                                <td>
                                <input name="stock_go_qty" type="text" disabled id="stock_go_qty" value="<%=formatnumber(stock_go_qty,0)%>"  style="width:80px;text-align:right" readonly="true">
                                </td>
                                <td>
                                <input name="stock_jj_qty" type="text" disabled id="stock_jj_qty" value="<%=formatnumber(stock_JJ_qty,0)%>"  style="width:80px;text-align:right" readonly="true">
                                </td>
							</tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                
                <input type="hidden" name="stock_goods_type" value="<%=stock_goods_type%>">
                <input type="hidden" name="stock_code" value="<%=stock_code%>">
                <input type="hidden" name="stock_goods_code" value="<%=stock_goods_code%>">
			</form>
		</div>				
	</body>
</html>

