<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
slip_seq = request("slip_seq")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql="select * from card_slip where slip_seq = "&cint(slip_seq)
Set rs=DbConn.Execute(Sql)

title_line = "카드 전표 수정"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
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
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.account.value =="") {
					alert('계정과목을 입력하세요.');
					frm.account.focus();
					return false;}
				if(document.frm.account_item.value =="") {
					alert('항목을 입력하세요.');
					frm.account_item.focus();
					return false;}
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function cost_cal(txtObj){
				price = parseInt(document.frm.price.value.replace(/,/g,""));			
				cost_vat = parseInt(document.frm.cost_vat.value.replace(/,/g,""));			
				cost = price - cost_vat;
				cost = String(cost);
				num_len = cost.length;
				sil_len = num_len;
				cost = String(cost);
				if (cost.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) cost = cost.substr(0,num_len -3) + "," + cost.substr(num_len -3,3);
				if (sil_len > 6) cost = cost.substr(0,num_len -6) + "," + cost.substr(num_len -6,3) + "," + cost.substr(num_len -2,3);

				document.frm.cost.value = cost; 

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
	<body oncontextmenu="return false" ondragstart="return false">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="card_slip_mod_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="40%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">사용일</th>
								<td class="left"><%=rs("slip_date")%></td>
                                <th>사용자</th>
								<td class="left"><%=rs("org_name")%>/<%=rs("emp_name")%></td>
							</tr>
							<tr>
								<th class="first">거래처</th>
								<td class="left"><%=rs("customer")%><%=rs("customer_no")%></td>
								<th>부가세</th>
								<td class="left"><input name="cost_vat" type="text" id="cost_vat" style="width:100px;text-align:right" value="<%=formatnumber(rs("cost_vat"),0)%>" onKeyUp="cost_cal(this)"></td>
							</tr>
							<tr>
								<th class="first">합계</th>
								<td class="left"><input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(rs("price"),0)%>" readonly="true"></td>
								<th>공급가액</th>
								<td class="left"><input name="cost" type="text" id="cost" style="width:100px;text-align:right" value="<%=formatnumber(rs("cost"),0)%>" readonly="true"></td>
							</tr>
							<tr>
								<th class="first">계정과목</th>
								<td class="left"><select name="account" id="account" style="width:150px">
								  <%
                                    Sql="select * from account where account_code > '100' order by account_name asc"
                                    rs_acc.Open Sql, Dbconn, 1
                                    do until rs_acc.eof
								%>
								  <option value='<%=rs_acc("account_name")%>' <%If rs_acc("account_name") = rs("account") then %>selected<% end if %>><%=rs_acc("account_name")%></option>
								  <%
                                        rs_acc.movenext()
                                    loop
                                    rs_acc.close()						
                                %>
							    </select></td>
								<th>항목</th>
								<td class="left"><input name="account_item" type="text" id="account_item" style="width:150px" onKeyUp="checklength(this,50);" value="<%=rs("account_item")%>"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
				<%	if end_sw <> "Y" then	%>
                    <span class="btnType01"><input type="button" value="변경" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="slip_seq" value="<%=slip_seq%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

