<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

buy_no = request("buy_no")
buy_date = request("buy_date")
buy_seq = request("buy_seq")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

sql = "select * from met_order where (order_buy_no = '"&buy_no&"') and (order_buy_seq = '"&buy_seq&"') and (order_buy_date = '"&buy_date&"') ORDER BY order_no,order_seq,order_date ASC"
Rs.Open Sql, Dbconn, 1

'response.write(sql)

title_line = " 구매 발주 현황 "

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
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}		
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}					
			function chkfrm() {
						
				{
				a=confirm('구매품의를 취소하겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_buy_order_list.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="3%" >
                      <col width="6%" >
                      <col width="6%" >
                      <col width="6%" >
				      <col width="8%" >
                      <col width="10%" >
                      <col width="8%" >
				      <col width="6%" >
                      <col width="8%" >
                      <col width="10%" >
                      <col width="8%" >
				      <col width="6%" >
                      <col width="6%" >
                      <col width="*" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">순번</th>
                        <th scope="col">구매번호</th>
                        <th scope="col">용도구분</th>
				        <th scope="col">구매품의일</th>
                        <th scope="col">회사</th>
                        <th scope="col">부서</th>
                        <th scope="col">발주번호</th>
                        <th scope="col">발주일자</th>
                        <th scope="col">발주담당</th>
                        <th scope="col">발주거래처</th>
                        <th scope="col">발주품목</th>
                        <th scope="col">발주금액</th>
                        <th scope="col">진행상태</th>
                        <th scope="col">적요</th>
			          </tr>
			        </thead>
				    <tbody>
        <%
						seq = 0
						do until rs.eof
                           seq = seq + 1
						   order_no = rs("order_no")
						   order_seq = rs("order_seq")
						   order_date = rs("order_date")
						   
						   buy_ing = rs("order_ing")
						   buy_ing_gubun = ""
						   if buy_ing = "0" then
						         buy_ing_gubun = "구매품의"
						      elseif buy_ing = "1" then
							            buy_ing_gubun = "부분발주"
									 elseif buy_ing = "2" then
							                   buy_ing_gubun = "전체발주"
										    elseif buy_ing = "3" then
							                          buy_ing_gubun = "발주완료"
												   elseif buy_ing = "4" then
							                                 buy_ing_gubun = "입고"
						   end if
						   
						   sql = "select * from met_order_goods where (og_order_no = '"&order_no&"') and (og_order_seq = '"&order_seq&"') and (og_order_date = '"&order_date&"')  ORDER BY og_seq,og_goods_code ASC"
						   Set Rs_buy=DbConn.Execute(Sql)
						   if Rs_buy.eof or Rs_buy.bof then
								bg_goods_name = ""
							  else
							  	bg_goods_name = Rs_buy("og_goods_name")
						   end if
						   Rs_buy.close()
		%>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td>
                        <a href="#" onClick="pop_Window('met_buy_detail_list.asp?buy_no=<%=buy_no%>&buy_date=<%=buy_date%>&buy_seq=<%=buy_seq%>&u_type=<%=""%>','met_buy_detail_pop','scrollbars=yes,width=930,height=650')"><%=buy_no%></a>
                        </td>
                        <td><%=rs("order_goods_type")%>&nbsp;</td>
                        <td><%=rs("order_buy_date")%>&nbsp;</td>
                        <td><%=rs("order_company")%>&nbsp;</td>
                        <td><%=rs("order_org_name")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('met_buy_order_detail.asp?order_no=<%=rs("order_no")%>&order_date=<%=rs("order_date")%>&order_seq=<%=rs("order_seq")%>&u_type=<%=""%>','met_buy_order_detail_pop','scrollbars=yes,width=930,height=650')"><%=rs("order_no")%>&nbsp;<%=rs("order_seq")%></a>
                        </td>
                        <td><%=rs("order_date")%>&nbsp;</td>
                        <td><%=rs("order_emp_name")%>&nbsp;</td>
                        <td><%=rs("order_trade_name")%>&nbsp;</td>
                        <td><%=bg_goods_name%>&nbsp;외</td>
                        <td class="right"><%=formatnumber(rs("order_cost"),0)%></td>
                        <td><%=buy_ing_gubun%>&nbsp;</td>
                        <td class="left"><%=rs("order_memo")%>&nbsp;</td>
			          </tr>
				    <%
							rs.movenext()
						loop
						rs.close()
					%>
			        </tbody>
			      </table>
          	     <br>
     				<div class="noprint">
                        <div align=center>
                            <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                        </div>
					</div>
					<br>               		
                    <input type="hidden" name="user_id">
		            <input type="hidden" name="pass">
                    
                    <input type="hidden" name="order_no" value="<%=order_no%>">
					<input type="hidden" name="order_seq" value="<%=order_seq%>">
					<input type="hidden" name="order_date" value="<%=order_date%>">
	     </form>
		</div>				
	</div>        				
	</body>
</html>

