<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

rele_no = request("rele_no")
rele_seq = request("rele_seq")
rele_date = request("rele_date")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

sql = "select * from met_chulgo where (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"') and (rele_date = '"&rele_date&"') and (chulgo_id = '본사출고') ORDER BY chulgo_date,chulgo_stock,chulgo_seq ASC"
Rs.Open Sql, Dbconn, 1

'response.write(sql)

title_line = " 출고의뢰 실출고 현황 "

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
				a=confirm('출고으뢰를 취소하겠습니까?')
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
				<form action="met_chulgo_reg_list.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="3%" >
                      <col width="8%" >
                      <col width="8%" >
				      <col width="8%" >
                      <col width="8%" >
				      <col width="12%" >

				      <col width="12%" >
				      <col width="12%" >
				      <col width="8%" >
                      <col width="8%" >
                      <col width="*" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">순번</th>
                        <th scope="col">의뢰일자</th>
                        <th scope="col">의뢰번호</th>
                        <th scope="col">용도구분</th>
                        <th scope="col">신청자</th>
                        <th scope="col">의뢰소속</th>

                        <th scope="col">출고품목</th>
                        <th scope="col">출고창고</th>
                        <th scope="col">출고일자</th>
                        <th scope="col">출고상태</th>
                        <th scope="col">적요</th>
			          </tr>
			        </thead>
				    <tbody>
        <%
						seq = 0
						do until rs.eof
                           seq = seq + 1
						   chulgo_date = rs("chulgo_date")
						   chulgo_stock = rs("chulgo_stock")
						   chulgo_seq = rs("chulgo_seq")
					       
						   sql = "select * from met_chulgo_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')  ORDER BY cg_goods_seq,cg_goods_code ASC"
						   Set Rs_good=DbConn.Execute(Sql)
						   if Rs_good.eof or Rs_good.bof then
								bg_goods_name = ""
							  else
							  	bg_goods_name = Rs_good("cg_goods_name")
						   end if
						   Rs_good.close()
						   
						   sql = "select * from met_chulgo_reg where (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"') and (rele_date = '"&rele_date&"')"
						   Set Rs_reg=DbConn.Execute(Sql)
						   if Rs_reg.eof or Rs_reg.bof then
								rele_emp_name = ""
								rele_org_name = ""
								chulgo_ing = ""
							  else
							  	rele_emp_name = Rs_reg("rele_emp_name")
								rele_org_name = Rs_reg("rele_org_name")
								chulgo_ing = Rs_reg("chulgo_ing")
						   end if
						   Rs_reg.close()
		%>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td><%=rs("rele_date")%>&nbsp;</td>
                        <td>
						<a href="#" onClick="pop_Window('met_chulgo_reg_detail.asp?rele_no=<%=rs("rele_no")%>&rele_date=<%=rs("rele_date")%>&rele_seq=<%=rs("rele_seq")%>&u_type=<%=""%>','met_chulgo_reg_detail_pop','scrollbars=yes,width=1000,height=650')"><%=rs("rele_no")%>&nbsp;<%=rs("rele_seq")%></a>
                        </td>
						<td><%=rs("chulgo_goods_type")%>&nbsp;</td>
                        <td><%=rele_emp_name%>&nbsp;</td>
                        <td><%=rele_org_name%>&nbsp;</td>

                        <td><%=bg_goods_name%>&nbsp;외</td>
                        <td><%=rs("chulgo_stock_name")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('met_chulgo_cust_detail.asp?chulgo_date=<%=rs("chulgo_date")%>&chulgo_stock=<%=rs("chulgo_stock")%>&chulgo_seq=<%=rs("chulgo_seq")%>&u_type=<%=""%>','met_chulgo_detail_pop','scrollbars=yes,width=1000,height=650')"><%=rs("chulgo_date")%></a>
                        </td>
                        <td><%=chulgo_ing%>&nbsp;</td>
                        <td><%=rs("chulgo_memo")%>&nbsp;</td>
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

