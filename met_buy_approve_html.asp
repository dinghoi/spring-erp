<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

buy_no = request("buy_no")
buy_date = request("buy_date")
buy_seq = request("buy_seq")

title_line = "구매품의 결재 요청"

sql = "select * from met_buy where (buy_no = '"&buy_no&"') and (buy_date = '"&buy_date&"') and (buy_seq = '"&buy_seq&"')"
Set rs=DbConn.Execute(Sql)
if not rs.eof then
    	buy_no = rs("buy_no")
		buy_date = rs("buy_date")
		buy_goods_type = rs("buy_goods_type")
		buy_company = rs("buy_company")
	    buy_bonbu = rs("buy_bonbu")
		buy_saupbu = rs("buy_saupbu")
		buy_team = rs("buy_team")
	    buy_org_code = rs("buy_org_code")
	    buy_org_name = rs("buy_org_name")
	    buy_emp_no = rs("buy_emp_no")
	    buy_emp_name = rs("buy_emp_name")
	    buy_bill_collect = rs("buy_bill_collect")
        buy_collect_due_date = rs("buy_collect_due_date")
	    buy_trade_no = rs("buy_trade_no")
        buy_trade_name = rs("buy_trade_name")
        buy_trade_person = rs("buy_trade_person")
		buy_trade_email = rs("buy_trade_email")
        buy_out_method = rs("buy_out_method")
        buy_out_request_date = rs("buy_out_request_date")
        buy_price = rs("buy_price")
        buy_cost = rs("buy_cost")
        buy_cost_vat = rs("buy_cost_vat")
        buy_memo = rs("buy_memo")
        if buy_memo = "" or isnull(buy_memo) then
	           buy_memo = rs("buy_memo")
           else
	           buy_memo = replace(buy_memo,chr(10),"<br>")
        end if
        buy_ing = rs("buy_ing")
		buy_sign_yn = rs("buy_sign_yn")
	    buy_sign_no = rs("buy_sign_no")
	    buy_sign_date = rs("buy_sign_date")
		buy_att_file = rs("buy_att_file")

	    if buy_out_request_date = "0000-00-00" then
	          buy_out_request_date = ""
	    end if
   else
		buy_company = ""
	    buy_bonbu = ""
		buy_saupbu = ""
		buy_team = ""
	    buy_org_code = ""
	    buy_org_name = ""
	    buy_emp_no = ""
	    buy_emp_name = ""
	    buy_bill_collect = ""
        buy_collect_due_date = ""
	    buy_trade_no = ""
        buy_trade_name = ""
        buy_trade_person = ""
		buy_trade_email = ""
        buy_out_method = ""
        buy_out_request_date = ""
        buy_price = 0
        buy_cost = 0
        buy_cost_vat = 0
        buy_memo = ""
        buy_ing = ""
		buy_att_file = ""
end if
rs.close()

sql = "select * from met_buy_goods where (bg_no = '"&buy_no&"') and (buy_seq = '"&buy_seq&"') and (bg_date = '"&buy_date&"') ORDER BY bg_seq,bg_goods_code ASC"

Rs.Open Sql, Dbconn, 1

view_att_file = buy_att_file
path = "/met_upload"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>구매품의 결재 요청</title>
		<style type="text/css">
		body{font-family:'Dotum','','Gulim','',sans-serif;font-size: 12px;}

		#wrap{ width: 920px; margin:0 auto; }
		#container{margin:20px 0;}
		h3.otit{color:#02880a; padding-left:12px; margin:20px 0 15px; font-size:15px;}
		.stit{margin:15px 0 10px 0; display:block; color:#000000; padding-left:15px; font-size:12px;}
		.step{font-size:16px; padding-left:12px; color:red;}
		.tit{font-size:16px; padding-left:12px; color:#02880a;}
		.insa{font-size:16px; padding-left:12px; color:#00008c;}
		.brown{font-size:16px; padding-left:12px; color:#7D0000}
		.btit{font-size:24px; padding-left:12px; color:#000000;}
		h3.teof{color:#515254; padding-left:12px; margin:20px 0 15px; font-size:14px;}
		/*button*/
		.btnType01{display:inline-block; min-width:75px;height:28px;  text-align:center;border:1px solid #dedee0; border-bottom:1px solid #acafb6; text-decoration:none !important; font-weight:bold; font-size:12px; color:#666 !important;overflow:hidden;
		background:#f0f0f2 -webkit-linear-gradient(top, #ffffff, #f0f0f2);
		background:#f0f0f2 -moz-linear-gradient(top, #ffffff, #f0f0f2);
		background:#f0f0f2 -o-linear-gradient(top, #ffffff, #f0f0f2);
		filter:progid:DXImageTransform.Microsoft.gradient(startColorStr=#ffffff, endColorStr=#f0f0f2)}
		.btnType01 input,.btnType01 button{display:inline-block;margin-top:-2px; padding:0 10px; background:none; color:#666 !important; border:0; cursor:pointer;outline:none !important; font-weight:bold}
		.btnType01 input{margin:0; height:30px}
		a.btnType01{padding:0 10px;min-width:55px; line-height:28px;margin:0; }
		a.btnType01 img{position:absolute;top:0;left:0}
		/* board view */
		.tableView{clear:both; width:100%;border-top:2px solid #515254; border-bottom:1px solid #cbcbcb;table-layout:fixed;word-Break:break-all}
		.tableView img{vertical-align:middle}
		.tableView thead th{padding-top:5px;height:25px;line-height:1.1em; text-align:center; border-left:1px solid #e3e3e3; background-color:#f8f8f8; color:#515254;}
		.tableView tbody th{padding-top:5px;height:25px;line-height:1.1em; border-left:1px solid #e3e3e3; border-top:1px solid #e3e3e3; background-color:#f8f8f8;color:#515254;}
		.tableView td{padding:8px 0 4px;border-top:1px solid #e3e3e3; border-left:1px solid #e3e3e3; text-align:center;}
		.tableView th:first-child,
		.tableView th.first{border-left:none;}
		.tableView textarea{width:98%; font-size:12px; padding:5px; resize:none;}
		.tableView .inputTxta{padding:5px 0; border-left:none; text-align:center; text-align:left; padding:15px}
		
		.tableView .inputTxta textarea{height:60px;}
		
		.tableView .left {padding-left:10px; text-align:left}
		.tableView .right {padding-right:10px; text-align:right}
		
		.tableView dl{overflow:hidden; padding-top:5px;}
		.tableView dt{float:left;clear:left; padding-top:2px; margin:0 10px 0 0; font-weight:bold; font-weight:normal;}
		.tableView dd{float:left; margin-top:3px;}
		.tableView dd *{vertical-align:middle;}
		.tableView input{text-align:center;}

		/* board list */
		.tableList{clear:both; width:100%;border-top:2px solid #515254; border-bottom:1px solid #cbcbcb;table-layout:fixed;word-Break:break-all}
		.tableList img{vertical-align:middle}
		.tableList thead th{padding-top:5px;height:28px;line-height:1.1em; text-align:center; border-left:1px solid #e3e3e3; background-color:#f8f8f8}
		.tableList tbody th{padding-top:5px;height:28px;line-height:1.1em; text-align:center; border-left:1px solid #e3e3e3; border-top:1px solid #e3e3e3; background-color:#FFECFF}
		.tableList tfoot th{padding-top:5px;height:28px;line-height:1.1em; text-align:center; border-left:1px solid #e3e3e3; background-color:#FFECFF; color:#515254;}
		.tableList td{padding:8px 0 4px;border-top:1px solid #e3e3e3; text-align:center; border-left:1px solid #e3e3e3}
		.tableList th:first-child,
		.tableList th.first{border-left:none;}
		.tableList td:first-child,
		.tableList td.first{border-left:none;}
		.tableList .left {padding-left:10px; text-align:left}
		.tableList .right {padding-right:5px; text-align:right}
		.tableList .left a:hover{text-decoration:underline;}
		.noData{padding:10px 0 6px;text-align:center;border-top:2px solid #515254; border-bottom:1px solid #cbcbcb;}

		</style>
	</head>
	<body>
		<div id="wrap">			
			<h3 class="insa"><%=title_line%></h3>
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
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
                                <th>구매번호</th>
							    <td class="left"><%=buy_no%>&nbsp;<%=buy_seq%></td>
							    <th>구매유형</th>
							    <td class="left"><%=buy_goods_type%></td>
							    <th>구매일자</th>
							    <td class="left"><%=buy_date%></td>
 							</tr>
                            <tr>
							    <th>구매회사</th>
							    <td class="left"><%=buy_company%></td>
							    <th>사업부</th>
							    <td class="left"><%=buy_saupbu%></td>
							    <th>구매담당</th>
							    <td class="left"><%=buy_org_name%>&nbsp;<%=buy_emp_name%></td>
						    </tr>
							<tr>
                                <th>구매처</th>
							    <td class="left"><%=buy_trade_name%></td>
							    <th>사업자번호</th>
							    <td class="left"><%=buy_trade_no%></td>
							    <th>담당자</th>
							    <td class="left"><%=buy_trade_person%></td>
						    </tr>
                            <tr>
                                <th>이메일</th>
							    <td class="left"><%=buy_trade_email%></td>
							    <th>대금<br>지급방법</th>
							    <td class="left"><%=buy_bill_collect%></td>
							    <th>지급예정일</th>
							    <td class="left"><%=buy_collect_due_date%></td>
						    </tr>
							<tr>
							  <th>비고</th>
							  <td colspan="5" class="left"><%=buy_memo%></td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 구매 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="16%" >
							<col width="14%" >
							<col width="8%" >
							<col width="12%" >
							<col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
								<th scope="col">수량</th>
								<th scope="col">구입단가</th>
								<th scope="col">구입금액</th>
							</tr>
						</thead>
						<tbody>     
						<%
							buy_cost_tot = 0
							i = 0
							do until rs.eof or rs.bof
							     i = i + 1
							
							     buy_hap = rs("bg_qty") * rs("bg_unit_cost")
							     buy_cost_tot = buy_cost_tot + buy_hap
							
						%>
							<tr>
								<td class="first"><%=i%></td>
                                <td><%=rs("bg_goods_type")%>&nbsp;</td>
								<td><%=rs("bg_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("bg_goods_code")%>&nbsp;</td>
                                <td><%=rs("bg_goods_name")%>&nbsp;</td>
                                <td><%=rs("bg_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("bg_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("bg_unit_cost"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(buy_hap,0)%>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
						</tbody>
					</table>
                    <br>
                    <table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="*" >
						</colgroup>
						<tbody>
                        <% 
						    buy_vat_hap = int(buy_cost_tot * (10 / 100))
							buy_tot_price = buy_cost_tot + buy_vat_hap
						%>
							<tr>
							  <th>구매총액</th>
							  <td class="right"><%=formatnumber(buy_tot_price,0)%></td>
							  <th>구매금액</th>
							  <td class="right"><%=formatnumber(buy_cost_tot,0)%></td>
							  <th>부가세</th>
							  <td class="right"><%=formatnumber(buy_vat_hap,0)%></td>
						    </tr>
							<tr>
							  <th>첨부</th>
							  <td colspan="5" class="left">
                        <% 
                           If buy_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=buy_att_file%>"><%=buy_att_file%></a>
                        <%    Else %>
				                    &nbsp;
                        <% 
						   End If %>
                              </td>
						    </tr>
						</tbody>
					</table>
					<br>
        	</div>
		</div>
    </body>
</html>

