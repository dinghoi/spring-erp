<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim pummok_tab(4,20)
dim cost_tab(6,40)

slip_id = request("slip_id")
slip_no = request("slip_no")
slip_seq = request("slip_seq")
cancel_yn = request("cancel_yn")

title_line = "전표 결재 요청"
if cancel_yn = "Y" then
	title_line = "전표 취소 결재 요청"
end if

Sql="select * from sales_slip where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
Set rs=DbConn.Execute(Sql)

view_att_file = rs("att_file")
if rs("slip_id") = "1" then	
	view_slip_id = "대기전표"
  else
	view_slip_id = "수주전표"
end if
if rs("sales_yn") = "Y" then	
	view_sales_yn = "매출"
  else
	view_sales_yn = "비매출"
end if
if rs("bill_issue_yn") = "Y" then	
	view_bill_issue_yn = "발행"
  else
	view_bill_issue_yn = "미발행"
end if

buy_cost = rs("buy_cost")
sales_cost = rs("sales_cost")
sales_cost_vat = rs("sales_cost_vat")
sales_price = rs("sales_price")
margin_cost = rs("margin_cost")
if rs("sales_cost") = 0 then
	margin_per = 0
  else
	margin_per = rs("margin_cost")/rs("sales_cost") * 100
end if
view_att_file = rs("att_file")
path = "/sales_file"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 전표 결재 요청</title>
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
			<h3 class="tit"><%=title_line%></h3>
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
							  <th>전표유형<br>전표번호</th>
							  <td class="left"><%=view_slip_id%>&nbsp;<%=slip_no%>-<%=slip_seq%></td>
							  <th>매출조직</th>
							  <td class="left"><%=rs("sales_company")%>&nbsp;<%=rs("sales_company")%></td>
							  <th>영업사원</th>
							  <td class="left"><%=rs("emp_name")%>&nbsp;<%=rs("org_name")%></td>
						    </tr>
							<tr>
							  <th>거래처</th>
							  <td class="left"><%=rs("trade_name")%></td>
							  <th>사업자번호</th>
							  <td class="left"><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%></td>
							  <th>거래처<br>
						      담당자</th>
							  <td class="left"><%=rs("trade_person")%>&nbsp;</td>
                          </tr>
							<tr>
							  <th>연락처</th>
							  <td class="left"><%=rs("trade_person_tel_no")%>&nbsp;</td>
							  <th>계산서 메일</th>
							  <td class="left"><%=rs("trade_email")%></td>
							  <th>매출구분</th>
							  <td class="left"><%=view_sales_yn%></td>
                          </tr>
							<tr>
							  <th>매출일자</th>
							  <td class="left"><%=rs("sales_date")%></td>
							  <th>제품출고<br>
요청일</th>
							  <td class="left"><%=rs("out_request_date")%></td>
							  <th>계산서<br>
발행여부</th>
							  <td class="left"><%=view_bill_issue_yn%></td>
						    </tr>
							<tr>
							  <th>계산서<br>
발행예정일</th>
							  <td class="left"><%=rs("bill_due_date")%></td>
							  <th>계산서발행일</th>
							  <td class="left"><%=rs("bill_issue_date")%></td>
							  <th>수금상태</th>
							  <td class="left"><%=rs("collect_stat")%></td>
						    </tr>
							<tr>
							  <th>수금방법</th>
							  <td class="left"><%=rs("bill_collect")%></td>
							  <th>수금예정일</th>
							  <td class="left"><%=rs("collect_due_date")%></td>
							  <th>수금일</th>
							  <td class="left"><%=rs("collect_date")%>&nbsp;</td>
						    </tr>
							<tr>
							  <th>비고</th>
							  <td colspan="5" class="left"><%=rs("slip_memo")%></td>
						    </tr>
						</tbody>
					</table>
				<h3 class="stit">* 품목 내역</h3>
           		<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
							<col width="12%" >
							<col width="*" >
							<col width="6%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">유형</th>
								<th scope="col">품목</th>
								<th scope="col">규격</th>
								<th scope="col">수량</th>
								<th scope="col">매입단가</th>
								<th scope="col">판매단가</th>
								<th scope="col">판매총액</th>
								<th scope="col">마진단가</th>
								<th scope="col">마진총액</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						rs.close()
						Sql="select * from sales_slip_detail where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"' order by goods_seq asc"
						Rs.Open Sql, Dbconn, 1
						do until rs.eof
							i = i + 1
						%>
			  				<tr>
								<td class="first"><%=i%></td>
								<td><%=rs("srv_type")%></td>
								<td><%=rs("pummok")%></td>
								<td><%=rs("standard")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("qty"),0)%></td>
								<td class="right"><%=formatnumber(rs("buy_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("qty")*rs("sales_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_cost")-rs("buy_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("qty")*(rs("sales_cost")-rs("buy_cost")),0)%></td>
							</tr>
						<%
							rs.movenext()
						loop
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
							<tr>
							  <th>매입총액</th>
							  <td class="right"><%=formatnumber(buy_cost,0)%></td>
							  <th>매출총액</th>
							  <td class="right"><%=formatnumber(sales_cost,0)%></td>
							  <th>매출부가세</th>
							  <td class="right"><%=formatnumber(sales_cost_vat,0)%></td>
						    </tr>
							<tr>
							  <th>총매출액</th>
							  <td class="right"><%=formatnumber(sales_price,0)%></td>
							  <th>마진총액</th>
							  <td class="right"><%=formatnumber(margin_cost,0)%></td>
							  <th>마진비율</th>
							  <td class="right"><%=formatnumber(margin_per,2)%>%</td>
                          </tr>
							<tr>
							  <th>첨부파일</th>
							  <td colspan="5" class="left">
						<% if view_att_file = "" or isnull(view_att_file) then	%>
                              &nbsp;
						<%   else	%>
							  <a href="http//intra.k-won.co.kr/download.asp?path=<%=path%>&att_file=<%=view_att_file%>"><%=view_att_file%></a>
						<% end if	%>
                              </td>
						    </tr>
						</tbody>
					</table>
					<br>
        	</div>
		</div>
    </body>
</html>

