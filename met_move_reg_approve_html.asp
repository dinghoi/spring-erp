<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

rele_date = request("rele_date")
rele_stock = request("rele_stock")
rele_seq = request("rele_seq")

title_line = "창고이동 출고의뢰 결재 요청"

sql = "select * from met_mv_reg where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')"
Set rs=DbConn.Execute(Sql)
if not rs.eof then
    	rele_stock = rs("rele_stock")
        rele_seq = rs("rele_seq")
	    rele_date = rs("rele_date")
        rele_id = rs("rele_id")
        rele_goods_type = rs("rele_goods_type")
		rele_stock_company = rs("rele_stock_company")
        rele_stock_name = rs("rele_stock_name")
        rele_emp_no = rs("rele_emp_no")
        rele_emp_name = rs("rele_emp_name")
        rele_company = rs("rele_company")
        rele_bonbu = rs("rele_bonbu")
        rele_saupbu = rs("rele_saupbu")
        rele_team = rs("rele_team")
        rele_org_name = rs("rele_org_name")

        chulgo_rele_date = rs("chulgo_rele_date")
		chulgo_ing = rs("chulgo_ing")
        chulgo_date = rs("chulgo_date")
        chulgo_stock = rs("chulgo_stock")
        chulgo_stock_name = rs("chulgo_stock_name")
	    chulgo_stock_company = rs("chulgo_stock_company")
	    rele_att_file = rs("rele_att_file")
	    rele_memo = rs("rele_memo")
        rele_sign_yn = rs("rele_sign_yn")
	    rele_sign_no = rs("rele_sign_no")
	    rele_sign_date = rs("rele_sign_date")
	    if chulgo_date = "0000-00-00" then
	          chulgo_date = ""
	    end if
   else
		rele_stock = ""
        rele_seq = ""
	    rele_date = ""
        rele_id = ""
        rele_goods_type = ""
        rele_stock_company = ""
        rele_stock_name = ""
		rele_emp_no = ""
        rele_emp_name = ""
        rele_company = ""
        rele_bonbu = ""
        rele_saupbu = ""
        rele_team = ""
        rele_org_name = ""

        chulgo_rele_date = ""
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
rs.close()

sql = "select * from met_mv_reg_goods where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')  ORDER BY rl_goods_seq,rl_goods_code ASC"

Rs.Open Sql, Dbconn, 1

view_att_file = buy_att_file
path = "/met_upload"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>창고이동 출고의뢰 결재 요청</title>
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
                                <th>신청일자</th>
							    <td class="left"><%=rele_date%></td>
							    <th>용도구분</th>
							    <td class="left"><%=rele_goods_type%></td>
							    <th>신청창고</th>
							    <td class="left"><%=rele_stock_name%>&nbsp;(<%=rele_stock_company%>)</td>
 							</tr>
                            <tr>
							    <th>회사</th>
							    <td class="left"><%=rele_company%></td>
							    <th>사업부</th>
							    <td class="left"><%=rele_saupbu%></td>
							    <th>신청자</th>
							    <td class="left"><%=rele_emp_name%>&nbsp;(<%=rele_org_name%>)</td>
						    </tr>
							<tr>
                                <th>출고요청일</th>
							    <td class="left"><%=chulgo_rele_date%></td>
							    <th>출고처창고</th>
							    <td colspan="3" class="left"><%=chulgo_stock_name%>&nbsp;(<%=chulgo_stock_company%>)</td>
						    </tr>
                            <tr>
                                <th>실출고일</th>
							    <td class="left"><%=chulgo_date%>&nbsp;</td>
							    <th>신청창고<br>입고일</th>
							    <td colspan="3" class="left"><%=in_stock_date%>&nbsp;</td>
						    </tr>
							<tr>
							  <th>비고</th>
							  <td colspan="5" class="left"><%=rele_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 창고이동 출고의뢰 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="8%" >
                            <col width="*" >
                            <col width="12%" >
							<col width="18%" >
							<col width="18%" >
							<col width="8%" >
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
							</tr>
						</thead>
						<tbody>     
						<%
							i = 0
							do until rs.eof or rs.bof
							     i = i + 1
							
						%>
							<tr>
								<td class="first"><%=i%></td>
                                <td><%=rs("rl_goods_type")%>&nbsp;</td>
                                <td><%=rs("rl_goods_grade")%>&nbsp;</td>
								<td><%=rs("rl_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("rl_goods_code")%>&nbsp;</td>
                                <td><%=rs("rl_goods_name")%>&nbsp;</td>
                                <td><%=rs("rl_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rl_qty"),0)%>&nbsp;</td>
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
							<tr>
							  <th>첨부</th>
							  <td colspan="5" class="left">
                        <% 
                           If rele_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=rele_att_file%>"><%=rele_att_file%></a>
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

