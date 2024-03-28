<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
company = request.form("company")
gubun = request.form("gubun")
code_seq = request.form("code_seq")
asset_name = request.form("asset_name")
asset_name = Replace(asset_name,"'","&quot;")
asset_name = Replace(asset_name,"""","&quot;")
asset_no = request.form("asset_no")
asset_cnt = int(request.form("asset_cnt"))
buy_date = request.form("buy_date")
reg_id = user_id

no_yymm = mid(asset_no,3,6)
in_year = mid(buy_date,1,4)
in_month = mid(buy_date,6,2)
in_yymm = in_year + in_month
set dbconn = server.CreateObject("adodb.connection")
dbconn.open dbconnect

if in_yymm <> no_yymm then
	sql="select max(asset_no) as max_no from asset where mid(asset_no,1,2) = '" + company + "' and mid(asset_no,3,6) = '" + in_yymm + "'"
	set rs_no=dbconn.execute(sql)
	if isnull(rs_no("max_no")) then
		asset_no = company + in_yymm + "0000"
		no_cnt = 0
	  else  	
		asset_no = rs_no("max_no")
	end if
end if
no_cnt = int(right(asset_no,4))

i = 0
do until i = asset_cnt
	i = i + 1	
	no_cnt = no_cnt + 1
	as_no = "000" + cstr(no_cnt)
	new_asset_no = mid(asset_no,1,8) + right(as_no,4)
	
	sql="insert into asset (asset_no,company,gubun,code_seq,asset_name,sticker_yn,buy_date,inst_process,reg_id,reg_date) values ('"&new_asset_no&"','"&company&"','"&gubun&"','"&code_seq&"','"&asset_name&"','N','"&buy_date&"','N','"&reg_id&"',now())"
	dbconn.execute(sql)
loop

sql="insert into asset_buy (buy_date,company,gubun,code_seq,asset_name,buy_cnt,reg_id,reg_date) values ('"&buy_date&"','"&company&"','"&gubun&"','"&code_seq&"','"&asset_name&"','"&asset_cnt&"','"&reg_id&"',now())"
dbconn.execute(sql)

end_msg = cstr(asset_cnt) +" 건 자산번호를 부여하였습니다...."
response.write"<script language=javascript>"
response.write"alert('"&end_msg&"');"
'response.write"location.replace('k1_asset_send_mg.asp');"
response.write"self.opener.location.reload('asset_send_mg.asp');"		
response.write"window.close();"		
response.write"</script>"
Response.End

dbconn.Close()
Set dbconn = Nothing

%>
