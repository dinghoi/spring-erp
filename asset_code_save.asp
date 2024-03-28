<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	on Error resume next

	dquot = "'"
	u_type = request.form("u_type")
	company = request.form("company")
	gubun = request.form("gubun")
	code_seq = cstr(request.form("code_seq"))
	asset_name = request.form("asset_name")
	asset_name = Replace(asset_name,"'","&quot;")
	asset_name = Replace(asset_name,"""","&quot;")
	maker = request.form("maker")
	cpu = request.form("cpu")
	mem = request.form("mem")
	hdd = request.form("hdd")
	os = request.form("os")
	spec = request.form("spec")
	spec = Replace(spec,"'","&quot;")
	spec = Replace(spec,"""","&quot;")
	rental = request.form("rental")
	unit_price = int(request.form("unit_price"))

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if u_type = "U" then
		sql = "update asset_code set asset_name='"+asset_name+"', maker='"+maker+"', cpu='"+cpu+"', mem='"+mem+"', hdd='"+hdd+"', os='"+os+"', spec='"+spec+"', rental='"+rental+"', unit_price="&unit_price&", mod_id='"+user_id+"', mod_date=now() where company='"+company+"' and gubun = '"+gubun+"' and code_seq='"+code_seq+"'"
	  else
		sql="select max(code_seq) as max_seq from asset_code where company='" + company + "' and gubun = '" + gubun + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			code_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			code_seq = right(max_seq,3)
		end if

		sql="insert into asset_code (company,gubun,code_seq,asset_name,maker,cpu,mem,hdd,os,spec,rental,unit_price,reg_id,reg_date) values ('"+company+"','"+gubun+"','"+code_seq+"','"+asset_name+"','"+maker+"','"+cpu+"','"+mem+"','"+hdd+"','"+os+"','"+spec+"','"+rental+"','"&unit_price&"','"+user_id+"',now())"
	end if
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
