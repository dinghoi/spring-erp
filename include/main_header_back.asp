			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/></h1>
				<h2><img src="/image/main_title.gif" alt="A/S 관리 시스템" width="225" height="25"/></h2>
				<%
				mi_view = 0
				Sql = "select count(*) from sign_msg where recv_id = '"&user_id&"' and read_yn = 'N'"
				Set rs_mi = Dbconn.Execute (sql)
				mi_view = cint(rs_mi(0))
				rs_mi.close()
				%>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%>님.</strong>
					</p>
					<a href="#" onClick="pop_Window('sign_process_mg.asp','sign_process_mg_pop','scrollbars=yes,width=1250,height=600')"><img src="image/close_icon.gif" width="16" height="13"><%=mi_view%>건&nbsp;</a>
					<a href="#" onclick="javascript:pop_user_mod();"><img src="/image/user_mod.gif" alt="개인정보변경"/></a>
                    <a href="logout.asp"><img src="/image/logout.gif" alt="로그아웃"/></a>
				</div>	
                <div id="gnb">
                    <ul>
                        <li></li>    
                    </ul>
                </div>
			</div>
