<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
	<!--<h1><img src="/image/logo_long_ko_main.png" alt="K-ONE" width="121" height="40"/></h1>--><!--개발서버용-->
	<h1><img src="/images/logo_long_ko_main.png" alt="K-ONE" width="121" height="40"/></h1>

	<h2>
		<div style="float:left;margin-top:5px;">
			<!--<img src="/image/logo_long_en.png" alt="K-ONE" width="103" height="24"/>--><!--개발서버용-->
			<img src="/images/logo_long_en.png" alt="K-ONE" width="103" height="24"/>

		</div>
		<div style="float:left;margin-top:4px;">
			<!--<img src="/image/nkp_title.gif" alt="Information Portal" width="219" height="23"/>--><!--개발서버용-->
			<img src="/images/nkp_title.gif" alt="Information Portal" width="219" height="23"/>

		</div>
	</h2>
	<div class="login">
		<p><strong><%=user_name%>&nbsp;<%=user_grade%>님</strong></p>
		<a href="#" onclick="javascript:pop_user_mod();"><img src="/image/user_mod.gif" alt="개인정보변경"/></a>
		<a href="/main/logout.asp"><img src="/image/logout.gif" alt="로그아웃"/></a>
	</div>
	<div id="gnb">
		<ul><li></li></ul>
	</div>
</div>
