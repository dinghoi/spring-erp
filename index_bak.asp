<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<title>NKP 시스템</title>
<link href="css/login.css" rel="stylesheet" type="text/css">
</head>

<body topmargin="0" leftmargin="0">
<div id="div_mian"></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="middle"><div id="div_aline">
      <div id="div_logo"><img src="image/nkp_title.gif" alt="(주)케이원정보통신" width="320" height="24" class="fr" /> <img src="images/logo.gif" alt="(주)케이원정보통신" width="120" height="56" class="mt10" /></div>
      <div id="login_fild">
        <div id="login_bg">
          <div id="login_img"><img src="images/img1.png" width="210" height="217" alt=""/></div>
          <div id="login_text"><img src="images/text1.gif" width="160" height="150" alt=""/></div>
          <div id="div_login">
            <div id="title_login"></div>
            <div id="form_login">
              <form action="login.asp" method="post">
                <label for="textfield"></label>
                <input name="id" type="text" id="id" class="id" tabindex="1" onfocus="this.className='id2'"/>
                <div class="login">
                  <input name="login" type="image" tabindex="3" src="images/loginpage_bt.gif" width="57" height="55" alt="로그인" />
                </div>
                <input name="pass" type="password" id="pw" class="pw" tabindex="2" onfocus="this.className='pw2'"/>
              </form>
            </div>
          </div>
          <div id="text_login"><img src="images/login_copyright.gif" width="390" height="40" alt=""/></div>
        </div>
      </div>
      <div id="copyright"><img src="images/copyright.gif" alt="" width="640" height="70" class="fr" /></div>
    </div></td>
  </tr>
</table>
<!--#include virtual="/include/google_analytics.asp" -->
</body>
</html>
