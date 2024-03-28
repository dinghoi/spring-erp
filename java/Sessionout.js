function expireSession()
{
  window.location = "login.jsp";
}
setTimeout('expireSession()', <%= request.getSession().getMaxInactiveInterval() * 1000 %>);
