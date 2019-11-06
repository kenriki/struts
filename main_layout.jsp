<%@ page language="java" contentType="text/html; charset=UTF-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ja" lang="ja">
<head>
<title>${param.title}</title>
</head>
<body>
<jsp:include page="header.jsp"/>
<jsp:include page="navigation.jsp"/>
<div id="main">
${param.content}
</div>
<jsp:include page="footer.jsp"/>
</body>
</html>
