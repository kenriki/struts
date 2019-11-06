<%@ page contentType="text/html; charset=utf-8" %>
<html>
  <head>
    <title>result</title>
  </head>
  <body>
    <%= request.getAttribute( "num1" ) %> + 
    <%= request.getAttribute( "num2" ) %> = 
    <%= request.getAttribute( "resultNum" ) %>
  </body>
</html>
