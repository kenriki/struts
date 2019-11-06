import java.io.*;
import javax.servlet.*;
import javax.servlet.http.*;

public class Calculator extends HttpServlet {

  public void doPost(HttpServletRequest request, HttpServletResponse response)
    throws IOException, ServletException{

    float num1 = 0;
    float num2 = 0;
    float resultNum;

    try {
      num1 = Float.parseFloat(request.getParameter("num1"));
      num2 = Float.parseFloat(request.getParameter("num2"));
      resultNum = num1 + num2;
    } catch (NumberFormatException e) {
      resultNum = 0;
    }

    request.setAttribute("num1", num1);
    request.setAttribute("num2", num2);
    request.setAttribute("resultNum", resultNum);

    getServletConfig().getServletContext().
      getRequestDispatcher("/jsp/result.jsp").forward(request, response);
  }
}
数字
