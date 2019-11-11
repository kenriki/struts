package test;

import java.io.IOException;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

@WebServlet("/TestP")
public class TestP extends HttpServlet {
	private static final long serialVersionUID = 1L;

	public TestP() {
		super();
	}

	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		// 呼び出し元Jspからデータ受け取り
		request.setCharacterEncoding("UTF8");
		String jsp = request.getParameter("fromJsp");

		// 呼び出し先Jspに渡すデータセット
		request.setAttribute("fromServlet", jsp + " + サーブレットで追加");

		// resultP.jsp にページ遷移
		RequestDispatcher dispatch = request.getRequestDispatcher("resultP.jsp");
		dispatch.forward(request, response);
	}
}
