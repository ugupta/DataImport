package org.openmf.mifos.dataimport.web;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@WebServlet(name = "DownloadServiceServlet", urlPatterns = {"/download"})
public class DownloadServiceServlet extends HttpServlet {
	private static final long serialVersionUID = 2L;
	
	private static final Logger logger = LoggerFactory.getLogger(DownloadServiceServlet.class);
    
	@Override
	public void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
		String fileName=request.getParameter("type")+"-"+request.getParameter("documentType");
		logger.info(fileName);
		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-Disposition", "attachment;filename="+fileName);
		InputStream is = this.getClass().getResourceAsStream("/office/"+fileName);
		int read=0;
		byte[] bytes = new byte[30000];
		OutputStream os = response.getOutputStream();
		
		while((read = is.read(bytes))!= -1){
			os.write(bytes, 0, read);
		}
		os.flush();
		os.close();	
	}

}
