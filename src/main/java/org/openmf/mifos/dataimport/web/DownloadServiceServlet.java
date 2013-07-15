package org.openmf.mifos.dataimport.web;

import java.io.IOException;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.populator.WorkbookPopulator;
import org.openmf.mifos.dataimport.populator.WorkbookPopulatorFactory;

@WebServlet(name = "DownloadServiceServlet", urlPatterns = {"/download"})
public class DownloadServiceServlet extends HttpServlet {
	
	private static final long serialVersionUID = 2L;
	
	private Workbook workbook;
    
	@Override
	public void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
		String fileName=request.getParameter("template");
		try{
		WorkbookPopulator populator = WorkbookPopulatorFactory.createWorkbookPopulator(fileName);
        downloadAndPopulate(populator);
		fileName=fileName+".xls";
		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-Disposition", "attachment;filename="+fileName);
		writeToStream(response.getOutputStream());
		
		}catch(Exception e){
			throw new ServletException("Cannot download template. " + fileName, e);
		}
	}
	
	void downloadAndPopulate(WorkbookPopulator populator) throws IOException {
        populator.downloadAndParse();
        workbook = new HSSFWorkbook();
        populator.populate(workbook);
    }
	
	 void writeToStream(OutputStream stream) throws IOException {
	             workbook.write(stream);
	        stream.flush();
	        stream.close();
	    }


}
