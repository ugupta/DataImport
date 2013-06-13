package org.openmf.mifos.dataimport.web;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.openmf.mifos.dataimport.DataImportHandler;
import org.openmf.mifos.dataimport.ImportHandlerFactory;
import org.openmf.mifos.dataimport.Result;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@WebServlet(name = "DataImportServlet", urlPatterns = {"/import"})
@MultipartConfig(maxFileSize=10000000, fileSizeThreshold=10000000)
public class DataImportServlet extends HttpServlet {

    private static final long serialVersionUID = 1L;

    private static final Logger logger = LoggerFactory.getLogger(DataImportServlet.class);

    @Override
    public void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        try {
            List<FileItem> items = new ServletFileUpload(new DiskFileItemFactory()).parseRequest(request);
            for (FileItem item : items) {
                if (!item.isFormField()) {
                   // String fieldname = item.getFieldName();
                   // String filename = FilenameUtils.getName(item.getName());
                    if (!item.getContentType().equals("application/vnd.ms-excel")) { throw new FileUploadException("Only excel files accepted! provided : " +item.getContentType() ); }
                    InputStream content = item.getInputStream();
                    DataImportHandler handler = ImportHandlerFactory.createImportHandler(content);
                    parse(response, handler);
                }
            }
        } catch (FileUploadException e) {
            throw new ServletException("Cannot upload request.", e);
        }

    }

    private void parse(HttpServletResponse response, DataImportHandler handler) throws IOException {
        Result parseResult = handler.parse();
        if (parseResult.isSuccess()) {
            upload(response, handler);
            response.getWriter().write("Success!");
            return;
        }
        writeErrors(parseResult, response);
    }

    private void upload(HttpServletResponse response, DataImportHandler handler) throws IOException {
        Result uploadResult = handler.upload();
        if (uploadResult.isSuccess()) {
            response.setStatus(HttpServletResponse.SC_OK);
            response.getWriter().write("Complete!");
            logger.info("Upload successful!");
            return;
        }
        writeErrors(uploadResult, response);
    }
    

    private void writeErrors(Result parseResult, HttpServletResponse response) throws IOException {
        for(String e : parseResult.getErrors()) {
            response.getWriter().println(e);
        }
    }

}
