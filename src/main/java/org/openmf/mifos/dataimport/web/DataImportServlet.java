package org.openmf.mifos.dataimport.web;

import java.io.DataOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;

import org.openmf.mifos.dataimport.DataImportHandler;
import org.openmf.mifos.dataimport.ImportFormatType;
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

        String filename = "";
        try {
            Part part = request.getPart("file");
            filename = readFileName(part);
            ImportFormatType type = ImportFormatType.of(part.getContentType());
            InputStream content = part.getInputStream();
            DataImportHandler handler = ImportHandlerFactory.createImportHandler(content, type);
            Result result = parseAndUpload(handler);
            writeResult(result, response.getOutputStream());
        } catch (IOException e) {
            throw new ServletException("Cannot import request. " + filename, e);
        }

    }

    String readFileName(Part part) {
        String filename = null;
        for (String s : part.getHeader("content-disposition").split(";")) {
            if (s.trim().startsWith("filename")) {
                filename = s.split("=")[1].replaceAll("\"", "");
            }
        }
        return filename;
    }

    Result parseAndUpload(DataImportHandler handler) throws IOException {
        Result result = handler.parse();
        if (result.isSuccess()) {
            result = handler.upload();
        }
        return result;
    }

    void writeResult(Result result, OutputStream stream) throws IOException {
        DataOutputStream ds = new DataOutputStream(stream);
        if(result.isSuccess()) {
            ds.writeUTF("Import complete");
            logger.debug("failed" + result);
        }
        for(String e : result.getErrors()) {
            ds.writeUTF(e);
        }
        ds.flush();
        ds.close();
    }

}
