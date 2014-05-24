package org.openmf.mifos.dataimport.web;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;

import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.handler.DataImportHandler;
import org.openmf.mifos.dataimport.handler.ImportFormatType;
import org.openmf.mifos.dataimport.handler.ImportHandlerFactory;
import org.openmf.mifos.dataimport.handler.Result;
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
            ImportFormatType.of(part.getContentType());
            InputStream content = part.getInputStream();
            Workbook workbook = new HSSFWorkbook(content);
            DataImportHandler handler = ImportHandlerFactory.createImportHandler(workbook);
            Result result = parseAndUpload(handler);
            writeResult(workbook, result, response);
        } catch (IOException e) {
            throw new ServletException("Cannot import request. " + filename, e);
        }

    }

    private String readFileName(Part part) {
        String filename = null;
        for (String s : part.getHeader("content-disposition").split(";")) {
            if (s.trim().startsWith("filename")) {
                filename = s.split("=")[1].replaceAll("\"", "");
            }
        }
        return filename;
    }

    private Result parseAndUpload(DataImportHandler handler) throws IOException {
        Result result = handler.parse();
        if (result.isSuccess()) {
            result = handler.upload();
        }
        return result;
    }

    private void writeResult(Workbook workbook, Result result, HttpServletResponse response) throws IOException {
    	OutputStream stream = response.getOutputStream();
        OutputStreamWriter out = new OutputStreamWriter(stream,"UTF-8");
        if(result.isSuccess()) {
        	out.write("<html>"
        			+ "<head>"
        			+ "<style>"
        			+ "div.step{padding:15px;margin-bottom:20px;background-image: -webkit-linear-gradient(top,#fff 0,#f8f8f8 100%);background-image: linear-gradient(to bottom,#fff 0,#f8f8f8 100%);background-repeat: repeat-x;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#ffffffff', endColorstr='#fff8f8f8', GradientType=0);filter: progid:DXImageTransform.Microsoft.gradient(enabled=false);border-radius: 4px;-webkit-box-shadow: inset 0 1px 0 rgba(255,255,255,.15),0 1px 5px rgba(0,0,0,.075);box-shadow: inset 0 1px 0 rgba(255,255,255,.15),0 1px 5px rgba(0,0,0,.075);}"
        			+ "*{margin:0;padding:0;}body{font-family:Helvetica Neue, Arial, sans-serif;background: #EEE;}"
					+ "div#title{color:#777;text-align: center;font-size: 2em;margin: 10px auto 0 auto;padding: 25px;font-weight: bold;}"
					+ "div#container{width:75%;margin:0 auto;}div#content {padding: 20px;}"
        			+ ".btn {text-decoration:none;color: #fff;background-color: #428bca;border-color: #357ebd;background-image: -webkit-linear-gradient(top,#428bca 0,#2d6ca2 100%);background-image: linear-gradient(to bottom,#428bca 0,#2d6ca2 100%);filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#ff428bca', endColorstr='#ff2d6ca2', GradientType=0);filter: progid:DXImageTransform.Microsoft.gradient(enabled=false);background-repeat: repeat-x;border-color: #2b669a;padding: 5px 10px;font-size: 12px;line-height: 1.5;border-radius: 3px;display: inline-block;margin-bottom: 0;font-weight: 400;text-align: center;vertical-align: middle;cursor: pointer;background-image: none;border: 1px solid transparent;white-space: nowrap;-webkit-user-select: none;-moz-user-select: none;-ms-user-select: none;}"
        			+ ".btn:hover{color: #fff;background-color: #3276b1;border-color: #285e8e;}"
        			+ "</style></head>"
        			+ "<body>"
        			+ "<div id='title'>Data Import Tool</div><"
        			+ "div id='container'>"
        			+ "<div id='content'>"
        			+ "<div class='step' style='border-left: 5px solid rgb(30, 224, 30);'>Import complete </div>"
        			+ "<div style='text-align:center;padding-top:25px;'><a href='/DataImportTool' class='btn'>Go Back</a></div>"
        			+ "</div>"
        			+ "</div>"
        			+ "</body>"
        			+ "</html>");
        } else {
        for(String e : result.getErrors())
            logger.debug("Failed: " + e);
        String fileName = "Re-Upload.xls";
		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-Disposition", "attachment;filename="+fileName);
        workbook.write(stream);
        }
        out.flush();
        out.close();
    }

}
