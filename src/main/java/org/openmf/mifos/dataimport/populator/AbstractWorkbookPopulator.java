package org.openmf.mifos.dataimport.populator;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Office;
import org.openmf.mifos.dataimport.dto.client.CompactClient;
import org.openmf.mifos.dataimport.dto.client.CompactGroup;

public abstract class AbstractWorkbookPopulator implements WorkbookPopulator {

	    protected void writeInt(int colIndex, Row row, int value) {
	            row.createCell(colIndex).setCellValue(value);
	    }
	    
	    protected void writeLong(int colIndex, Row row, long value) {
            row.createCell(colIndex).setCellValue(value);
        }

	    protected void writeString(int colIndex, Row row, String value) {
	           row.createCell(colIndex).setCellValue(value);
	    }
	    
	    protected void writeDouble(int colIndex, Row row, double value) {
	    	    row.createCell(colIndex).setCellValue(value);
	    }

	    protected void writeFormula(int colIndex, Row row, String formula) {
	    	    row.createCell(colIndex).setCellType(Cell.CELL_TYPE_FORMULA);
	    	    row.createCell(colIndex).setCellFormula(formula);
	    }
	    
	    protected CellStyle getDateCellStyle(Workbook workbook) {
	    	CellStyle dateCellStyle = workbook.createCellStyle();
	        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
	        dateCellStyle.setDataFormat(df);
	        return dateCellStyle;
	    }
	    
	    protected void writeDate(int colIndex, Row row, String value, CellStyle dateCellStyle) {
	    	try {
	    		//To make validation between functions inclusive.
	    	    Date date = new SimpleDateFormat("dd/MM/yyyy", Locale.ENGLISH).parse(value);
	    		Calendar cal = Calendar.getInstance();
	    		cal.setTime(date);
	    	    cal.set(Calendar.HOUR_OF_DAY, 0);
	    	    cal.set(Calendar.MINUTE, 0);
	    	    cal.set(Calendar.SECOND, 0);
	    	    cal.set(Calendar.MILLISECOND, 0);
	    	    Date dateWithoutTime = cal.getTime();
	    	    row.createCell(colIndex).setCellValue(dateWithoutTime);
	    	    row.getCell(colIndex).setCellStyle(dateCellStyle);
	    	    } catch (ParseException pe) {
	    	    	throw new IllegalArgumentException("ParseException");
	    	    } 
	    }
	    
	    protected void setClientAndGroupDateLookupTable(Sheet sheet, List<CompactClient> clients, List<CompactGroup> groups,
	    		int nameCol, int activationDateCol) {
	    	Workbook workbook = sheet.getWorkbook();
	    	CellStyle dateCellStyle = workbook.createCellStyle();
	        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
	        dateCellStyle.setDataFormat(df);	
	    	int rowIndex = 0;
    		for(CompactClient client: clients) {
    			Row row = sheet.getRow(++rowIndex);
    			if(row == null)
    			    row = sheet.createRow(rowIndex);
    			writeString(nameCol, row, client.getDisplayName().replaceAll("[ )(] ", "_"));
    			writeDate(activationDateCol, row, client.getActivationDate().get(2) + "/" + client.getActivationDate().get(1) + "/" + client.getActivationDate().get(0), dateCellStyle);
    		}
    		
    		for(CompactGroup group: groups) {
    			Row row = sheet.getRow(++rowIndex);
    			writeString(nameCol, row, group.getName().replaceAll("[ )(] ", "_"));
    			writeDate(activationDateCol, row, group.getActivationDate().get(2) + "/" + group.getActivationDate().get(1) + "/" + group.getActivationDate().get(0), dateCellStyle);
    		}
	    }
	    
	    protected void setOfficeDateLookupTable(Sheet sheet, List<Office> offices, int officeNameCol, int activationDateCol) {
	    	Workbook workbook = sheet.getWorkbook();
	    	CellStyle dateCellStyle = workbook.createCellStyle();
	        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
	        dateCellStyle.setDataFormat(df);
	    	int rowIndex = 0;
	    	for(Office office:offices) {
	        	Row row = sheet.createRow(++rowIndex);
	        	writeString(officeNameCol, row, office.getName().trim().replaceAll("[ )(]", "_"));
	        	writeDate(activationDateCol, row, ""+office.getOpeningDate().get(2)+"/"+office.getOpeningDate().get(1)+"/"+office.getOpeningDate().get(0), dateCellStyle);
	        }
	    }
}
