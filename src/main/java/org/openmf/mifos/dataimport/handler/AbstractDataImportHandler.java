package org.openmf.mifos.dataimport.handler;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public abstract class AbstractDataImportHandler implements DataImportHandler {

    protected Integer getNumberOfRows(Sheet sheet, int primaryColumn) {
        Integer noOfEntries = 1;
        // getLastRowNum and getPhysicalNumberOfRows showing false values
        // sometimes
           while (sheet.getRow(noOfEntries) !=null && sheet.getRow(noOfEntries).getCell(primaryColumn) != null) {
               noOfEntries++;
           }
        	
        return noOfEntries;
    }
    
    protected boolean isNotImported(Row row, int statusColumn) {
		return !readAsString(statusColumn, row).equals("Imported");
	}

    protected String readAsInt(int colIndex, Row row) {
        try {
        	Cell c = row.getCell(colIndex);
        	if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
        		return "";
        	return ((Double) c.getNumericCellValue()).intValue() + "";
        } catch (RuntimeException re) {
            return row.getCell(colIndex).getStringCellValue();
        }
    }
    
    protected String readAsLong(int colIndex, Row row) {
        try {
        	Cell c = row.getCell(colIndex);
        	if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
        		return "";
        	return ((Double) c.getNumericCellValue()).longValue() + "";
        } catch (RuntimeException re) {
            return row.getCell(colIndex).getStringCellValue();
        }
    }
    
    protected Double readAsDouble(int colIndex, Row row) {
    	Cell c = row.getCell(colIndex);
    	if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
    		return 0.0;
    	return row.getCell(colIndex).getNumericCellValue();
    }

    protected String readAsString(int colIndex, Row row) {
        try {
        	Cell c = row.getCell(colIndex);
        	if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
        		return "";
            return c.getStringCellValue().trim();
        } catch (Exception e) {
            return ((Double)row.getCell(colIndex).getNumericCellValue()).intValue() + "";
        }
    }

    protected String readAsDate(int colIndex, Row row) {
    	try{
    		Cell c = row.getCell(colIndex);
    		if(c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
    			return "";
    		DateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy");
            return dateFormat.format(c.getDateCellValue());
    	}  catch  (Exception e) {
    		return "";
    	}
    }
    
    protected String readAsDateWithoutYear(int colIndex, Row row) {
    	try{
    		Cell c = row.getCell(colIndex);
    		if(c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
    			return "";
    		DateFormat dateFormat = new SimpleDateFormat("dd MMMM");
            return dateFormat.format(c.getDateCellValue());
    	}  catch  (Exception e) {
    		return "";
    	}
    }
    
    protected Boolean readAsBoolean(int colIndex, Row row) {
    	try{
    	    Cell c = row.getCell(colIndex);
		    if(c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
			    return false;
    	    return c.getBooleanCellValue();
    	} catch (Exception e) {
    		String booleanString = row.getCell(colIndex).getStringCellValue().trim();
    		if(booleanString.equalsIgnoreCase("TRUE"))
    			return true;
    		else
    			return false;
    	}
    }
    
    protected void writeString(int colIndex, Row row, String value) {
        row.createCell(colIndex).setCellValue(value);
    }
    
    protected CellStyle getCellStyle(Workbook workbook, IndexedColors color) {
    	CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        return style;
    }
    
    protected String parseStatus(String errorMessage) {
    	StringBuffer message = new StringBuffer();
    	JsonObject obj = new JsonParser().parse(errorMessage.trim()).getAsJsonObject();
        JsonArray array = obj.getAsJsonArray("errors");
        Iterator<JsonElement> iterator = array.iterator();
        while(iterator.hasNext()) {
        	JsonObject json = iterator.next().getAsJsonObject();
        	String parameterName = json.get("parameterName").toString();
        	String defaultUserMessage = json.get("defaultUserMessage").toString();
        	message = message.append(parameterName.substring(1, parameterName.length() - 1) + ":" + defaultUserMessage.substring(1, defaultUserMessage.length() - 1) + "\t");
         }
    	 return message.toString();
    }
    
    protected Integer getIdByName (Sheet sheet, String name) {
    	String sheetName = sheet.getSheetName();
    	if(!sheetName.equals("Products")) {
    	for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getRichStringCellValue().getString().trim().equals(name)) {
                    	if(sheetName.equals("Offices"))
                            return ((Double)row.getCell(cell.getColumnIndex() - 1).getNumericCellValue()).intValue(); 
                    	else if(sheetName.equals("Extras"))
                    		return ((Double)row.getCell(cell.getColumnIndex() - 1).getNumericCellValue()).intValue();
                    	else if(sheetName.equals("GlAccounts"))
                    		return ((Double)row.getCell(cell.getColumnIndex() - 1).getNumericCellValue()).intValue();
                    	else if(sheetName.equals("Clients") || sheetName.equals("Center")|| sheetName.equals("Groups") || sheetName.equals("Staff") ) 
                    		/* Monkey patch - if office and center names are the same it errors out from looking
                    		 for a number in a string. You have to check if the cell is a number, if not get the
                    		  next column which will work when office and branch share names on group import.
							*/
							if (row.getCell(cell.getColumnIndex() + 1).getCellType() == Cell.CELL_TYPE_NUMERIC){
								return ((Double)row.getCell(cell.getColumnIndex() + 1).getNumericCellValue()).intValue();
							}
							else {
								return ((Double)row.getCell(cell.getColumnIndex() + 2).getNumericCellValue()).intValue();
							}
                    }
            }
          }
    	} else if (sheetName.equals("Products")) {
    		for(Row row : sheet) {
    			for(int i = 0; i < 2; i++) {
    				Cell cell = row.getCell(i);
    				if(cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getRichStringCellValue().getString().trim().equals(name)) {
    						return ((Double)row.getCell(cell.getColumnIndex() - 1).getNumericCellValue()).intValue();
    				}
    			}
    		}
    	}
        return 0;
    }

	protected String getCodeByName(Sheet sheet, String name) {
		String sheetName = sheet.getSheetName();
		sheetName.equals("Extras");
		{
			for (Row row : sheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == Cell.CELL_TYPE_STRING
							&& cell.getRichStringCellValue().getString().trim()
									.equals(name)) {
							return row.getCell(cell.getColumnIndex() - 1)
									.getStringCellValue().toString();

					}
				}
			}
		}
		return "";
	}
}