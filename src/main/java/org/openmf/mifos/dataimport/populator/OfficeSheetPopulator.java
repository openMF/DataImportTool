package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Office;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class OfficeSheetPopulator extends AbstractWorkbookPopulator {
	
	private static final Logger logger = LoggerFactory.getLogger(OfficeSheetPopulator.class);
	
	private final RestClient client;
	
	private String content;
	
	private List<Office> offices;
	private ArrayList<String> officeNames;
	
	private static final int ID_COL = 0;
	private static final int OFFICE_NAME_COL = 1;

    public OfficeSheetPopulator(RestClient client) {
        this.client = client;
    }
    
    @Override
    public Result downloadAndParse() {
    	Result result = new Result();
        try {
        	client.createAuthToken();
        	offices = new ArrayList<Office>();
        	officeNames = new ArrayList<String>();
            content = client.get("offices?limit=-1");
            parseOffices();
        } catch (Exception e) {
            result.addError(e.getMessage());
            logger.error(e.getMessage());
        }
        return result;
    }

    @Override
    public Result populate(Workbook workbook) {
    	Result result = new Result();
    	try{
        int rowIndex = 1;
        Sheet officeSheet = workbook.createSheet("Offices");
        setLayout(officeSheet);
        
        populateOffices(officeSheet, rowIndex);
        officeSheet.protectSheet("");
    	} catch (Exception e) {
    		result.addError(e.getMessage());
    		logger.error(e.getMessage());
    	}
        return result;
    }
    
    private void populateOffices(Sheet officeSheet, int rowIndex) {
    	for(Office office:offices) {
        	Row row = officeSheet.createRow(rowIndex);
        	writeInt(ID_COL, row, office.getId());
        	writeString(OFFICE_NAME_COL, row, office.getName().trim().replaceAll("[ )(]", "_"));
        	rowIndex++;
        }
    }
    
    private void parseOffices() {
    	Gson gson = new Gson();
        JsonElement json = new JsonParser().parse(content);
        JsonArray array = json.getAsJsonArray();
        Iterator<JsonElement> iterator = array.iterator();
        while(iterator.hasNext()) {
        	json = iterator.next();
        	Office office = gson.fromJson(json, Office.class);
        	offices.add(office);
        	officeNames.add(office.getName().trim().replaceAll("[ )(]", "_"));
        }
    }
    
    private void setLayout(Sheet worksheet) {
    	worksheet.setColumnWidth(ID_COL, 2000);
        worksheet.setColumnWidth(OFFICE_NAME_COL, 7000);
        Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        writeString(ID_COL, rowHeader, "ID");
        writeString(OFFICE_NAME_COL, rowHeader, "Name");
    }
    
    public List<Office> getOffices() {
        return offices;
    }
    
    public String[] getOfficeNames() {
        return officeNames.toArray(new String[officeNames.size()]);
    }

}
