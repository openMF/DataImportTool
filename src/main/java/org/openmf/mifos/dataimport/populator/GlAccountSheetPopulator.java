package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.accounting.GlAccount;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class GlAccountSheetPopulator extends AbstractWorkbookPopulator {

	private static final Logger logger = LoggerFactory
			.getLogger(GlAccountSheetPopulator.class);

	private final RestClient client;
	private static final int ID_COL = 0;
	private static final int ACCOUNT_NAME_COL = 1;

	private List<GlAccount> glAccounts;
	private ArrayList<String> glAccountNames;

	public GlAccountSheetPopulator(RestClient client) {
		this.client = client;
	}

	private String content;

	@Override
	public Result downloadAndParse() {
		Result result = new Result();
        try {
        	client.createAuthToken();
        	glAccounts = new ArrayList<GlAccount>();
        	glAccountNames = new ArrayList<String>();
            content = client.get("glaccounts");
            parseglAccounts();
        } catch (Exception e) {
            result.addError(e.getMessage());
            logger.error(e.getMessage());
        }
		return result;
	}

	private void parseglAccounts() {
		Gson gson = new Gson();
        JsonElement json = new JsonParser().parse(content);
        JsonArray array = json.getAsJsonArray();
        Iterator<JsonElement> iterator = array.iterator();
        while(iterator.hasNext()) {
        	json = iterator.next();
        	GlAccount glAccount = gson.fromJson(json, GlAccount.class);
        	if(glAccount.getUsage().getValue().equals("DETAIL"))
        		glAccounts.add(glAccount);
        	glAccountNames.add(glAccount.getName().trim().replaceAll("[ )(]", "_"));
        }
	}

	@Override
	public Result populate(Workbook workbook) {
		Result result = new Result();
    	try{
        int rowIndex = 1;
        Sheet glAccountSheet = workbook.createSheet("GlAccounts");
        setLayout(glAccountSheet);
        
        populateglAccounts(glAccountSheet, rowIndex);
        glAccountSheet.protectSheet("");
    	} catch (Exception e) {
    		result.addError(e.getMessage());
    		logger.error(e.getMessage());
    	}
        return result;
	}

	private void populateglAccounts(Sheet GlAccountSheet, int rowIndex) {
		for(GlAccount glAccount:glAccounts) {
        	Row row = GlAccountSheet.createRow(rowIndex);
        	writeInt(ID_COL, row, glAccount.getId());
        	writeString(ACCOUNT_NAME_COL, row, glAccount.getName().trim().replaceAll("[ )(]", "_"));
        	rowIndex++;
        }
	}

	private void setLayout(Sheet worksheet) {
		worksheet.setColumnWidth(ID_COL, 2000);
        worksheet.setColumnWidth(ACCOUNT_NAME_COL, 7000);
        Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        writeString(ID_COL, rowHeader, "Gl Account ID");
        writeString(ACCOUNT_NAME_COL, rowHeader, "Gl Account Name");
    }
	
	public Integer getGlAccountNamesSize(){
		return glAccountNames.size();
	}
	}


