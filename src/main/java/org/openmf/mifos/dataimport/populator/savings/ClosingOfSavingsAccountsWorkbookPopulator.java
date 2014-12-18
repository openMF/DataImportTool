package org.openmf.mifos.dataimport.populator.savings;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.openmf.mifos.dataimport.dto.savings.CompactSavingsAccount;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.openmf.mifos.dataimport.populator.AbstractWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.ClientSheetPopulator;
import org.openmf.mifos.dataimport.populator.OfficeSheetPopulator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class ClosingOfSavingsAccountsWorkbookPopulator extends AbstractWorkbookPopulator  {
	
	
	 private static final Logger logger = LoggerFactory.getLogger(ClosingOfSavingsAccountsWorkbookPopulator.class);
		
		private final RestClient restClient;
		
		private String content;
		
		private OfficeSheetPopulator officeSheetPopulator;
		private ClientSheetPopulator clientSheetPopulator;
		private List<CompactSavingsAccount> savings;
		
		
		private static final int OFFICE_NAME_COL = 0;
	    private static final int CLIENT_NAME_COL = 1;
	    private static final int ACCOUNT_TYPE = 2;
	    private static final int SAVINGS_ACCOUNT_NO_COL = 3;
	    private static final int PRODUCT_COL = 4;
	    private static final int OPENING_BALANCE_COL = 5;
	    private static final int CLOSED_ON_DATE = 6;
	    private static final int ON_ACCOUNT_CLOSURE_ID = 7;
	    private static final int TO_SAVINGS_ACCOUNT_ID = 8;
	    private static final int LOOKUP_CLIENT_NAME_COL = 15;
	    private static final int LOOKUP_ACCOUNT_NO_COL = 16;
	    private static final int LOOKUP_PRODUCT_COL = 17;
	    private static final int LOOKUP_OPENING_BALANCE_COL = 18;
	    private static final int LOOKUP_SAVINGS_ACTIVATION_DATE_COL = 19;
		
	    public ClosingOfSavingsAccountsWorkbookPopulator(RestClient restClient, OfficeSheetPopulator officeSheetPopulator,
    		ClientSheetPopulator clientSheetPopulator){
	        this.restClient = restClient;
	        this.officeSheetPopulator = officeSheetPopulator;
	        this.clientSheetPopulator = clientSheetPopulator;
	    	savings = new ArrayList<CompactSavingsAccount>();
		
	    }
	    @Override  
	    public Result downloadAndParse() {
			Result result =  officeSheetPopulator.downloadAndParse();
			if(result.isSuccess())
				result = clientSheetPopulator.downloadAndParse();
			if(result.isSuccess()) {
				try {
		        	restClient.createAuthToken();
		            content = restClient.get("savingsaccounts?limit=-1");
		            Gson gson = new Gson();
		            JsonParser parser = new JsonParser();
		            JsonObject obj = parser.parse(content).getAsJsonObject();
		            JsonArray array = obj.getAsJsonArray("pageItems");
		            Iterator<JsonElement> iterator = array.iterator();
		            while(iterator.hasNext()) {
		            	JsonElement json = iterator.next();
		            	CompactSavingsAccount savingsAccount = gson.fromJson(json, CompactSavingsAccount.class);
		            	if(savingsAccount.isActive())
		            	  savings.add(savingsAccount);
		            } 
		       } catch (Exception e) {
		           result.addError(e.getMessage());
		           logger.error(e.getMessage());
		       }
			}
	    	return result;
	    }
	    
	    @Override
	    public Result populate(Workbook workbook) {
	    	Sheet closingOfSavingsAccountSheet = workbook.createSheet("ClosingOfSavingsAccounts");
	    	setLayout(closingOfSavingsAccountSheet);
	    	Result result = officeSheetPopulator.populate(workbook);
	    	if(result.isSuccess())
	    		result = clientSheetPopulator.populate(workbook);
	    	
	    	if(result.isSuccess())
	    		result = populateSavingsTable(closingOfSavingsAccountSheet);
	        if(result.isSuccess())
	            result = setRules(closingOfSavingsAccountSheet);
	        setDefaults(closingOfSavingsAccountSheet);
	        return result;
	    }
		
		
		
			
		
		private Result populateSavingsTable(Sheet closingOfSavingsAccountSheet) {
		    	Result result = new Result();
		    	Workbook workbook = closingOfSavingsAccountSheet.getWorkbook();
		    	CellStyle dateCellStyle = workbook.createCellStyle();
		        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
		        dateCellStyle.setDataFormat(df);
				int rowIndex = 1;
		    	Row row;
		    	Collections.sort(savings, CompactSavingsAccount.ClientNameComparator);
		    	try{
		    		for(CompactSavingsAccount savingsAccount : savings) {
		    			row = closingOfSavingsAccountSheet.createRow(rowIndex++);
		    			writeString(LOOKUP_CLIENT_NAME_COL, row, savingsAccount.getClientName()  + "(" + savingsAccount.getClientId() + ")");
		    			writeLong(LOOKUP_ACCOUNT_NO_COL, row, Long.parseLong(savingsAccount.getAccountNo()));
		    			writeString(LOOKUP_PRODUCT_COL, row, savingsAccount.getSavingsProductName());
		    			if(savingsAccount.getMinRequiredOpeningBalance() != null)
		    			   writeDouble(LOOKUP_OPENING_BALANCE_COL, row, savingsAccount.getMinRequiredOpeningBalance());
		    			writeDate(LOOKUP_SAVINGS_ACTIVATION_DATE_COL, row, savingsAccount.getTimeline().getActivatedOnDate().get(2) + "/" + savingsAccount.getTimeline().getActivatedOnDate().get(1) + "/" + savingsAccount.getTimeline().getActivatedOnDate().get(0), dateCellStyle);
		    		}
			    } catch (Exception e) {
				result.addError(e.getMessage());
				logger.error(e.getMessage());
			    }
		    	return result;
		    	}
		
		 private void setLayout(Sheet worksheet) {
		    	Row rowHeader = worksheet.createRow(0);
		        rowHeader.setHeight((short)500);
		        worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
		        worksheet.setColumnWidth(CLIENT_NAME_COL, 5000);
		        worksheet.setColumnWidth(ACCOUNT_TYPE, 5000);
		        worksheet.setColumnWidth(SAVINGS_ACCOUNT_NO_COL, 3000);
		        worksheet.setColumnWidth(PRODUCT_COL, 4000);
		        worksheet.setColumnWidth(OPENING_BALANCE_COL, 4000);
		        worksheet.setColumnWidth(CLOSED_ON_DATE, 4000);
		        worksheet.setColumnWidth(ON_ACCOUNT_CLOSURE_ID, 4000);
		        worksheet.setColumnWidth(TO_SAVINGS_ACCOUNT_ID, 4000);
		        writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
		        writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
		        writeString(ACCOUNT_TYPE, rowHeader, "Account Type");
		        writeString(SAVINGS_ACCOUNT_NO_COL, rowHeader, "Account No.*");
		        writeString(PRODUCT_COL, rowHeader, "Product Name");
		        writeString(OPENING_BALANCE_COL, rowHeader, "Opening Balance");
		        writeString(CLOSED_ON_DATE, rowHeader, "Clising Date ");
		        writeString(ON_ACCOUNT_CLOSURE_ID,rowHeader,"Action(Account Transfer(200) or cash(100) ");
		        writeString(TO_SAVINGS_ACCOUNT_ID,rowHeader, "Transfered Account No.");
		        writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Lookup Client");
		        writeString(LOOKUP_ACCOUNT_NO_COL, rowHeader, "Lookup Account");
		        writeString(LOOKUP_PRODUCT_COL, rowHeader, "Lookup Product");
		        writeString(LOOKUP_OPENING_BALANCE_COL, rowHeader, "Lookup Opening Balance");
		        writeString(LOOKUP_SAVINGS_ACTIVATION_DATE_COL, rowHeader, "Lookup Savings Activation Date");
}
		  private Result setRules(Sheet worksheet) {
		    	Result result = new Result();
		    	try {
		    		CellRangeAddressList officeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), OFFICE_NAME_COL, OFFICE_NAME_COL);
		        	CellRangeAddressList clientNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), CLIENT_NAME_COL, CLIENT_NAME_COL);
		        	CellRangeAddressList accountNumberRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), SAVINGS_ACCOUNT_NO_COL, SAVINGS_ACCOUNT_NO_COL);
		        	CellRangeAddressList accountTypeRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), ACCOUNT_TYPE, ACCOUNT_TYPE);
		        	
		        	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
		        	
		        	setNames(worksheet);
		        	
		        	DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
		        	DataValidationConstraint clientNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Client_\",$A1))");
		        	DataValidationConstraint accountNumberConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Account_\",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($B1,\" \",\"_\"),\"(\",\"_\"),\")\",\"_\")))");
		        	DataValidationConstraint accountTypeConstraint = validationHelper.createExplicitListConstraint(new String[] {"savingsaccounts","fixeddepositaccounts","recurringdepositaccounts"});
		        	
		        	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
		        	DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
		        	DataValidation accountNumberValidation = validationHelper.createValidation(accountNumberConstraint, accountNumberRange);
		        	DataValidation accountTypeValidation = validationHelper.createValidation(accountTypeConstraint, accountTypeRange);
		        	
		        	
		        	worksheet.addValidationData(officeValidation);
		            worksheet.addValidationData(clientValidation);
		            worksheet.addValidationData(accountNumberValidation);
		            worksheet.addValidationData(accountTypeValidation);
		           
		        	
		    	} catch (RuntimeException re) {
		    		result.addError(re.getMessage());
		    	}
		       return result;
		    }
		   private void setDefaults(Sheet worksheet) {
		    	try {
		    		for(Integer rowNo = 1; rowNo < 3000; rowNo++)
		    		{
		    			Row row = worksheet.getRow(rowNo);
		    			if(row == null)
		    				row = worksheet.createRow(rowNo);
		    			writeFormula(PRODUCT_COL, row, "IF(ISERROR(VLOOKUP($D"+ (rowNo+1) +",$Q$2:$S$" + (savings.size() + 1) + ",2,FALSE)),\"\",VLOOKUP($D"+ (rowNo+1) +",$Q$2:$S$" + (savings.size() + 1) + ",2,FALSE))");
		    			writeFormula(OPENING_BALANCE_COL, row, "IF(ISERROR(VLOOKUP($D"+ (rowNo+1) +",$Q$2:$S$" + (savings.size() + 1) + ",3,FALSE)),\"\",VLOOKUP($D"+ (rowNo+1) +",$Q$2:$S$" + (savings.size() + 1) + ",3,FALSE))");
		    		}
		    	} catch (Exception e) {
		    		logger.error(e.getMessage());
		    	}
		    }
		    
		    private void setNames(Sheet worksheet) {
		    	Workbook closingOfSavingsAccountWorkbook = worksheet.getWorkbook();
		    	ArrayList<String> officeNames = new ArrayList<String>(Arrays.asList(officeSheetPopulator.getOfficeNames()));
		    	
		    	//Office Names
		    	Name officeGroup = closingOfSavingsAccountWorkbook.createName();
		    	officeGroup.setNameName("Office");
		    	officeGroup.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));
		    	
		    	//Clients Named after Offices
		    	for(Integer i = 0; i < officeNames.size(); i++) {
		    		Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator.getOfficeNameToBeginEndIndexesOfClients().get(i);
		    		Name name = closingOfSavingsAccountWorkbook.createName();
		    		if(officeNameToBeginEndIndexesOfClients != null) {
		    	       name.setNameName("Client_" + officeNames.get(i));
		    	       name.setRefersToFormula("Clients!$B$" + officeNameToBeginEndIndexesOfClients[0] + ":$B$" + officeNameToBeginEndIndexesOfClients[1]);
		    		}
		    	}
		    	
		    	//Counting clients with active savings and starting and end addresses of cells for naming
		    	HashMap<String, Integer[]> clientNameToBeginEndIndexes = new HashMap<String, Integer[]>();
		    	ArrayList<String> clientsWithActiveSavings = new ArrayList<String>();
		    	ArrayList<String> clientIdsWithActiveSavings = new ArrayList<String>();
		    	int startIndex = 1, endIndex = 1;
		    	String clientName = "";
		    	String clientId = "";
		    	for(int i = 0; i < savings.size(); i++){
		    		if(!clientName.equals(savings.get(i).getClientName())) {
		    			endIndex = i + 1;
		    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
		    			startIndex = i + 2;
		    			clientName = savings.get(i).getClientName();
		    			clientId = savings.get(i).getClientId();
		    			clientsWithActiveSavings.add(clientName);
		    			clientIdsWithActiveSavings.add(clientId);
		    		}
		    		if(i == savings.size()-1) {
		    			endIndex = i + 2;
		    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
		    		}
		    	}
		    	
		    	//Account Number Named  after Clients
		    	for(int j = 0; j < clientsWithActiveSavings.size(); j++) {
		    		Name name = closingOfSavingsAccountWorkbook.createName();
		    		name.setNameName("Account_" + clientsWithActiveSavings.get(j).replaceAll(" ", "_") + "_" + clientIdsWithActiveSavings.get(j) + "_");
		    		
		    	
		    	}
		    	
		    	
		    
}}