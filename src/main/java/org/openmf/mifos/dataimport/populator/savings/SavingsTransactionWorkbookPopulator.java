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
import org.openmf.mifos.dataimport.populator.ExtrasSheetPopulator;
import org.openmf.mifos.dataimport.populator.OfficeSheetPopulator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class SavingsTransactionWorkbookPopulator extends AbstractWorkbookPopulator {
	
    private static final Logger logger = LoggerFactory.getLogger(SavingsTransactionWorkbookPopulator.class);
	
	private final RestClient restClient;
	
	private String content;
	
	private OfficeSheetPopulator officeSheetPopulator;
	private ClientSheetPopulator clientSheetPopulator;
	private ExtrasSheetPopulator extrasSheetPopulator;
	private List<CompactSavingsAccount> savings;
	
	private static final int OFFICE_NAME_COL = 0;
    private static final int CLIENT_NAME_COL = 1;
    private static final int SAVINGS_ACCOUNT_NO_COL = 2;
    private static final int PRODUCT_COL = 3;
    private static final int OPENING_BALANCE_COL = 4;
    private static final int TRANSACTION_TYPE_COL = 5;
    private static final int AMOUNT_COL = 6;
    private static final int TRANSACTION_DATE_COL = 7;
    private static final int PAYMENT_TYPE_COL = 8;
    private static final int ACCOUNT_NO_COL = 9;
    private static final int CHECK_NO_COL = 10;
    private static final int ROUTING_CODE_COL = 11;	
    private static final int RECEIPT_NO_COL = 12;
    private static final int BANK_NO_COL = 13;
    private static final int LOOKUP_CLIENT_NAME_COL = 15;
    private static final int LOOKUP_ACCOUNT_NO_COL = 16;
    private static final int LOOKUP_PRODUCT_COL = 17;
    private static final int LOOKUP_OPENING_BALANCE_COL = 18;
    private static final int LOOKUP_SAVINGS_ACTIVATION_DATE_COL = 19;
    
    public SavingsTransactionWorkbookPopulator(RestClient restClient, OfficeSheetPopulator officeSheetPopulator,
    		ClientSheetPopulator clientSheetPopulator, ExtrasSheetPopulator extrasSheetPopulator) {
        this.restClient = restClient;
        this.officeSheetPopulator = officeSheetPopulator;
        this.clientSheetPopulator = clientSheetPopulator;
        this.extrasSheetPopulator = extrasSheetPopulator;
		savings = new ArrayList<CompactSavingsAccount>();
    }
    
    @Override
    public Result downloadAndParse() {
		Result result =  officeSheetPopulator.downloadAndParse();
		if(result.isSuccess())
			result = clientSheetPopulator.downloadAndParse();
		if(result.isSuccess())
    		result = extrasSheetPopulator.downloadAndParse();
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
    	Sheet savingsTransactionSheet = workbook.createSheet("SavingsTransaction");
    	setLayout(savingsTransactionSheet);
    	Result result = officeSheetPopulator.populate(workbook);
    	if(result.isSuccess())
    		result = clientSheetPopulator.populate(workbook);
    	if(result.isSuccess())
    		result = extrasSheetPopulator.populate(workbook);
    	if(result.isSuccess())
    		result = populateLoansTable(savingsTransactionSheet);
        if(result.isSuccess())
            result = setRules(savingsTransactionSheet);
        setDefaults(savingsTransactionSheet);
        return result;
    }
    
    private Result populateLoansTable(Sheet savingsTransactionSheet) {
    	Result result = new Result();
    	Workbook workbook = savingsTransactionSheet.getWorkbook();
    	CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
        dateCellStyle.setDataFormat(df);
		int rowIndex = 1;
    	Row row;
    	Collections.sort(savings, CompactSavingsAccount.ClientNameComparator);
    	try{
    		for(CompactSavingsAccount savingsAccount : savings) {
    			row = savingsTransactionSheet.createRow(rowIndex++);
    			writeString(LOOKUP_CLIENT_NAME_COL, row, savingsAccount.getClientName());
    			writeLong(LOOKUP_ACCOUNT_NO_COL, row, Long.parseLong(savingsAccount.getAccountNo()));
    			writeString(LOOKUP_PRODUCT_COL, row, savingsAccount.getSavingsProductName());
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
        worksheet.setColumnWidth(SAVINGS_ACCOUNT_NO_COL, 3000);
        worksheet.setColumnWidth(PRODUCT_COL, 4000);
        worksheet.setColumnWidth(OPENING_BALANCE_COL, 4000);
        worksheet.setColumnWidth(TRANSACTION_TYPE_COL, 3300);
        worksheet.setColumnWidth(AMOUNT_COL, 4000);
        worksheet.setColumnWidth(TRANSACTION_DATE_COL, 3000);
        worksheet.setColumnWidth(PAYMENT_TYPE_COL, 3000);
        worksheet.setColumnWidth(ACCOUNT_NO_COL, 3000);
        worksheet.setColumnWidth(CHECK_NO_COL, 3000);
        worksheet.setColumnWidth(RECEIPT_NO_COL, 3000);
        worksheet.setColumnWidth(ROUTING_CODE_COL, 3000);
        worksheet.setColumnWidth(BANK_NO_COL, 3000);
        worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 5000);
        worksheet.setColumnWidth(LOOKUP_ACCOUNT_NO_COL, 3000);
        worksheet.setColumnWidth(LOOKUP_PRODUCT_COL, 3000);
        worksheet.setColumnWidth(LOOKUP_OPENING_BALANCE_COL, 3700);
        worksheet.setColumnWidth(LOOKUP_SAVINGS_ACTIVATION_DATE_COL, 3500);
        writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
        writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
        writeString(SAVINGS_ACCOUNT_NO_COL, rowHeader, "Account No.*");
        writeString(PRODUCT_COL, rowHeader, "Product Name");
        writeString(OPENING_BALANCE_COL, rowHeader, "Opening Balance");
        writeString(TRANSACTION_TYPE_COL, rowHeader, "Transaction Type*");
        writeString(AMOUNT_COL, rowHeader, "Amount*");
        writeString(TRANSACTION_DATE_COL, rowHeader, "Date*");
        writeString(PAYMENT_TYPE_COL, rowHeader, "Type*");
        writeString(ACCOUNT_NO_COL, rowHeader, "Account No");
        writeString(CHECK_NO_COL, rowHeader, "Check No");
        writeString(RECEIPT_NO_COL, rowHeader, "Receipt No");
        writeString(ROUTING_CODE_COL, rowHeader, "Routing Code");
        writeString(BANK_NO_COL, rowHeader, "Bank No");
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
        	CellRangeAddressList transactionTypeRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), TRANSACTION_TYPE_COL, TRANSACTION_TYPE_COL);
        	CellRangeAddressList paymentTypeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), PAYMENT_TYPE_COL, PAYMENT_TYPE_COL);
        	CellRangeAddressList transactionDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), TRANSACTION_DATE_COL, TRANSACTION_DATE_COL);
        	
        	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
        	
        	setNames(worksheet);
        	
        	DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
        	DataValidationConstraint clientNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Client_\",$A1))");
        	DataValidationConstraint accountNumberConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Account_\",SUBSTITUTE($B1,\" \",\"_\")))");
        	DataValidationConstraint transactionTypeConstraint = validationHelper.createExplicitListConstraint(new String[] {"Withdrawal","Deposit"});
        	DataValidationConstraint paymentTypeConstraint = validationHelper.createFormulaListConstraint("PaymentTypes");
        	DataValidationConstraint transactionDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=VLOOKUP($C1,$Q$2:$T$" + (savings.size() + 1) + ",4,FALSE)", "=TODAY()", "dd/mm/yy");
        	
        	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
        	DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
        	DataValidation accountNumberValidation = validationHelper.createValidation(accountNumberConstraint, accountNumberRange);
        	DataValidation transactionTypeValidation = validationHelper.createValidation(transactionTypeConstraint, transactionTypeRange);
        	DataValidation paymentTypeValidation = validationHelper.createValidation(paymentTypeConstraint, paymentTypeRange);
        	DataValidation transactionDateValidation = validationHelper.createValidation(transactionDateConstraint, transactionDateRange);
        	
        	worksheet.addValidationData(officeValidation);
            worksheet.addValidationData(clientValidation);
            worksheet.addValidationData(accountNumberValidation);
            worksheet.addValidationData(transactionTypeValidation);
            worksheet.addValidationData(paymentTypeValidation);
            worksheet.addValidationData(transactionDateValidation);
        	
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
    			writeFormula(PRODUCT_COL, row, "IF(ISERROR(VLOOKUP($C"+ (rowNo+1) +",$Q$2:$S$" + (savings.size() + 1) + ",2,FALSE)),\"\",VLOOKUP($C"+ (rowNo+1) +",$Q$2:$S$" + (savings.size() + 1) + ",2,FALSE))");
    			writeFormula(OPENING_BALANCE_COL, row, "IF(ISERROR(VLOOKUP($C"+ (rowNo+1) +",$Q$2:$S$" + (savings.size() + 1) + ",3,FALSE)),\"\",VLOOKUP($C"+ (rowNo+1) +",$Q$2:$S$" + (savings.size() + 1) + ",3,FALSE))");
    		}
    	} catch (Exception e) {
    		logger.error(e.getMessage());
    	}
    }
    
    private void setNames(Sheet worksheet) {
    	Workbook savingsTransactionWorkbook = worksheet.getWorkbook();
    	ArrayList<String> officeNames = new ArrayList<String>(Arrays.asList(officeSheetPopulator.getOfficeNames()));
    	
    	//Office Names
    	Name officeGroup = savingsTransactionWorkbook.createName();
    	officeGroup.setNameName("Office");
    	officeGroup.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));
    	
    	//Clients Named after Offices
    	for(Integer i = 0; i < officeNames.size(); i++) {
    		Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator.getOfficeNameToBeginEndIndexesOfClients().get(i);
    		Name name = savingsTransactionWorkbook.createName();
    		if(officeNameToBeginEndIndexesOfClients != null) {
    	       name.setNameName("Client_" + officeNames.get(i));
    	       name.setRefersToFormula("Clients!$B$" + officeNameToBeginEndIndexesOfClients[0] + ":$B$" + officeNameToBeginEndIndexesOfClients[1]);
    		}
    	}
    	
    	//Counting clients with active savings and starting and end addresses of cells for naming
    	HashMap<String, Integer[]> clientNameToBeginEndIndexes = new HashMap<String, Integer[]>();
    	ArrayList<String> clientsWithActiveSavings = new ArrayList<String>();
    	int startIndex = 1, endIndex = 1;
    	String clientName = "";
    	
    	for(int i = 0; i < savings.size(); i++){
    		if(!clientName.equals(savings.get(i).getClientName())) {
    			endIndex = i + 1;
    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
    			startIndex = i + 2;
    			clientName = savings.get(i).getClientName();
    			clientsWithActiveSavings.add(clientName);
    		}
    		if(i == savings.size()-1) {
    			endIndex = i + 2;
    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
    		}
    	}
    	
    	//Account Number Named  after Clients
    	for(int j = 0; j < clientsWithActiveSavings.size(); j++) {
    		Name name = savingsTransactionWorkbook.createName();
    		name.setNameName("Account_" + clientsWithActiveSavings.get(j).replaceAll(" ", "_"));
    		name.setRefersToFormula("SavingsTransaction!$Q$" + clientNameToBeginEndIndexes.get(clientsWithActiveSavings.get(j))[0] + ":$Q$" + clientNameToBeginEndIndexes.get(clientsWithActiveSavings.get(j))[1]);
    	}
    	
    	//Payment Type Name
    	Name paymentTypeGroup = savingsTransactionWorkbook.createName();
    	paymentTypeGroup.setNameName("PaymentTypes");
    	paymentTypeGroup.setRefersToFormula("Extras!$D$2:$D$" + (extrasSheetPopulator.getPaymentTypesSize() + 1));
    }
}
