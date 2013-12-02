package org.openmf.mifos.dataimport.populator.loan;

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
import org.openmf.mifos.dataimport.dto.loan.CompactLoan;
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

public class LoanRepaymentWorkbookPopulator extends AbstractWorkbookPopulator {
	
    private static final Logger logger = LoggerFactory.getLogger(LoanRepaymentWorkbookPopulator.class);
	
	private final RestClient restClient;
	
	private String content;
	
	private OfficeSheetPopulator officeSheetPopulator;
	private ClientSheetPopulator clientSheetPopulator;
	private ExtrasSheetPopulator extrasSheetPopulator;
	private List<CompactLoan> loans;
	
	private static final int OFFICE_NAME_COL = 0;
    private static final int CLIENT_NAME_COL = 1;
    private static final int LOAN_ACCOUNT_NO_COL = 2;
    private static final int PRODUCT_COL = 3;
    private static final int PRINCIPAL_COL = 4;
    private static final int AMOUNT_COL = 5;
    private static final int REPAID_ON_DATE_COL = 6;
    private static final int REPAYMENT_TYPE_COL = 7;
    private static final int ACCOUNT_NO_COL = 8;
    private static final int CHECK_NO_COL = 9;
    private static final int ROUTING_CODE_COL = 10;	
    private static final int RECEIPT_NO_COL = 11;
    private static final int BANK_NO_COL = 12;
    private static final int LOOKUP_CLIENT_NAME_COL = 14;
    private static final int LOOKUP_ACCOUNT_NO_COL = 15;
    private static final int LOOKUP_PRODUCT_COL = 16;
    private static final int LOOKUP_PRINCIPAL_COL = 17;
    private static final int LOOKUP_LOAN_DISBURSEMENT_DATE_COL = 18;
	
	public LoanRepaymentWorkbookPopulator(RestClient restClient, OfficeSheetPopulator officeSheetPopulator,
			ClientSheetPopulator clientSheetPopulator, ExtrasSheetPopulator extrasSheetPopulator) {
        this.restClient = restClient;
        this.officeSheetPopulator = officeSheetPopulator;
        this.clientSheetPopulator = clientSheetPopulator;
        this.extrasSheetPopulator = extrasSheetPopulator;
		loans = new ArrayList<CompactLoan>();
    }
	
	@Override
    public Result downloadAndParse() {
		Result result =  officeSheetPopulator.downloadAndParse();
		if(result.isSuccess())
			result = clientSheetPopulator.downloadAndParse();
		if(result.isSuccess())
    		result = extrasSheetPopulator.downloadAndParse();
		if(result.isSuccess())
			result = parseLoans();
    	return result;
    }

    @Override
    public Result populate(Workbook workbook) {
    	Sheet loanRepaymentSheet = workbook.createSheet("LoanRepayment");
    	setLayout(loanRepaymentSheet);
    	Result result = officeSheetPopulator.populate(workbook);
    	if(result.isSuccess())
    		result = clientSheetPopulator.populate(workbook);
    	if(result.isSuccess())
    		result = extrasSheetPopulator.populate(workbook);
    	if(result.isSuccess())
    		result = populateLoansTable(loanRepaymentSheet);
        if(result.isSuccess())
            result = setRules(loanRepaymentSheet);
        setDefaults(loanRepaymentSheet);
        return result;
    }
    
    private Result parseLoans() {
    	Result result = new Result();
    	try {
        	restClient.createAuthToken();
            content = restClient.get("loans?limit=-1");
            Gson gson = new Gson();
            JsonParser parser = new JsonParser();
            JsonObject obj = parser.parse(content).getAsJsonObject();
            JsonArray array = obj.getAsJsonArray("pageItems");
            Iterator<JsonElement> iterator = array.iterator();
            while(iterator.hasNext()) {
            	JsonElement json = iterator.next();
            	CompactLoan loan = gson.fromJson(json, CompactLoan.class);
            	if(loan.isActive())
            	  loans.add(loan);
            } 
       } catch (Exception e) {
           result.addError(e.getMessage());
           logger.error(e.getMessage());
       }
       return result;	
    }
    
    private Result populateLoansTable(Sheet loanRepaymentSheet) {
    	Result result = new Result();
    	int rowIndex = 1;
    	Row row;
    	Workbook workbook = loanRepaymentSheet.getWorkbook();
    	CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
        dateCellStyle.setDataFormat(df);
    	Collections.sort(loans, CompactLoan.ClientNameComparator);
    	try{
    		for(CompactLoan loan : loans) {
    			row = loanRepaymentSheet.createRow(rowIndex++);
    			writeString(LOOKUP_CLIENT_NAME_COL, row, loan.getClientName());
    			writeLong(LOOKUP_ACCOUNT_NO_COL, row, Long.parseLong(loan.getAccountNo()));
    			writeString(LOOKUP_PRODUCT_COL, row, loan.getLoanProductName());
    			writeDouble(LOOKUP_PRINCIPAL_COL, row, loan.getPrincipal());
    			writeDate(LOOKUP_LOAN_DISBURSEMENT_DATE_COL, row, loan.getTimeline().getActualDisbursementDate().get(2) + "/" + loan.getTimeline().getActualDisbursementDate().get(1) + "/" + loan.getTimeline().getActualDisbursementDate().get(0), dateCellStyle);
    		}
	   } catch (Exception e) {
		result.addError(e.getMessage());
	    }
    	return result;
    }

    private void setLayout(Sheet worksheet) {
    	Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
        worksheet.setColumnWidth(CLIENT_NAME_COL, 5000);
        worksheet.setColumnWidth(LOAN_ACCOUNT_NO_COL, 3000);
        worksheet.setColumnWidth(PRODUCT_COL, 4000);
        worksheet.setColumnWidth(PRINCIPAL_COL, 4000);
        worksheet.setColumnWidth(AMOUNT_COL, 4000);
        worksheet.setColumnWidth(REPAID_ON_DATE_COL, 3000);
        worksheet.setColumnWidth(REPAYMENT_TYPE_COL, 3000);
        worksheet.setColumnWidth(ACCOUNT_NO_COL, 3000);
        worksheet.setColumnWidth(CHECK_NO_COL, 3000);
        worksheet.setColumnWidth(RECEIPT_NO_COL, 3000);
        worksheet.setColumnWidth(ROUTING_CODE_COL, 3000);
        worksheet.setColumnWidth(BANK_NO_COL, 3000);
        worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 5000);
        worksheet.setColumnWidth(LOOKUP_ACCOUNT_NO_COL, 3000);
        worksheet.setColumnWidth(LOOKUP_PRODUCT_COL, 3000);
        worksheet.setColumnWidth(LOOKUP_PRINCIPAL_COL, 3700);
        worksheet.setColumnWidth(LOOKUP_LOAN_DISBURSEMENT_DATE_COL, 3700);
        writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
        writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
        writeString(LOAN_ACCOUNT_NO_COL, rowHeader, "Account No.*");
        writeString(PRODUCT_COL, rowHeader, "Product Name");
        writeString(PRINCIPAL_COL, rowHeader, "Principal");
        writeString(AMOUNT_COL, rowHeader, "Amount Repaid*");
        writeString(REPAID_ON_DATE_COL, rowHeader, "Date*");
        writeString(REPAYMENT_TYPE_COL, rowHeader, "Type*");
        writeString(ACCOUNT_NO_COL, rowHeader, "Account No");
        writeString(CHECK_NO_COL, rowHeader, "Check No");
        writeString(RECEIPT_NO_COL, rowHeader, "Receipt No");
        writeString(ROUTING_CODE_COL, rowHeader, "Routing Code");
        writeString(BANK_NO_COL, rowHeader, "Bank No");
        writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Lookup Client");
        writeString(LOOKUP_ACCOUNT_NO_COL, rowHeader, "Lookup Account");
        writeString(LOOKUP_PRODUCT_COL, rowHeader, "Lookup Product");
        writeString(LOOKUP_PRINCIPAL_COL, rowHeader, "Lookup Principal");
        writeString(LOOKUP_LOAN_DISBURSEMENT_DATE_COL, rowHeader, "Lookup Loan Disbursement Date");
        
    }
    
    private Result setRules(Sheet worksheet) {
    	Result result = new Result();
    	try {
    		CellRangeAddressList officeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), OFFICE_NAME_COL, OFFICE_NAME_COL);
        	CellRangeAddressList clientNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), CLIENT_NAME_COL, CLIENT_NAME_COL);
        	CellRangeAddressList accountNumberRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_ACCOUNT_NO_COL, LOAN_ACCOUNT_NO_COL);
        	CellRangeAddressList repaymentTypeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAYMENT_TYPE_COL, REPAYMENT_TYPE_COL);
        	CellRangeAddressList repaymentDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAID_ON_DATE_COL, REPAID_ON_DATE_COL);
        	
        	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
        
        	setNames(worksheet);
        	
        	DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
        	DataValidationConstraint clientNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Client_\",$A1))");
        	DataValidationConstraint accountNumberConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Account_\",SUBSTITUTE($B1,\" \",\"_\")))");
        	DataValidationConstraint paymentTypeConstraint = validationHelper.createFormulaListConstraint("PaymentTypes");
        	DataValidationConstraint repaymentDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=VLOOKUP($C1,$P$2:$S$" + (loans.size() + 1) + ",4,FALSE)", "=TODAY()", "dd/mm/yy");
        	
        	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
        	DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
        	DataValidation accountNumberValidation = validationHelper.createValidation(accountNumberConstraint, accountNumberRange);
        	DataValidation repaymentTypeValidation = validationHelper.createValidation(paymentTypeConstraint, repaymentTypeRange);
        	DataValidation repaymentDateValidation = validationHelper.createValidation(repaymentDateConstraint, repaymentDateRange);
        	
        	worksheet.addValidationData(officeValidation);
            worksheet.addValidationData(clientValidation);
            worksheet.addValidationData(accountNumberValidation);
            worksheet.addValidationData(repaymentTypeValidation);
            worksheet.addValidationData(repaymentDateValidation);
        	
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
    			writeFormula(PRODUCT_COL, row, "IF(ISERROR(VLOOKUP($C"+ (rowNo+1) +",$P$2:$R$" + (loans.size() + 1) + ",2,FALSE)),\"\",VLOOKUP($C"+ (rowNo+1) +",$P$2:$R$" + (loans.size() + 1) + ",2,FALSE))");
    			writeFormula(PRINCIPAL_COL, row, "IF(ISERROR(VLOOKUP($C"+ (rowNo+1) +",$P$2:$R$" + (loans.size() + 1) + ",3,FALSE)),\"\",VLOOKUP($C"+ (rowNo+1) +",$P$2:$R$" + (loans.size() + 1) + ",3,FALSE))");
    		}
    	} catch (Exception e) {
    		logger.error(e.getMessage());
    	}
    }
    
    private void setNames(Sheet worksheet) {
    	ArrayList<String> officeNames = new ArrayList<String>(Arrays.asList(officeSheetPopulator.getOfficeNames()));
    	Workbook loanRepaymentWorkbook = worksheet.getWorkbook();
    	//Office Names
    	Name officeGroup = loanRepaymentWorkbook.createName();
    	officeGroup.setNameName("Office");
    	officeGroup.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));
    	
    	//Clients Named after Offices
    	for(Integer i = 0; i < officeNames.size(); i++) {
    		Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator.getOfficeNameToBeginEndIndexesOfClients().get(i);
    		Name name = loanRepaymentWorkbook.createName();
    		if(officeNameToBeginEndIndexesOfClients != null) {
    	       name.setNameName("Client_" + officeNames.get(i));
    	       name.setRefersToFormula("Clients!$B$" + officeNameToBeginEndIndexesOfClients[0] + ":$B$" + officeNameToBeginEndIndexesOfClients[1]);
    		}
    	}
    	
    	//Counting clients with active loans and starting and end addresses of cells
    	HashMap<String, Integer[]> clientNameToBeginEndIndexes = new HashMap<String, Integer[]>();
    	ArrayList<String> clientsWithActiveLoans = new ArrayList<String>();
    	int startIndex = 1, endIndex = 1;
    	String clientName = "";
    	
    	for(int i = 0; i < loans.size(); i++){
    		if(!clientName.equals(loans.get(i).getClientName())) {
    			endIndex = i + 1;
    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
    			startIndex = i + 2;
    			clientName = loans.get(i).getClientName();
    			clientsWithActiveLoans.add(clientName);
    		}
    		if(i == loans.size()-1) {
    			endIndex = i + 2;
    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
    		}
    	}
    	
    	//Account Number Named  after Clients
    	for(int j = 0; j < clientsWithActiveLoans.size(); j++) {
    		Name name = loanRepaymentWorkbook.createName();
    		name.setNameName("Account_" + clientsWithActiveLoans.get(j).replaceAll(" ", "_"));
    		name.setRefersToFormula("LoanRepayment!$P$" + clientNameToBeginEndIndexes.get(clientsWithActiveLoans.get(j))[0] + ":$P$" + clientNameToBeginEndIndexes.get(clientsWithActiveLoans.get(j))[1]);
    	}
    	
    	//Payment Type Name
    	Name paymentTypeGroup = loanRepaymentWorkbook.createName();
    	paymentTypeGroup.setNameName("PaymentTypes");
    	paymentTypeGroup.setRefersToFormula("Extras!$D$2:$D$" + (extrasSheetPopulator.getPaymentTypesSize() + 1));
    }

}
