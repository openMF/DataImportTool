package org.openmf.mifos.dataimport.handler.savings;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.ClosingOfSavingsAccounts;
import org.openmf.mifos.dataimport.handler.AbstractDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;

public class ClosingOfSavingsAccountHandler  extends AbstractDataImportHandler{
	 private static final Logger logger = LoggerFactory.getLogger(ClosingOfSavingsAccountHandler.class);
		
		private final RestClient restClient;
		 private final Workbook workbook;
		
	    private List<ClosingOfSavingsAccounts> closedOnDate = new ArrayList<ClosingOfSavingsAccounts>();
	    private String savingsAccountId = "";
	    
	private static final int ACCOUNT_TYPE_COL=2;
	private static final int SAVINGS_ACCOUNT_NO_COL = 3;
	private static final int CLOSED_ON_DATE = 6;
    private static final int ON_ACCOUNT_CLOSURE_ID = 7;
    private static final int TO_SAVINGS_ACCOUNT_ID = 8;
    private static final int STATUS_COL = 10;
   
    
    public ClosingOfSavingsAccountHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
    }

	@Override
	public Result parse() {
		Result result = new Result();
        Sheet savingsSheet = workbook.getSheet("ClosingOfSavingsAccounts");
        Integer noOfEntries = getNumberOfRows(savingsSheet, 0);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = savingsSheet.getRow(rowIndex);
                if (isNotImported(row, STATUS_COL)) {
                    closedOnDate.add(paseAsSavingsClosed(row));
                }
            } catch (RuntimeException re) {
            	logger.error("row = " + rowIndex, re);
                result.addError("Row = " + rowIndex + " , " + re.getMessage());
            }
        }
        return result;
	}

	

	
private ClosingOfSavingsAccounts paseAsSavingsClosed(Row row) {
	 String savingsAccountIdCheck = readAsInt(SAVINGS_ACCOUNT_NO_COL, row);
     if(!savingsAccountIdCheck.equals(""))
     	savingsAccountId = savingsAccountIdCheck;
     String closedOnDate = readAsDate(CLOSED_ON_DATE,row);
     String onAccountClosure = readAsLong(ON_ACCOUNT_CLOSURE_ID,row);
     String toSavingsAccountId = readAsLong(TO_SAVINGS_ACCOUNT_ID, row);
     String accountType = readAsString(ACCOUNT_TYPE_COL, row);
     return new ClosingOfSavingsAccounts(Integer.parseInt(savingsAccountId),accountType,closedOnDate,onAccountClosure,toSavingsAccountId,  row.getRowNum());	
	}

@Override
public Result upload() {
	Result result = new Result();
    Sheet savingsTransactionSheet = workbook.getSheet("ClosingOfSavingsAccounts");
    restClient.createAuthToken();
    for (ClosingOfSavingsAccounts transaction : closedOnDate) {
        try {
            Gson gson = new Gson();
            String payload = gson.toJson(transaction);
            restClient.post(transaction.getAccountType() +"/"+transaction.getAccountId()+ "?command=close", payload);
            
            Cell statusCell = savingsTransactionSheet.getRow(transaction.getRowIndex()).createCell(STATUS_COL);
            statusCell.setCellValue("Imported");
            statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
        } catch (Exception e) {
        	Cell savingsAccountIdCell = savingsTransactionSheet.getRow(transaction.getRowIndex()).createCell(SAVINGS_ACCOUNT_NO_COL);
        	savingsAccountIdCell.setCellValue(transaction.getAccountId());
        	String message = parseStatus(e.getMessage());
        	
        	Cell statusCell = savingsTransactionSheet.getRow(transaction.getRowIndex()).createCell(STATUS_COL);
        	statusCell.setCellValue(message);
            statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));
            result.addError("Row = " + transaction.getRowIndex() + " ," + message);
        }
    }
    savingsTransactionSheet.setColumnWidth(STATUS_COL, 15000);
	writeString(STATUS_COL, savingsTransactionSheet.getRow(0), "Status");
	return result;
}


   

public static int getClosedOnDate() {
	return CLOSED_ON_DATE;
}

public static int getOnAccountClosureId() {
	return ON_ACCOUNT_CLOSURE_ID;
}

public static int getToSavingsAccountId() {
	return TO_SAVINGS_ACCOUNT_ID;
}

	}
	
	


