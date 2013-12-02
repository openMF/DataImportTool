package org.openmf.mifos.dataimport.handler.savings;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Transaction;
import org.openmf.mifos.dataimport.handler.AbstractDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;

public class SavingsTransactionDataImportHandler extends AbstractDataImportHandler {

	private static final Logger logger = LoggerFactory.getLogger(SavingsTransactionDataImportHandler.class);
	
    private final RestClient restClient;
    private final Workbook workbook;
    
    private List<Transaction> savingsTransactions;
    private String savingsAccountId = "";
    
    private static final int SAVINGS_ACCOUNT_NO_COL = 2;
    private static final int TRANSACTION_TYPE_COL = 5;
    private static final int AMOUNT_COL = 6;
    private static final int TRANSACTION_DATE_COL = 7;
    private static final int PAYMENT_TYPE_COL = 8;
    private static final int ACCOUNT_NO_COL = 9;
    private static final int CHECK_NO_COL = 10;
    private static final int ROUTING_CODE_COL = 11;	
    private static final int RECEIPT_NO_COL = 12;
    private static final int BANK_NO_COL = 13;
    private static final int STATUS_COL = 13;
    
    public SavingsTransactionDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
        savingsTransactions = new ArrayList<Transaction>();
    }
    
    @Override
    public Result parse() {
        Result result = new Result();
        Sheet savingsTransactionSheet = workbook.getSheet("SavingsTransaction");
        Integer noOfEntries = getNumberOfRows(savingsTransactionSheet, AMOUNT_COL);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = savingsTransactionSheet.getRow(rowIndex);
                if(isNotImported(row, STATUS_COL))
                	savingsTransactions.add(parseAsTransaction(row));
            } catch (Exception e) {
                logger.error("row = " + rowIndex, e);
                result.addError("Row = " + rowIndex + " , " + e.getMessage());
            }
        }
        return result;
    }
    
    private Transaction parseAsTransaction(Row row) {
   	    String savingsAccountIdCheck = readAsInt(SAVINGS_ACCOUNT_NO_COL, row);
        if(!savingsAccountIdCheck.equals(""))
        	savingsAccountId = savingsAccountIdCheck;
        String transactionType = readAsString(TRANSACTION_TYPE_COL, row);
        String amount = readAsDouble(AMOUNT_COL, row).toString();
        String transactionDate = readAsDate(TRANSACTION_DATE_COL, row);
        String paymentType = readAsString(PAYMENT_TYPE_COL, row);
        String paymentTypeId = getIdByName(workbook.getSheet("Extras"), paymentType).toString();
        String accountNumber = readAsLong(ACCOUNT_NO_COL, row);
        String checkNumber = readAsLong(CHECK_NO_COL, row);
        String routingCode = readAsLong(ROUTING_CODE_COL, row);
        String receiptNumber = readAsLong(RECEIPT_NO_COL, row);
        String bankNumber = readAsLong(BANK_NO_COL, row);
        return new Transaction(amount, transactionDate, paymentTypeId, accountNumber,
        		checkNumber, routingCode, receiptNumber, bankNumber, Integer.parseInt(savingsAccountId), transactionType, row.getRowNum());
   }
    
    @Override
    public Result upload() {
        Result result = new Result();
        Sheet savingsTransactionSheet = workbook.getSheet("SavingsTransaction");
        restClient.createAuthToken();
        for (Transaction transaction : savingsTransactions) {
            try {
                Gson gson = new Gson();
                String payload = gson.toJson(transaction);
                logger.info("ID: "+transaction.getAccountId()+" : "+payload);
                restClient.post("savingsaccounts/" + transaction.getAccountId() + "/transactions?command=" + transaction.getTransactionType(), payload);
                
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
    
    public List<Transaction> getSavingsTransactions() {
    	return savingsTransactions;
    }
    
}
