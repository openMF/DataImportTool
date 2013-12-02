package org.openmf.mifos.dataimport.handler.loan;

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

public class LoanRepaymentDataImportHandler extends AbstractDataImportHandler {

	private static final Logger logger = LoggerFactory.getLogger(LoanRepaymentDataImportHandler.class);
	
    private final RestClient restClient;
    private final Workbook workbook;
    
    private List<Transaction> loanRepayments;
    private String loanAccountId = "";
    
    private static final int LOAN_ACCOUNT_NO_COL = 2;
	private static final int AMOUNT_COL = 5;
    private static final int REPAID_ON_DATE_COL = 6;
    private static final int REPAYMENT_TYPE_COL = 7;
    private static final int ACCOUNT_NO_COL = 8;
    private static final int CHECK_NO_COL = 9;
    private static final int ROUTING_CODE_COL = 10;	
    private static final int RECEIPT_NO_COL = 11;
    private static final int BANK_NO_COL = 12;
    private static final int STATUS_COL = 13;

    public LoanRepaymentDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
        loanRepayments = new ArrayList<Transaction>();
    }
    
    @Override
    public Result parse() {
        Result result = new Result();
        Sheet loanRepaymentSheet = workbook.getSheet("LoanRepayment");
        Integer noOfEntries = getNumberOfRows(loanRepaymentSheet, AMOUNT_COL);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = loanRepaymentSheet.getRow(rowIndex);
                if(isNotImported(row, STATUS_COL))
                    loanRepayments.add(parseAsLoanRepayment(row));
            } catch (Exception e) {
                logger.error("row = " + rowIndex, e);
                result.addError("Row = " + rowIndex + " , " + e.getMessage());
            }
        }
        return result;
    }
        
        private Transaction parseAsLoanRepayment(Row row) {
        	 String loanAccountIdCheck = readAsInt(LOAN_ACCOUNT_NO_COL, row);
             if(!loanAccountIdCheck.equals(""))
                loanAccountId = loanAccountIdCheck;
             String repaymentAmount = readAsDouble(AMOUNT_COL, row).toString();
             String repaymentDate = readAsDate(REPAID_ON_DATE_COL, row);
             String repaymentType = readAsString(REPAYMENT_TYPE_COL, row);
             String repaymentTypeId = getIdByName(workbook.getSheet("Extras"), repaymentType).toString();
             String accountNumber = readAsLong(ACCOUNT_NO_COL, row);
             String checkNumber = readAsLong(CHECK_NO_COL, row);
             String routingCode = readAsLong(ROUTING_CODE_COL, row);
             String receiptNumber = readAsLong(RECEIPT_NO_COL, row);
             String bankNumber = readAsLong(BANK_NO_COL, row);
             return new Transaction(repaymentAmount, repaymentDate, repaymentTypeId, accountNumber,
             		checkNumber, routingCode, receiptNumber, bankNumber, Integer.parseInt(loanAccountId), "", row.getRowNum());
        }
    
    @Override
    public Result upload() {
        Result result = new Result();
        Sheet loanRepaymentSheet = workbook.getSheet("LoanRepayment");
        restClient.createAuthToken();
        for (Transaction loanRepayment : loanRepayments) {
            try {
                Gson gson = new Gson();
                String payload = gson.toJson(loanRepayment);
                logger.info("ID: "+loanRepayment.getAccountId()+" : "+payload);
                restClient.post("loans/" + loanRepayment.getAccountId() + "/transactions?command=repayment", payload);
                Cell statusCell = loanRepaymentSheet.getRow(loanRepayment.getRowIndex()).createCell(STATUS_COL);
                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (Exception e) {
            	Cell loanAccountIdCell = loanRepaymentSheet.getRow(loanRepayment.getRowIndex()).createCell(LOAN_ACCOUNT_NO_COL);
                loanAccountIdCell.setCellValue(loanRepayment.getAccountId());
            	String message = parseStatus(e.getMessage());
            	Cell statusCell = loanRepaymentSheet.getRow(loanRepayment.getRowIndex()).createCell(STATUS_COL);
            	statusCell.setCellValue(message);
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));
                result.addError("Row = " + loanRepayment.getRowIndex() + " ," + message);
            }
        }
        loanRepaymentSheet.setColumnWidth(STATUS_COL, 15000);
    	writeString(STATUS_COL, loanRepaymentSheet.getRow(0), "Status");
        return result;
    }
    
    public List<Transaction> getLoanRepayments() {
    	return loanRepayments;
    }
}
