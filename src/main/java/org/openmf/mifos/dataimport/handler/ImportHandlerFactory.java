package org.openmf.mifos.dataimport.handler;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.handler.accounting.AddJournalEntriesHandler;
import org.openmf.mifos.dataimport.handler.client.CenterDataImportHandler;
import org.openmf.mifos.dataimport.handler.client.ClientDataImportHandler;
import org.openmf.mifos.dataimport.handler.client.GroupDataImportHandler;
import org.openmf.mifos.dataimport.handler.loan.AddGuarantorDataImportHandler;
import org.openmf.mifos.dataimport.handler.loan.LoanDataImportHandler;
import org.openmf.mifos.dataimport.handler.loan.LoanRepaymentDataImportHandler;
import org.openmf.mifos.dataimport.handler.savings.ClosingOfSavingsAccountHandler;
import org.openmf.mifos.dataimport.handler.savings.FixedDepositImportHandler;
import org.openmf.mifos.dataimport.handler.savings.RecurringDepositAccountTransactionDataImportHandler;
import org.openmf.mifos.dataimport.handler.savings.RecurringDepositImportHandler;
import org.openmf.mifos.dataimport.handler.savings.SavingsDataImportHandler;
import org.openmf.mifos.dataimport.handler.savings.SavingsTransactionDataImportHandler;
import org.openmf.mifos.dataimport.http.MifosRestClient;


public class ImportHandlerFactory {
    
    public static final DataImportHandler createImportHandler(Workbook workbook) throws IOException {
        
        if(workbook.getSheetIndex("Clients") == 0) {
            	return new ClientDataImportHandler(workbook, new MifosRestClient());
        }else if(workbook.getSheetIndex("Groups") == 0) {
    	    return new GroupDataImportHandler(workbook, new MifosRestClient());
        }else if(workbook.getSheetIndex("Centers") == 0) {
    	    return new CenterDataImportHandler(workbook, new MifosRestClient());
        }else if(workbook.getSheetIndex("Loans") == 0) {
        	    return new LoanDataImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("LoanRepayment") == 0) {
        	    return new LoanRepaymentDataImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("Savings") == 0) {
    	    return new SavingsDataImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("SavingsTransaction") == 0) {
    	    return new SavingsTransactionDataImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("FixedDeposit") == 0) {
        	return new FixedDepositImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("RecurringDeposit") == 0) {
        	return new RecurringDepositImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("RecurringDepositTransaction") == 0) {
        	return new RecurringDepositAccountTransactionDataImportHandler(workbook, new MifosRestClient());
        }else if(workbook.getSheetIndex("ClosingOfSavingsAccounts") == 0) {
            	return new ClosingOfSavingsAccountHandler(workbook, new MifosRestClient());
        }else if(workbook.getSheetIndex("AddJournalEntries") == 0) {
        	return new AddJournalEntriesHandler(workbook, new MifosRestClient());
        }else if(workbook.getSheetIndex("guarantor") == 0) {
        	return new AddGuarantorDataImportHandler(workbook, new MifosRestClient());
        }
        
        
        
        throw new IllegalArgumentException("No work sheet found for processing : active sheet " + workbook.getSheetName(0));
    }

}
