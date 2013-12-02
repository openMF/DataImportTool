package org.openmf.mifos.dataimport.handler;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.Mock;
import org.mockito.runners.MockitoJUnitRunner;
import org.openmf.mifos.dataimport.dto.Transaction;
import org.openmf.mifos.dataimport.handler.savings.SavingsTransactionDataImportHandler;
import org.openmf.mifos.dataimport.http.RestClient;

@RunWith(MockitoJUnitRunner.class)
public class SavingsTransactionDataImportHandlerTest {

	
	    @Mock
	    RestClient restClient;
	    
	    @Test
	    public void shouldParseSavingsTransaction() throws IOException {
	        
	        InputStream is = this.getClass().getClassLoader().getResourceAsStream("savings/savingsTransactionHistory.xls");
	        Workbook book = new HSSFWorkbook(is);
	        SavingsTransactionDataImportHandler handler = new SavingsTransactionDataImportHandler(book, restClient);
	        Result result = handler.parse();
	        Assert.assertTrue(result.isSuccess());
	        Assert.assertEquals(2, handler.getSavingsTransactions().size());
	        Transaction savingsTransaction = handler.getSavingsTransactions().get(0);
	        Transaction savingsTransactionWithoutId = handler.getSavingsTransactions().get(1);
	        Assert.assertEquals("7", savingsTransaction.getAccountId().toString());
	        Assert.assertEquals("7", savingsTransactionWithoutId.getAccountId().toString());
	        Assert.assertEquals("100.0", savingsTransaction.getTransactionAmount());
	        Assert.assertEquals("75.0", savingsTransactionWithoutId.getTransactionAmount());
	        Assert.assertEquals("06 September 2013", savingsTransaction.getTransactionDate());
	        Assert.assertEquals("07 September 2013", savingsTransactionWithoutId.getTransactionDate());
	        Assert.assertEquals("15", savingsTransaction.getPaymentTypeId());
	        Assert.assertEquals("15", handler.getIdByName(book.getSheet("Extras"), "Cash").toString());
	        Assert.assertEquals("17", savingsTransactionWithoutId.getPaymentTypeId());
	        Assert.assertEquals("17", handler.getIdByName(book.getSheet("Extras"), "Check").toString());
	        Assert.assertEquals("12", savingsTransaction.getAccountNumber());
	        Assert.assertEquals("13", savingsTransaction.getCheckNumber());
	        Assert.assertEquals("14", savingsTransaction.getRoutingCode());
	        Assert.assertEquals("15", savingsTransaction.getReceiptNumber());
	        Assert.assertEquals("16", savingsTransaction.getBankNumber());
	    }
}
