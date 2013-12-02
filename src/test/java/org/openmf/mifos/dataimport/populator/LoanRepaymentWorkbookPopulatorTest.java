package org.openmf.mifos.dataimport.populator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.runners.MockitoJUnitRunner;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.openmf.mifos.dataimport.populator.loan.LoanRepaymentWorkbookPopulator;

@RunWith(MockitoJUnitRunner.class)
public class LoanRepaymentWorkbookPopulatorTest {

	@Mock
	RestClient restClient;

	private static final int LOOKUP_CLIENT_NAME_COL = 14;
    private static final int LOOKUP_ACCOUNT_NO_COL = 15;
    private static final int LOOKUP_PRODUCT_COL = 16;
    private static final int LOOKUP_PRINCIPAL_COL = 17;
    
    @Test
    public void shouldPopulateLoanRepaymentWorkbook() {
    	
    	Mockito.when(restClient.get("loans?limit=-1")).thenReturn("{\"totalFilteredRecords\": 1,\"pageItems\": [{" +
      "\"id\": 3,\"accountNo\": \"000000003\",\"status\": {\"id\": 300,\"code\": \"loanStatusType.active\",\"value\": \"Active\",\"pendingApproval\": false,\"waitingForDisbursal\": false," +
      "\"active\": true,\"closedObligationsMet\": false,\"closedWrittenOff\": false,\"closedRescheduled\": false,\"closed\": false,\"overpaid\": false},\"clientId\": 2," +
      "\"clientName\": \"Billy T Bob\",\"clientOfficeId\": 2,\"loanProductId\": 1,\"loanProductName\": \"HM\",\"loanProductDescription\": \"HM\",\"fundId\": 1,\"fundName\": \"Fund1\"," +
      "\"loanOfficerId\": 2,\"loanOfficerName\": \"Dzeko, Edin\",\"loanType\": {\"id\": 1, \"code\": \"accountType.individual\",\"value\": \"Individual\"}," +
      "\"currency\": {\"code\": \"USD\",\"name\": \"US Dollar\",\"decimalPlaces\": 2,\"inMultiplesOf\": 5,\"displaySymbol\": \"$\",\"nameCode\": \"currency.USD\"," +
      "\"displayLabel\": \"US Dollar ($)\"}, \"principal\": 25000.000000, \"termFrequency\": 12, \"termPeriodFrequencyType\": {\"id\": 2,\"code\": \"termFrequency.periodFrequencyType.months\"," +
      "\"value\": \"Months\"},\"numberOfRepayments\": 12,\"repaymentEvery\": 1,\"repaymentFrequencyType\": {\"id\": 2,\"code\": \"repaymentFrequency.periodFrequencyType.months\"," +
      "\"value\": \"Months\"},\"interestRatePerPeriod\": 7.000000,\"interestRateFrequencyType\": {\"id\": 3,\"code\": \"interestRateFrequency.periodFrequencyType.years\"," +
      "\"value\": \"Per year\"},\"annualInterestRate\": 7.000000,\"amortizationType\": {\"id\": 0,\"code\": \"amortizationType.equal.principal\",\"value\": \"Equal principal payments\"}," +
      "\"interestType\": {\"id\": 0,\"code\": \"interestType.declining.balance\",\"value\": \"Declining Balance\"},\"interestCalculationPeriodType\": {\"id\": 0," +
      "\"code\": \"interestCalculationPeriodType.daily\",\"value\": \"Daily\"},\"inArrearsTolerance\": 3.000000,\"transactionProcessingStrategyId\": 4,\"transactionProcessingStrategyName\": \"RBI (India)\"," +
      "\"graceOnPrincipalPayment\": 1,\"graceOnInterestPayment\": 2,\"graceOnInterestCharged\": 1,\"syncDisbursementWithMeeting\": false,\"timeline\": {\"submittedOnDate\": [2013,8,2]," +
      "\"submittedByUsername\": \"mifos\",\"submittedByFirstname\": \"App\",\"submittedByLastname\": \"Administrator\",\"approvedOnDate\": [2013,8,3],\"approvedByUsername\": \"mifos\"," +
      "\"approvedByFirstname\": \"App\",\"approvedByLastname\": \"Administrator\",\"expectedDisbursementDate\": [2013,8,2],\"actualDisbursementDate\": [2013,8,4]," +
      "\"disbursedByUsername\": \"mifos\", \"disbursedByFirstname\": \"App\", \"disbursedByLastname\": \"Administrator\",\"expectedMaturityDate\": [2014,8,4]}," +
      "\"summary\": {\"currency\": {\"code\": \"USD\",\"name\": \"US Dollar\",\"decimalPlaces\": 2,\"inMultiplesOf\": 5,\"displaySymbol\": \"$\",\"nameCode\": \"currency.USD\"," +
      "\"displayLabel\": \"US Dollar ($)\"}},\"feeChargesAtDisbursementCharged\": 0,\"loanCounter\": 1,\"loanProductCounter\": 1}]}");
    	
    	Mockito.when(restClient.get("clients?limit=-1")).thenReturn("{\"totalFilteredRecords\": 2,\"pageItems\": [{\"id\": 1,\"accountNo\": \"000000001\"," +
    	 		"\"status\": {\"id\": 300,\"code\": \"clientStatusType.active\",\"value\": \"Active\"},\"active\": true,\"activationDate\": [2013,7,1]," +
    	 		"\"firstname\": \"Arsene\",\"middlename\": \"K\",\"lastname\": \"Wenger\",\"displayName\": \"Arsene K Wenger\",\"officeId\": 1," +
    	 		"\"officeName\": \"Head Office\",\"staffId\": 1,\"staffName\": \"Chatta, Sahil\"},{\"id\": 2,\"accountNo\": \"000000002\"," +
    	 		"\"status\": {\"id\": 300,\"code\": \"clientStatusType.active\",\"value\": \"Active\"},\"active\": true,\"activationDate\": [2013,7,1]," +
    	 		"\"firstname\": \"Billy\",\"middlename\": \"T\",\"lastname\": \"Bob\",\"displayName\": \"Billy T Bob\",\"officeId\": 2,\"officeName\": \"Office1\"," +
    	 		"\"staffId\": 2,\"staffName\": \"Dzeko, Edin\"}]}");
		
		Mockito.when(restClient.get("offices?limit=-1")).thenReturn("[{\"id\":1,\"name\":\"Head Office\",\"nameDecorated\":\"Head Office\",\"externalId\": \"1\"," +
        		"\"openingDate\":[2009,1,1],\"hierarchy\": \".\"},{\"id\": 2,\"name\": \"Office1\",\"nameDecorated\": \"....Office1\",\"openingDate\":[2013,4,1]," +
        		"\"hierarchy\": \".2.\",\"parentId\": 1,\"parentName\": \"Head Office\"}]");
		
		Mockito.when(restClient.get("funds")).thenReturn("[{\"id\": 1,\"name\": \"Fund1\"}]");
        Mockito.when(restClient.get("codes/12/codevalues")).thenReturn("[{\"id\": 10,\"name\": \"Cash\",\"position\": 1},{\"id\": 11,\"name\": \"MPesa\",\"position\": 2}]");
    	
    	LoanRepaymentWorkbookPopulator loanRepaymentWorkbookPopulator = new LoanRepaymentWorkbookPopulator(restClient,
    			new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
    	loanRepaymentWorkbookPopulator.downloadAndParse();
    	
    	Workbook loanRepaymentWorkbook = new HSSFWorkbook();
    	Result result = loanRepaymentWorkbookPopulator.populate(loanRepaymentWorkbook);
    	Assert.assertTrue(result.isSuccess());
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("loans?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("clients?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("funds");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("codes/12/codevalues");
    	
    	Sheet loanRepaymentSheet = loanRepaymentWorkbook.getSheet("LoanRepayment");
    	Row row = loanRepaymentSheet.getRow(0);
    	
    	//If test fails, also check if column letters embedded in formulas in setDefault and setRules have changed or not.
    	Assert.assertEquals("Office Name*", row.getCell(0).getStringCellValue());
    	Assert.assertEquals("Client Name*", row.getCell(1).getStringCellValue());
    	Assert.assertEquals("Account No.*", row.getCell(2).getStringCellValue());
    	Assert.assertEquals("Product Name", row.getCell(3).getStringCellValue());
    	Assert.assertEquals("Principal", row.getCell(4).getStringCellValue());
    	Assert.assertEquals("Amount Repaid*", row.getCell(5).getStringCellValue());
    	Assert.assertEquals("Date*", row.getCell(6).getStringCellValue());
    	Assert.assertEquals("Type*", row.getCell(7).getStringCellValue());
    	Assert.assertEquals("Account No", row.getCell(8).getStringCellValue());
    	Assert.assertEquals("Check No", row.getCell(9).getStringCellValue());
    	Assert.assertEquals("Routing Code", row.getCell(10).getStringCellValue());
    	Assert.assertEquals("Receipt No", row.getCell(11).getStringCellValue());
    	Assert.assertEquals("Bank No", row.getCell(12).getStringCellValue());
    	Assert.assertEquals("Lookup Client", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("Lookup Account", row.getCell(LOOKUP_ACCOUNT_NO_COL).getStringCellValue());
    	Assert.assertEquals("Lookup Product", row.getCell(LOOKUP_PRODUCT_COL).getStringCellValue());
    	Assert.assertEquals("Lookup Principal", row.getCell(LOOKUP_PRINCIPAL_COL).getStringCellValue());
    	
    	//Lookup Table test
    	row = loanRepaymentSheet.getRow(1);
    	Assert.assertEquals("Billy T Bob", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("3.0", ((Double)row.getCell(LOOKUP_ACCOUNT_NO_COL).getNumericCellValue()).toString());
    	Assert.assertEquals("HM", row.getCell(LOOKUP_PRODUCT_COL).getStringCellValue());
    	Assert.assertEquals("25000.0", ((Double)row.getCell(LOOKUP_PRINCIPAL_COL).getNumericCellValue()).toString());
    }
}
