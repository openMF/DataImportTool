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
import org.openmf.mifos.dataimport.populator.savings.SavingsTransactionWorkbookPopulator;

@RunWith(MockitoJUnitRunner.class)
public class SavingsTransactionWorkbookPopulatorTest {
	
	@Mock
	RestClient restClient;
	
	private static final int LOOKUP_CLIENT_NAME_COL = 15;
    private static final int LOOKUP_ACCOUNT_NO_COL = 16;
    private static final int LOOKUP_PRODUCT_COL = 17;
    private static final int LOOKUP_OPENING_BALANCE_COL = 18;
	
	@Test
    public void shouldPopulateLoanRepaymentWorkbook() {
		
		Mockito.when(restClient.get("savingsaccounts?limit=-1")).thenReturn("{\"totalFilteredRecords\": 1,\"pageItems\": [{" +
      "\"id\": 6,\"accountNo\": \"000000006\",\"clientId\": 1,\"clientName\": \"Arsene K Wenger\",\"savingsProductId\": 1,\"savingsProductName\": \"SP1\",\"fieldOfficerId\": 1," +
      "\"fieldOfficerName\": \"Chatta, Sahil\",\"status\": {\"id\": 300,\"code\": \"savingsAccountStatusType.active\",\"value\": \"Active\",\"submittedAndPendingApproval\": false," +
      "\"approved\": false,\"rejected\": false,\"withdrawnByApplicant\": false,\"active\": true,\"closed\": false},\"timeline\": {\"submittedOnDate\": [2013,7,2]," +
      "\"approvedOnDate\": [2013,7,3],\"approvedByUsername\": \"mifos\",\"approvedByFirstname\": \"App\",\"approvedByLastname\": \"Administrator\",\"activatedOnDate\": [2013,7,4]}," +
      "\"currency\": {\"code\": \"USD\",\"name\": \"US Dollar\",\"decimalPlaces\": 2,\"displaySymbol\": \"$\",\"nameCode\": \"currency.USD\",\"displayLabel\": \"US Dollar ($)\"}," +
      "\"nominalAnnualInterestRate\": 8.000000,\"interestCompoundingPeriodType\": {\"id\": 1,\"code\": \"savings.interest.period.savingsCompoundingInterestPeriodType.daily\"," +
      "\"value\": \"Daily\"},\"interestPostingPeriodType\": {\"id\": 4,\"code\": \"savings.interest.posting.period.savingsPostingInterestPeriodType.monthly\",\"value\": \"Monthly\"}," +
      "\"interestCalculationType\": {\"id\": 1,\"code\": \"savingsInterestCalculationType.dailybalance\",\"value\": \"Daily Balance\"},\"interestCalculationDaysInYearType\": {" +
      "\"id\": 365,\"code\": \"savingsInterestCalculationDaysInYearType.days365\",\"value\": \"365 Days\"},\"minRequiredOpeningBalance\": 1200.000000,\"withdrawalFeeForTransfers\": true," +
      "\"summary\": {\"currency\": {\"code\": \"USD\",\"name\": \"US Dollar\",\"decimalPlaces\": 2,\"displaySymbol\": \"$\",\"nameCode\": \"currency.USD\",\"displayLabel\": \"US Dollar ($)\"}," +
      "\"totalDeposits\": 1300.000000,\"totalWithdrawals\": 200.000000,\"totalInterestEarned\": 19.830000,\"totalInterestPosted\": 15.910000,\"accountBalance\": 1115.910000}}]}");
		
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

		SavingsTransactionWorkbookPopulator savingsTransactionWorkbookPopulator = new SavingsTransactionWorkbookPopulator(restClient,
				new OfficeSheetPopulator(restClient), new ClientSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
	
		savingsTransactionWorkbookPopulator.downloadAndParse();
    	
    	Workbook savingsTransactionWorkbook = new HSSFWorkbook();
    	Result result = savingsTransactionWorkbookPopulator.populate(savingsTransactionWorkbook);
    	Assert.assertTrue(result.isSuccess());
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("savingsaccounts?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("clients?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("funds");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("codes/12/codevalues");
    	
    	Sheet savingsTransactionSheet = savingsTransactionWorkbook.getSheet("SavingsTransaction");
    	Row row = savingsTransactionSheet.getRow(0);
    	
    	//If test fails, also check if column letters embedded in formulas in setDefault and setRules have changed or not.
    	Assert.assertEquals("Office Name*", row.getCell(0).getStringCellValue());
    	Assert.assertEquals("Client Name*", row.getCell(1).getStringCellValue());
    	Assert.assertEquals("Account No.*", row.getCell(2).getStringCellValue());
    	Assert.assertEquals("Product Name", row.getCell(3).getStringCellValue());
    	Assert.assertEquals("Opening Balance", row.getCell(4).getStringCellValue());
    	Assert.assertEquals("Transaction Type*", row.getCell(5).getStringCellValue());
    	Assert.assertEquals("Amount*", row.getCell(6).getStringCellValue()); 
    	Assert.assertEquals("Date*", row.getCell(7).getStringCellValue()); 
    	Assert.assertEquals("Type*", row.getCell(8).getStringCellValue()); 
    	Assert.assertEquals("Account No", row.getCell(9).getStringCellValue());
    	Assert.assertEquals("Check No", row.getCell(10).getStringCellValue());
    	Assert.assertEquals("Routing Code", row.getCell(11).getStringCellValue());
    	Assert.assertEquals("Receipt No", row.getCell(12).getStringCellValue());
    	Assert.assertEquals("Bank No", row.getCell(13).getStringCellValue());
    	Assert.assertEquals("Lookup Client", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("Lookup Account", row.getCell(LOOKUP_ACCOUNT_NO_COL).getStringCellValue());
    	Assert.assertEquals("Lookup Product", row.getCell(LOOKUP_PRODUCT_COL).getStringCellValue());
    	Assert.assertEquals("Lookup Opening Balance", row.getCell(LOOKUP_OPENING_BALANCE_COL).getStringCellValue());
    	
    	//Lookup Table test
    	row = savingsTransactionSheet.getRow(1);
    	Assert.assertEquals("Arsene K Wenger", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("6.0", ((Double)row.getCell(LOOKUP_ACCOUNT_NO_COL).getNumericCellValue()).toString());
    	Assert.assertEquals("SP1", row.getCell(LOOKUP_PRODUCT_COL).getStringCellValue());
    	Assert.assertEquals("1200.0", ((Double)row.getCell(LOOKUP_OPENING_BALANCE_COL).getNumericCellValue()).toString());
	
	}

}
