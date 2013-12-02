package org.openmf.mifos.dataimport.populator;

import java.text.DateFormat;
import java.text.SimpleDateFormat;

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
import org.openmf.mifos.dataimport.populator.loan.LoanProductSheetPopulator;
import org.openmf.mifos.dataimport.populator.loan.LoanWorkbookPopulator;

@RunWith(MockitoJUnitRunner.class)
public class LoanWorkbookPopulatorTest {

	@Mock
	RestClient restClient;
	
	 private static final int LOOKUP_CLIENT_NAME_COL = 42;
	 private static final int LOOKUP_ACTIVATION_DATE_COL = 43;
	
	@Test
    public void shouldPopulateLoanWorkbook() {
		
		Mockito.when(restClient.get("clients?limit=-1")).thenReturn("{\"totalFilteredRecords\": 2,\"pageItems\": [{\"id\": 1,\"accountNo\": \"000000001\"," +
    	 		"\"status\": {\"id\": 300,\"code\": \"clientStatusType.active\",\"value\": \"Active\"},\"active\": true,\"activationDate\": [2013,7,1]," +
    	 		"\"firstname\": \"Arsene\",\"middlename\": \"K\",\"lastname\": \"Wenger\",\"displayName\": \"Arsene K Wenger\",\"officeId\": 1," +
    	 		"\"officeName\": \"Head Office\",\"staffId\": 1,\"staffName\": \"Chatta, Sahil\"},{\"id\": 2,\"accountNo\": \"000000002\"," +
    	 		"\"status\": {\"id\": 300,\"code\": \"clientStatusType.active\",\"value\": \"Active\"},\"active\": true,\"activationDate\": [2013,7,1]," +
    	 		"\"firstname\": \"Billy\",\"middlename\": \"T\",\"lastname\": \"Bob\",\"displayName\": \"Billy T Bob\",\"officeId\": 2,\"officeName\": \"Office1\"," +
    	 		"\"staffId\": 2,\"staffName\": \"Dzeko, Edin\"}]}");
		
		Mockito.when(restClient.get("groups?limit=-1")).thenReturn("{\"totalFilteredRecords\": 1,\"pageItems\": [{\"id\": 1, \"name\": \"Group 1\", \"externalId\":" +
    			" \"B1561\", \"status\": {\"id\": 300, \"code\": \"clientStatusType.active\", \"value\": \"Active\"},\"active\": true,\"activationDate\":" +
    			" [2013,9,1], \"officeId\": 1, \"officeName\": \"Head Office\", \"staffId\": 1, \"staffName\": \"Chatta, Sahil\", \"hierarchy\": \".1.\"}]}");
		
		Mockito.when(restClient.get("offices?limit=-1")).thenReturn("[{\"id\":1,\"name\":\"Head Office\",\"nameDecorated\":\"Head Office\",\"externalId\": \"1\"," +
        		"\"openingDate\":[2009,1,1],\"hierarchy\": \".\"},{\"id\": 2,\"name\": \"Office1\",\"nameDecorated\": \"....Office1\",\"openingDate\":[2013,4,1]," +
        		"\"hierarchy\": \".2.\",\"parentId\": 1,\"parentName\": \"Head Office\"}]");
		
    	Mockito.when(restClient.get("staff?limit=-1")).thenReturn("[{\"id\": 1, \"firstname\": \"Sahil\", \"lastname\": \"Chatta\", \"displayName\": \"Chatta, Sahil\"," +
        		" \"officeId\": 1,\"officeName\": \"Head Office\", \"isLoanOfficer\": true },{\"id\": 2, \"firstname\": \"Edin\", \"lastname\": \"Dzeko\",\"displayName\":" +
        		" \"Dzeko, Edin\",\"officeId\": 2,\"officeName\": \"Office1\",\"isLoanOfficer\": true}]");
    	
    	Mockito.when(restClient.get("loanproducts")).thenReturn("[{\"id\": 1,\"name\": \"HM\",\"description\": \"HM\",\"fundId\": 1,\"fundName\": \"Fund1\",\"includeInBorrowerCycle\": true," +
    		    "\"startDate\":[2012,4,1],\"closeDate\":[2014,6,1],\"status\": \"loanProduct.active\",\"currency\":{\"code\": \"USD\",\"name\": \"US Dollar\",\"decimalPlaces\": 2,\"inMultiplesOf\": 5," +
    		    "\"displaySymbol\": \"$\",\"nameCode\": \"currency.USD\",\"displayLabel\": \"US Dollar ($)\"},\"principal\": 20000.000000,\"minPrincipal\": 10000.000000,\"maxPrincipal\": 30000.000000," +
    		    "\"numberOfRepayments\": 12,\"minNumberOfRepayments\": 5,\"maxNumberOfRepayments\": 24,\"repaymentEvery\": 1,\"repaymentFrequencyType\":{\"id\":2,\"code\": \"repaymentFrequency.periodFrequencyType.months\"," +
    		    "\"value\":\"Months\"},\"interestRatePerPeriod\":7.000000,\"minInterestRatePerPeriod\":5.000000,\"maxInterestRatePerPeriod\": 9.000000,\"interestRateFrequencyType\":{\"id\": 3," +
    		    "\"code\": \"interestRateFrequency.periodFrequencyType.years\",\"value\": \"Per year\"},\"annualInterestRate\": 7.000000,\"amortizationType\":{\"id\": 1,\"code\": \"amortizationType.equal.installments\"," +
    		    "\"value\": \"Equal installments\"},\"interestType\":{\"id\": 0,\"code\":\"interestType.declining.balance\",\"value\": \"Declining Balance\"},\"interestCalculationPeriodType\": {\"id\": 1," +
    		    "\"code\": \"interestCalculationPeriodType.same.as.repayment.period\",\"value\": \"Same as repayment period\"},\"inArrearsTolerance\":3.000000,\"transactionProcessingStrategyId\": 4," +
    		    "\"transactionProcessingStrategyName\": \"RBI (India)\",\"graceOnPrincipalPayment\": 1,\"graceOnInterestPayment\": 2,\"graceOnInterestCharged\": 1,\"accountingRule\": {\"id\": 1," +
    		    "\"code\": \"accountingRuleType.none\",\"value\": \"NONE\"}}]");
    	
    	Mockito.when(restClient.get("funds")).thenReturn("[{\"id\": 1,\"name\": \"Fund1\"}]");
        Mockito.when(restClient.get("codes/12/codevalues")).thenReturn("[{\"id\": 10,\"name\": \"Cash\",\"position\": 1},{\"id\": 11,\"name\": \"MPesa\",\"position\": 2}]");
	    
        Boolean onlyLoanOfficers = Boolean.TRUE;
        LoanWorkbookPopulator loanWorkbookPopulator = new LoanWorkbookPopulator(new OfficeSheetPopulator(restClient),
        		new ClientSheetPopulator(restClient), new GroupSheetPopulator(restClient), new PersonnelSheetPopulator(onlyLoanOfficers, restClient),
   			 new LoanProductSheetPopulator(restClient), new ExtrasSheetPopulator(restClient));
        loanWorkbookPopulator.downloadAndParse();
        Workbook loanWorkbook = new HSSFWorkbook();
        Result result = loanWorkbookPopulator.populate(loanWorkbook);
        Assert.assertTrue(result.isSuccess());
        Mockito.verify(restClient, Mockito.atLeastOnce()).get("clients?limit=-1");
        Mockito.verify(restClient, Mockito.atLeastOnce()).get("groups?limit=-1");
        Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("staff?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("loanproducts");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("funds");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("codes/12/codevalues");
	
    	Sheet loanSheet = loanWorkbook.getSheet("Loans");
    	Row row = loanSheet.getRow(0);
    	
    	//If test fails, also check if column letters embedded in formulas in setDefault and setRules have changed or not.
    	Assert.assertEquals("Office Name*", row.getCell(0).getStringCellValue());
    	Assert.assertEquals("Client/Group Name*", row.getCell(2).getStringCellValue());
    	Assert.assertEquals("Product*", row.getCell(3).getStringCellValue());
    	Assert.assertEquals("Loan Officer*", row.getCell(4).getStringCellValue());
    	Assert.assertEquals("Submitted On*", row.getCell(5).getStringCellValue());
    	Assert.assertEquals("Approved On*", row.getCell(6).getStringCellValue());
    	Assert.assertEquals("Disbursed Date*", row.getCell(7).getStringCellValue());
    	Assert.assertEquals("Payment Type*", row.getCell(8).getStringCellValue());
    	Assert.assertEquals("Fund Name", row.getCell(9).getStringCellValue());
    	Assert.assertEquals("Principal*", row.getCell(10).getStringCellValue());
    	Assert.assertEquals("# of Repayments*", row.getCell(11).getStringCellValue());
    	Assert.assertEquals("Repaid Every*", row.getCell(12).getStringCellValue());
    	Assert.assertEquals("Loan Term*", row.getCell(14).getStringCellValue());
    	Assert.assertEquals("Nominal Interest %*", row.getCell(16).getStringCellValue());
    	Assert.assertEquals("Amortization*", row.getCell(18).getStringCellValue());
    	Assert.assertEquals("Interest Method*", row.getCell(19).getStringCellValue());
    	Assert.assertEquals("Interest Calculation Period*", row.getCell(20).getStringCellValue());
    	Assert.assertEquals("Arrears Tolerance", row.getCell(21).getStringCellValue());
    	Assert.assertEquals("Repayment Strategy*", row.getCell(22).getStringCellValue());
    	Assert.assertEquals("Grace-Principal Payment", row.getCell(23).getStringCellValue());
    	Assert.assertEquals("Grace-Interest Payment", row.getCell(24).getStringCellValue());
    	Assert.assertEquals("Interest-Free Period(s)", row.getCell(25).getStringCellValue());
    	Assert.assertEquals("Interest Charged From", row.getCell(26).getStringCellValue());
    	Assert.assertEquals("First Repayment On", row.getCell(27).getStringCellValue());
    	Assert.assertEquals("Amount Repaid", row.getCell(28).getStringCellValue());
    	Assert.assertEquals("Date-Last Repayment", row.getCell(29).getStringCellValue());
    	Assert.assertEquals("Repayment Type", row.getCell(30).getStringCellValue());
    	Assert.assertEquals("Client Name", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("Client Activation Date", row.getCell(LOOKUP_ACTIVATION_DATE_COL).getStringCellValue());
    	
    	//Date Lookup Table test
    	DateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy");
    	row = loanSheet.getRow(1);
    	Assert.assertEquals("Arsene K Wenger", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("01 July 2013", dateFormat.format(row.getCell(LOOKUP_ACTIVATION_DATE_COL).getDateCellValue()));
    	row = loanSheet.getRow(2);
    	Assert.assertEquals("Billy T Bob", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("01 July 2013", dateFormat.format(row.getCell(LOOKUP_ACTIVATION_DATE_COL).getDateCellValue()));
	
	}
}
