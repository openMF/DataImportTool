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
import org.openmf.mifos.dataimport.populator.savings.SavingsProductSheetPopulator;
import org.openmf.mifos.dataimport.populator.savings.SavingsWorkbookPopulator;

@RunWith(MockitoJUnitRunner.class)
public class SavingsWorkbookPopulatorTest {

	@Mock
	RestClient restClient;
	
	private static final int LOOKUP_CLIENT_NAME_COL = 31;
    private static final int LOOKUP_ACTIVATION_DATE_COL = 32;
	
	@Test
    public void shouldPopulateSavingsWorkbook() {
		
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
    	
    	Mockito.when(restClient.get("savingsproducts")).thenReturn("[{\"id\": 2,\"name\": \"SP2\",\"description\": \"SP2\",\"currency\": {\"code\": \"USD\",\"name\": \"US Dollar\",\"decimalPlaces\": 2," +
    		      "\"inMultiplesOf\": 5,\"displaySymbol\": \"$\",\"nameCode\": \"currency.USD\",\"displayLabel\": \"US Dollar ($)\"},\"nominalAnnualInterestRate\": 10.000000,\"interestCompoundingPeriodType\": {" +
    		      "\"id\": 1,\"code\": \"savings.interest.period.savingsCompoundingInterestPeriodType.daily\",\"value\": \"Daily\"},\"interestPostingPeriodType\": {\"id\": 4,\"code\": \"savings.interest.posting.period.savingsPostingInterestPeriodType.monthly\"," +
    		      "\"value\": \"Monthly\"},\"interestCalculationType\": {\"id\": 1,\"code\": \"savingsInterestCalculationType.dailybalance\",\"value\": \"Daily Balance\"},\"interestCalculationDaysInYearType\": {" +
    		      "\"id\": 365,\"code\": \"savingsInterestCalculationDaysInYearType.days365\",\"value\": \"365 Days\"},\"minRequiredOpeningBalance\": 870.000000,\"lockinPeriodFrequency\": 1,\"lockinPeriodFrequencyType\": {" +
    		      "\"id\": 0,\"code\": \"savings.lockin.savingsPeriodFrequencyType.days\",\"value\": \"Days\"},\"withdrawalFeeAmount\": 1.000000,\"withdrawalFeeType\": {\"id\": 1,\"code\": \"savingsWithdrawalFeesType.flat\"," +
    		      "\"value\": \"Flat\"},\"annualFeeAmount\": 3.000000,\"annualFeeOnMonthDay\": [9,1],\"accountingRule\": {\"id\": 1,\"code\": \"accountingRuleType.none\",\"value\": \"NONE\"}}]");
		
		Boolean onlyLoanOfficers = Boolean.TRUE;
		SavingsWorkbookPopulator savingsWorkbookPopulator = new SavingsWorkbookPopulator(new OfficeSheetPopulator(restClient),
				new ClientSheetPopulator(restClient), new GroupSheetPopulator(restClient),new PersonnelSheetPopulator(onlyLoanOfficers, restClient),
				new SavingsProductSheetPopulator(restClient));
		savingsWorkbookPopulator.downloadAndParse();
		Workbook savingsWorkbook = new HSSFWorkbook();
		Result result = savingsWorkbookPopulator.populate(savingsWorkbook);
		Assert.assertTrue(result.isSuccess());
		Mockito.verify(restClient, Mockito.atLeastOnce()).get("clients?limit=-1");
		Mockito.verify(restClient, Mockito.atLeastOnce()).get("groups?limit=-1");
		Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("staff?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("savingsproducts");
    	
    	Sheet savingsSheet = savingsWorkbook.getSheet("Savings");
    	Row row = savingsSheet.getRow(0);
    	
    	//If test fails, also check if column letters embedded in formulas in setDefault and setRules have changed or not.
    	Assert.assertEquals("Office Name*", row.getCell(0).getStringCellValue());
    	Assert.assertEquals("Client Name*", row.getCell(2).getStringCellValue());
    	Assert.assertEquals("Product*", row.getCell(3).getStringCellValue());
    	Assert.assertEquals("Field Officer*", row.getCell(4).getStringCellValue());
    	Assert.assertEquals("Submitted On*", row.getCell(5).getStringCellValue());
    	Assert.assertEquals("Approved On*", row.getCell(6).getStringCellValue());
    	Assert.assertEquals("Activation Date*", row.getCell(7).getStringCellValue());
    	Assert.assertEquals("Currency", row.getCell(8).getStringCellValue());
    	Assert.assertEquals("Decimal Places", row.getCell(9).getStringCellValue());
    	Assert.assertEquals("In Multiples Of", row.getCell(10).getStringCellValue());
    	Assert.assertEquals("Interest Rate %*", row.getCell(11).getStringCellValue());
    	Assert.assertEquals("Interest Compounding Period*", row.getCell(12).getStringCellValue());
    	Assert.assertEquals("Interest Posting Period*", row.getCell(13).getStringCellValue());
    	Assert.assertEquals("Interest Calculated*", row.getCell(14).getStringCellValue());
    	Assert.assertEquals("# Days in Year*", row.getCell(15).getStringCellValue());
    	Assert.assertEquals("Min Opening Balance", row.getCell(16).getStringCellValue());
    	Assert.assertEquals("Locked In For", row.getCell(17).getStringCellValue());
    	Assert.assertEquals("Withdrawal Fee", row.getCell(19).getStringCellValue());
    	Assert.assertEquals("Annual Fee", row.getCell(21).getStringCellValue());
    	Assert.assertEquals("On Date", row.getCell(22).getStringCellValue());
    	Assert.assertEquals("Apply Withdrawal Fee For Transfers", row.getCell(23).getStringCellValue());
    	Assert.assertEquals("Client Name", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("Client Activation Date", row.getCell(LOOKUP_ACTIVATION_DATE_COL).getStringCellValue());
    	
    	//Date Lookup Table test
    	DateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy");
    	row = savingsSheet.getRow(1);
    	Assert.assertEquals("Arsene K Wenger", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("01 July 2013", dateFormat.format(row.getCell(LOOKUP_ACTIVATION_DATE_COL).getDateCellValue()));
    	row = savingsSheet.getRow(2);
    	Assert.assertEquals("Billy T Bob", row.getCell(LOOKUP_CLIENT_NAME_COL).getStringCellValue());
    	Assert.assertEquals("01 July 2013", dateFormat.format(row.getCell(LOOKUP_ACTIVATION_DATE_COL).getDateCellValue()));
	}
}
