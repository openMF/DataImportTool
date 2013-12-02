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
import org.openmf.mifos.dataimport.populator.client.GroupWorkbookPopulator;

@RunWith(MockitoJUnitRunner.class)
public class GroupWorkbookPopulatorTest {

	@Mock
	RestClient restClient;
    
	private static final int LOOKUP_OFFICE_NAME_COL = 251;
    private static final int LOOKUP_OFFICE_OPENING_DATE_COL = 252;
    
    @Test
    public void shouldPopulateGroupWorkbook() {
    	Mockito.when(restClient.get("offices?limit=-1")).thenReturn("[{\"id\":1,\"name\":\"Head Office\",\"nameDecorated\":\"Head Office\",\"externalId\": \"1\"," +
        		"\"openingDate\":[2009,1,1],\"hierarchy\": \".\"},{\"id\": 2,\"name\": \"Office1\",\"nameDecorated\": \"....Office1\",\"openingDate\":[2013,4,1]," +
        		"\"hierarchy\": \".2.\",\"parentId\": 1,\"parentName\": \"Head Office\"}]");
    	Mockito.when(restClient.get("staff?limit=-1")).thenReturn("[{\"id\": 1, \"firstname\": \"Sahil\", \"lastname\": \"Chatta\", \"displayName\": \"Chatta, Sahil\"," +
        		" \"officeId\": 1,\"officeName\": \"Head Office\", \"isLoanOfficer\": true },{\"id\": 2, \"firstname\": \"Edin\", \"lastname\": \"Dzeko\",\"displayName\":" +
        		" \"Dzeko, Edin\",\"officeId\": 2,\"officeName\": \"Office1\",\"isLoanOfficer\": true}]");
    	Mockito.when(restClient.get("clients?limit=-1")).thenReturn("{\"totalFilteredRecords\": 2,\"pageItems\": [{\"id\": 1,\"accountNo\": \"000000001\"," +
    	 		"\"status\": {\"id\": 300,\"code\": \"clientStatusType.active\",\"value\": \"Active\"},\"active\": true,\"activationDate\": [2013,7,1]," +
    	 		"\"firstname\": \"Arsene\",\"middlename\": \"K\",\"lastname\": \"Wenger\",\"displayName\": \"Arsene K Wenger\",\"officeId\": 1," +
    	 		"\"officeName\": \"Head Office\",\"staffId\": 1,\"staffName\": \"Chatta, Sahil\"},{\"id\": 2,\"accountNo\": \"000000002\"," +
    	 		"\"status\": {\"id\": 300,\"code\": \"clientStatusType.active\",\"value\": \"Active\"},\"active\": true,\"activationDate\": [2013,7,1]," +
    	 		"\"firstname\": \"Billy\",\"middlename\": \"T\",\"lastname\": \"Bob\",\"displayName\": \"Billy T Bob\",\"officeId\": 2,\"officeName\": \"Office1\"," +
    	 		"\"staffId\": 2,\"staffName\": \"Dzeko, Edin\"}]}");
    	
    	Boolean onlyLoanOfficers = Boolean.FALSE;
    	GroupWorkbookPopulator groupWorkbookPopulator = new GroupWorkbookPopulator(new OfficeSheetPopulator(restClient),
    			new PersonnelSheetPopulator(onlyLoanOfficers, restClient), new ClientSheetPopulator(restClient));
    	groupWorkbookPopulator.downloadAndParse();
    	Workbook groupWorkbook = new HSSFWorkbook();
    	Result result = groupWorkbookPopulator.populate(groupWorkbook);
    	Assert.assertTrue(result.isSuccess());
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("staff?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("clients?limit=-1");
    	
    	Sheet groupSheet = groupWorkbook.getSheet("Groups");
    	Row row = groupSheet.getRow(0);
    	
    	//If test fails, also check if column letters embedded in formulas in setDefault and setRules have changed or not.
    	Assert.assertEquals("Group Name*", row.getCell(0).getStringCellValue());
    	Assert.assertEquals("Office Name*", row.getCell(1).getStringCellValue());
    	Assert.assertEquals("Staff Name*", row.getCell(2).getStringCellValue());
    	Assert.assertEquals("External ID", row.getCell(3).getStringCellValue());
    	Assert.assertEquals("Active*", row.getCell(4).getStringCellValue());
    	Assert.assertEquals("Activation Date*", row.getCell(5).getStringCellValue());
    	Assert.assertEquals("Meeting Start Date* (On or After)", row.getCell(6).getStringCellValue());
    	Assert.assertEquals("Repeat*", row.getCell(7).getStringCellValue());
    	Assert.assertEquals("Repeats*", row.getCell(8).getStringCellValue());
    	Assert.assertEquals("Every*", row.getCell(9).getStringCellValue());
    	Assert.assertEquals("Repeats On*", row.getCell(10).getStringCellValue());
    	Assert.assertEquals("Client Names* (Enter in consecutive cells horizontally)", row.getCell(14).getStringCellValue());
    	Assert.assertEquals("Office Name", row.getCell(LOOKUP_OFFICE_NAME_COL).getStringCellValue());
    	Assert.assertEquals("Opening Date", row.getCell(LOOKUP_OFFICE_OPENING_DATE_COL).getStringCellValue());
    	
    	//Date Lookup Table test
    	DateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy");
    	row = groupSheet.getRow(1);
    	Assert.assertEquals("Head_Office", row.getCell(LOOKUP_OFFICE_NAME_COL).getStringCellValue());
    	Assert.assertEquals("01 January 2009", dateFormat.format(row.getCell(LOOKUP_OFFICE_OPENING_DATE_COL).getDateCellValue()));
    	row = groupSheet.getRow(2);
    	Assert.assertEquals("Office1", row.getCell(LOOKUP_OFFICE_NAME_COL).getStringCellValue());
    	Assert.assertEquals("01 April 2013", dateFormat.format(row.getCell(LOOKUP_OFFICE_OPENING_DATE_COL).getDateCellValue()));
    }
}
