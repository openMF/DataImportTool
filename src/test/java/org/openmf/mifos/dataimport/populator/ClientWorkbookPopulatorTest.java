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
import org.openmf.mifos.dataimport.populator.client.ClientWorkbookPopulator;

@RunWith(MockitoJUnitRunner.class)
public class ClientWorkbookPopulatorTest {

    @Mock
	RestClient restClient;
    
    private static final int RELATIONAL_OFFICE_NAME_COL = 16;
    private static final int RELATIONAL_OFFICE_OPENING_DATE_COL = 17;
    
    @Test
    public void shouldPopulateClientWorkbook() {
    	
    	Mockito.when(restClient.get("offices?limit=-1")).thenReturn("[{\"id\":1,\"name\":\"Head Office\",\"nameDecorated\":\"Head Office\",\"externalId\": \"1\"," +
        		"\"openingDate\":[2009,1,1],\"hierarchy\": \".\"},{\"id\": 2,\"name\": \"Office1\",\"nameDecorated\": \"....Office1\",\"openingDate\":[2013,4,1]," +
        		"\"hierarchy\": \".2.\",\"parentId\": 1,\"parentName\": \"Head Office\"}]");
    	Mockito.when(restClient.get("staff?limit=-1")).thenReturn("[{\"id\": 1, \"firstname\": \"Sahil\", \"lastname\": \"Chatta\", \"displayName\": \"Chatta, Sahil\"," +
        		" \"officeId\": 1,\"officeName\": \"Head Office\", \"isLoanOfficer\": true },{\"id\": 2, \"firstname\": \"Edin\", \"lastname\": \"Dzeko\",\"displayName\":" +
        		" \"Dzeko, Edin\",\"officeId\": 2,\"officeName\": \"Office1\",\"isLoanOfficer\": true}]");
    	
    	Boolean onlyLoanOfficers = Boolean.FALSE;
    	ClientWorkbookPopulator individualClientWorkbookPopulator = new ClientWorkbookPopulator("individual",
    			new OfficeSheetPopulator(restClient), new PersonnelSheetPopulator(onlyLoanOfficers, restClient));
    	ClientWorkbookPopulator corporateClientWorkbookPopulator = new ClientWorkbookPopulator("corporate",
    			new OfficeSheetPopulator(restClient), new PersonnelSheetPopulator(onlyLoanOfficers, restClient));
    	individualClientWorkbookPopulator.downloadAndParse();
    	corporateClientWorkbookPopulator.downloadAndParse();
    	
    	Workbook individualClientWorkbook = new HSSFWorkbook();
    	Workbook corporateClientWorkbook = new HSSFWorkbook();
    	Result result = individualClientWorkbookPopulator.populate(individualClientWorkbook);
    	Assert.assertTrue(result.isSuccess());
    	result = corporateClientWorkbookPopulator.populate(corporateClientWorkbook);
    	Assert.assertTrue(result.isSuccess());
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("staff?limit=-1");
    	
    	Sheet individualClientSheet = individualClientWorkbook.getSheet("Clients");
    	Sheet corporateClientSheet = corporateClientWorkbook.getSheet("Clients");
    	Row row = individualClientSheet.getRow(0);
    	
    	//If test fails, also check if column letters embedded in formulas in setDefault and setRules have changed or not.
    	Assert.assertEquals("First Name*", row.getCell(0).getStringCellValue());
    	Assert.assertEquals("Last Name*", row.getCell(1).getStringCellValue());
    	Assert.assertEquals("Middle Name", row.getCell(2).getStringCellValue());
    	Assert.assertEquals("Office Name*", row.getCell(3).getStringCellValue());
    	Assert.assertEquals("Staff Name*", row.getCell(4).getStringCellValue());
    	Assert.assertEquals("External ID", row.getCell(5).getStringCellValue());
    	Assert.assertEquals("Activation Date*", row.getCell(6).getStringCellValue());
    	Assert.assertEquals("Active*", row.getCell(7).getStringCellValue());
    	Assert.assertEquals("Office Name", row.getCell(RELATIONAL_OFFICE_NAME_COL).getStringCellValue());
    	Assert.assertEquals("Opening Date", row.getCell(RELATIONAL_OFFICE_OPENING_DATE_COL).getStringCellValue());
    	
    	Assert.assertEquals("Full/Business Name*", corporateClientSheet.getRow(0).getCell(0).getStringCellValue());
    	
    	//Date Lookup Table test
    	DateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy");
    	row = individualClientSheet.getRow(1);
    	Assert.assertEquals("Head_Office", row.getCell(RELATIONAL_OFFICE_NAME_COL).getStringCellValue());
    	Assert.assertEquals("01 January 2009", dateFormat.format(row.getCell(RELATIONAL_OFFICE_OPENING_DATE_COL).getDateCellValue()));
    	row = individualClientSheet.getRow(2);
    	Assert.assertEquals("Office1", row.getCell(RELATIONAL_OFFICE_NAME_COL).getStringCellValue());
    	Assert.assertEquals("01 April 2013", dateFormat.format(row.getCell(RELATIONAL_OFFICE_OPENING_DATE_COL).getDateCellValue()));
    	
    }
}
