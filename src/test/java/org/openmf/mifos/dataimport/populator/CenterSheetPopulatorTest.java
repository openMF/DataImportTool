package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

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
import org.openmf.mifos.dataimport.dto.client.CompactCenter;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;

@RunWith(MockitoJUnitRunner.class)
public class CenterSheetPopulatorTest {
	
	CenterSheetPopulator populator;
	
	@Mock
	RestClient restClient;

	//private String[] officeNameToBeginEndIndexesOfCenters;
    
    @Test
    public void shouldDownloadAndParseCenters() {
    	
    	Mockito.when(restClient.get("centers?limit=-1")).thenReturn("{\"totalFilteredRecords\": 1,\"pageItems\": [{\"id\": 1, \"name\": \"Center 1\", \"externalId\":" +
    			" \"B1561\", \"status\": {\"id\": 300, \"code\": \"centerStatusType.active\", \"value\": \"Active\"},\"active\": true,\"activationDate\":" +
    			" [2013,9,1], \"officeId\": 1, \"officeName\": \"Head Office\", \"staffId\": 1, \"staffName\": \"Chatta, Sahil\", \"hierarchy\": \".1.\"}]}");
    	
    	Mockito.when(restClient.get("offices?limit=-1")).thenReturn("[{\"id\":1,\"name\":\"Head Office\",\"nameDecorated\":\"Head Office\",\"externalId\": \"1\"," +
        		"\"openingDate\":[2009,1,1],\"hierarchy\": \".\"},{\"id\": 2,\"name\": \"Office1\",\"nameDecorated\": \"....Office1\",\"openingDate\":[2013,4,1]," +
        		"\"hierarchy\": \".2.\",\"parentId\": 1,\"parentName\": \"Head Office\"}]");
    	
    	populator = new CenterSheetPopulator(restClient);
    	Result result = populator.downloadAndParse();
    	
    	Assert.assertTrue(result.isSuccess());
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("centers?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	
    	List<CompactCenter> centers = populator.getCenters();
    	CompactCenter center = centers.get(0);
    	Assert.assertEquals(1, centers.size());
    	
    	Assert.assertEquals("1", center.getId().toString());
    	Assert.assertEquals("Center 1", center.getName());
    	Assert.assertEquals("Head Office", center.getOfficeName());
    	Assert.assertEquals("2013", center.getActivationDate().get(0).toString());
    	Assert.assertEquals("9", center.getActivationDate().get(1).toString());
    	Assert.assertEquals("1", center.getActivationDate().get(2).toString());
    }
    
    @Test
    public void shouldPopulateCenterSheet() {
    	
    	Mockito.when(restClient.get("centers?limit=-1")).thenReturn("{\"totalFilteredRecords\": 1,\"pageItems\": [{\"id\": 1, \"name\": \"Center 1\", \"externalId\":" +
    			" \"B1561\", \"status\": {\"id\": 300, \"code\": \"centerStatusType.active\", \"value\": \"Active\"},\"active\": true,\"activationDate\":" +
    			" [2013,9,1], \"officeId\": 1, \"officeName\": \"Head Office\", \"staffId\": 1, \"staffName\": \"Chatta, Sahil\", \"hierarchy\": \".1.\"}]}");
    	
    	Mockito.when(restClient.get("offices?limit=-1")).thenReturn("[{\"id\":1,\"name\":\"Head Office\",\"nameDecorated\":\"Head Office\",\"externalId\": \"1\"," +
        		"\"openingDate\":[2009,1,1],\"hierarchy\": \".\"},{\"id\": 2,\"name\": \"Office1\",\"nameDecorated\": \"....Office1\",\"openingDate\":[2013,4,1]," +
        		"\"hierarchy\": \".2.\",\"parentId\": 1,\"parentName\": \"Head Office\"}]");
    	
    	populator = new CenterSheetPopulator(restClient);
     	populator.downloadAndParse();
    	Workbook book = new HSSFWorkbook();
    	Result result = populator.populate(book);
    	Integer[] officeNameToBeginEndIndexesOfCenters = populator.getOfficeNameToBeginEndIndexesOfCenters().get(0);
    	Map<String, ArrayList<String>> officeToCenters = populator.getOfficeToCenters();
    	Map<String, Integer> centerNameToCenterId = populator.getCenterNameToCenterId();
    	
    	Assert.assertTrue(result.isSuccess());
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("centers?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	
    	Sheet centerSheet = book.getSheet("Centers");
    	Row row = centerSheet.getRow(1);
    	Assert.assertEquals("Head_Office", row.getCell(0).getStringCellValue());
    	Assert.assertEquals("Center 1", row.getCell(1).getStringCellValue());
    	Assert.assertEquals("1.0", "" + row.getCell(2).getNumericCellValue());
    	
    	Assert.assertEquals("2", "" + officeNameToBeginEndIndexesOfCenters[0]);
    	Assert.assertEquals("2", "" + officeNameToBeginEndIndexesOfCenters[1]);
    	Assert.assertEquals("1", "" + officeToCenters.size());
    	Assert.assertEquals("1", "" + officeToCenters.get("Head_Office").size());
    	Assert.assertEquals("Center", "" + officeToCenters.get("Head_Office").get(0));
    	Assert.assertEquals("1", "" + centerNameToCenterId.size());
    	Assert.assertEquals("1", "" + centerNameToCenterId.get("Center 1"));
    }

}
