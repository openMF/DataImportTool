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
import org.openmf.mifos.dataimport.dto.client.CompactGroup;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;

@RunWith(MockitoJUnitRunner.class)
public class GroupSheetPopulatorTest {
	
	GroupSheetPopulator populator;
	
	@Mock
	RestClient restClient;
    
    @Test
    public void shouldDownloadAndParseGroups() {
    	
    	Mockito.when(restClient.get("groups?limit=-1")).thenReturn("{\"totalFilteredRecords\": 1,\"pageItems\": [{\"id\": 1, \"name\": \"Group 1\", \"externalId\":" +
    			" \"B1561\", \"status\": {\"id\": 300, \"code\": \"clientStatusType.active\", \"value\": \"Active\"},\"active\": true,\"activationDate\":" +
    			" [2013,9,1], \"officeId\": 1, \"officeName\": \"Head Office\", \"staffId\": 1, \"staffName\": \"Chatta, Sahil\", \"hierarchy\": \".1.\"}]}");
    	
    	Mockito.when(restClient.get("offices?limit=-1")).thenReturn("[{\"id\":1,\"name\":\"Head Office\",\"nameDecorated\":\"Head Office\",\"externalId\": \"1\"," +
        		"\"openingDate\":[2009,1,1],\"hierarchy\": \".\"},{\"id\": 2,\"name\": \"Office1\",\"nameDecorated\": \"....Office1\",\"openingDate\":[2013,4,1]," +
        		"\"hierarchy\": \".2.\",\"parentId\": 1,\"parentName\": \"Head Office\"}]");
    	
    	populator = new GroupSheetPopulator(restClient);
    	Result result = populator.downloadAndParse();
    	
    	Assert.assertTrue(result.isSuccess());
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("groups?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	
    	List<CompactGroup> groups = populator.getGroups();
    	CompactGroup group = groups.get(0);
    	Assert.assertEquals(1, groups.size());
    	
    	Assert.assertEquals("1", group.getId().toString());
    	Assert.assertEquals("Group 1", group.getName());
    	Assert.assertEquals("Head Office", group.getOfficeName());
    	Assert.assertEquals("2013", group.getActivationDate().get(0).toString());
    	Assert.assertEquals("9", group.getActivationDate().get(1).toString());
    	Assert.assertEquals("1", group.getActivationDate().get(2).toString());
    }
    
    @Test
    public void shouldPopulateGroupSheet() {
    	
    	Mockito.when(restClient.get("groups?limit=-1")).thenReturn("{\"totalFilteredRecords\": 1,\"pageItems\": [{\"id\": 1, \"name\": \"Group 1\", \"externalId\":" +
    			" \"B1561\", \"status\": {\"id\": 300, \"code\": \"clientStatusType.active\", \"value\": \"Active\"},\"active\": true,\"activationDate\":" +
    			" [2013,9,1], \"officeId\": 1, \"officeName\": \"Head Office\", \"staffId\": 1, \"staffName\": \"Chatta, Sahil\", \"hierarchy\": \".1.\"}]}");
    	
    	Mockito.when(restClient.get("offices?limit=-1")).thenReturn("[{\"id\":1,\"name\":\"Head Office\",\"nameDecorated\":\"Head Office\",\"externalId\": \"1\"," +
        		"\"openingDate\":[2009,1,1],\"hierarchy\": \".\"},{\"id\": 2,\"name\": \"Office1\",\"nameDecorated\": \"....Office1\",\"openingDate\":[2013,4,1]," +
        		"\"hierarchy\": \".2.\",\"parentId\": 1,\"parentName\": \"Head Office\"}]");
    	
    	populator = new GroupSheetPopulator(restClient);
     	populator.downloadAndParse();
    	Workbook book = new HSSFWorkbook();
    	Result result = populator.populate(book);
    	Integer[] officeNameToBeginEndIndexesOfGroups = populator.getOfficeNameToBeginEndIndexesOfGroups().get(0);
    	Map<String, ArrayList<String>> officeToGroups = populator.getOfficeToGroups();
    	Map<String, Integer> groupNameToGroupId = populator.getGroupNameToGroupId();
    	
    	Assert.assertTrue(result.isSuccess());
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("groups?limit=-1");
    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("offices?limit=-1");
    	
    	Sheet groupSheet = book.getSheet("Groups");
    	Row row = groupSheet.getRow(1);
    	Assert.assertEquals("Head_Office", row.getCell(0).getStringCellValue());
    	Assert.assertEquals("Group 1", row.getCell(1).getStringCellValue());
    	Assert.assertEquals("1.0", "" + row.getCell(2).getNumericCellValue());
    	
    	Assert.assertEquals("2", "" + officeNameToBeginEndIndexesOfGroups[0]);
    	Assert.assertEquals("2", "" + officeNameToBeginEndIndexesOfGroups[1]);
    	Assert.assertEquals("1", "" + officeToGroups.size());
    	Assert.assertEquals("1", "" + officeToGroups.get("Head_Office").size());
    	Assert.assertEquals("Group 1", "" + officeToGroups.get("Head_Office").get(0));
    	Assert.assertEquals("1", "" + groupNameToGroupId.size());
    	Assert.assertEquals("1", "" + groupNameToGroupId.get("Group 1"));
    }

}
