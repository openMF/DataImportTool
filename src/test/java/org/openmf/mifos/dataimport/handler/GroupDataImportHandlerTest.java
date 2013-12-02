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
import org.openmf.mifos.dataimport.dto.client.Group;
import org.openmf.mifos.dataimport.dto.client.WeeklyMeeting;
import org.openmf.mifos.dataimport.handler.client.GroupDataImportHandler;
import org.openmf.mifos.dataimport.http.RestClient;

@RunWith(MockitoJUnitRunner.class)
public class GroupDataImportHandlerTest {

	    @Mock
	    RestClient restClient;
	    
	    @Test
	    public void shouldParseGroupsAndMeetings() throws IOException {
	    	InputStream is = this.getClass().getClassLoader().getResourceAsStream("client/groups.xls");
	        Workbook book = new HSSFWorkbook(is);
	        GroupDataImportHandler handler = new GroupDataImportHandler(book, restClient);
	        Result result = handler.parse();
	        Assert.assertTrue(result.isSuccess());
	        Assert.assertEquals(1, handler.getGroups().size());
	        Group group = handler.getGroups().get(0);
	        WeeklyMeeting meeting = (WeeklyMeeting)handler.getMeetings().get(0);
	        Assert.assertEquals("Test Group X", group.getName());
	        Assert.assertEquals("1", group.getOfficeId());
	        Assert.assertEquals("1", handler.getIdByName(book.getSheet("Offices"), "Head_Office").toString());
	        Assert.assertEquals("1", group.getStaffId());
	        Assert.assertEquals("1", handler.getIdByName(book.getSheet("Staff"), "Sahil Chatta").toString());
	        Assert.assertEquals("4531",group.getExternalId());
	        Assert.assertEquals("true", group.isActive());
	        Assert.assertEquals("13 September 2013", group.getActivationDate());
	        Assert.assertEquals("1", group.getClientMembers().get(0));
	        Assert.assertEquals("1", handler.getIdByName(book.getSheet("Clients"), "Arsene K Wenger").toString());
	        
	        Assert.assertEquals("14 September 2013", meeting.getStartDate());
	        Assert.assertEquals("true", meeting.isRepeating());
	        Assert.assertEquals("Weekly", meeting.getRepeats());
	        Assert.assertEquals("2", meeting.getRepeatsEvery());
	        Assert.assertEquals("WE", meeting.getRepeatsOnDay());
	    }
}
