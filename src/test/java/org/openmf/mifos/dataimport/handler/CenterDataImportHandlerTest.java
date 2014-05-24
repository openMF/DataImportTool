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
import org.openmf.mifos.dataimport.dto.client.Center;
import org.openmf.mifos.dataimport.dto.client.WeeklyMeeting;
import org.openmf.mifos.dataimport.handler.client.CenterDataImportHandler;
import org.openmf.mifos.dataimport.http.RestClient;

@RunWith(MockitoJUnitRunner.class)
public class CenterDataImportHandlerTest {

		@Mock
	    RestClient restClient;
	    
	    @Test
	    public void shouldParsecentersAndMeetings() throws IOException {
	    	InputStream is = this.getClass().getClassLoader().getResourceAsStream("client/centers.xls");
	        Workbook book = new HSSFWorkbook(is);
	        CenterDataImportHandler handler = new CenterDataImportHandler(book, restClient);
	        Result result = handler.parse();
	        Assert.assertTrue(result.isSuccess());
	        Assert.assertEquals(1, handler.getCenters().size());
	        Center center = handler.getCenters().get(0);
	        WeeklyMeeting meeting = (WeeklyMeeting)handler.getMeetings().get(0);
	        Assert.assertEquals("Test center X", center.getName());
	        Assert.assertEquals("1", center.getOfficeId());
	        Assert.assertEquals("1", handler.getIdByName(book.getSheet("Offices"), "Head_Office").toString());
	        Assert.assertEquals("1", center.getStaffId());
	        Assert.assertEquals("1", handler.getIdByName(book.getSheet("Staff"), "Sahil Chatta").toString());
	        Assert.assertEquals("4531",center.getExternalId());
	        Assert.assertEquals("true", center.isActive());
	        Assert.assertEquals("13 September 2013", center.getActivationDate());
	        Assert.assertEquals("1", handler.getIdByName(book.getSheet("Clients"), "Arsene K Wenger").toString());
	        
	        Assert.assertEquals("14 September 2013", meeting.getStartDate());
	        Assert.assertEquals("true", meeting.isRepeating());
	        Assert.assertEquals("2", meeting.getFrequency());
	        Assert.assertEquals("2", meeting.getInterval());
	        Assert.assertEquals("3", meeting.getRepeatsOnDay());
	    }
}
