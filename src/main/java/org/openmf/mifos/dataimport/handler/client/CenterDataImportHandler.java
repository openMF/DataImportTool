package org.openmf.mifos.dataimport.handler.client;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.handler.AbstractDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.openmf.mifos.dataimport.dto.client.Center;
import org.openmf.mifos.dataimport.dto.client.Meeting;
import org.openmf.mifos.dataimport.dto.client.WeeklyMeeting;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.openmf.mifos.dataimport.utils.StringUtils;

import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class CenterDataImportHandler extends AbstractDataImportHandler   {
	private static final Logger logger = LoggerFactory.getLogger(CenterDataImportHandler.class);
	
	@SuppressWarnings("CPD-START")
	private static final int NAME_COL = 0;
    private static final int OFFICE_NAME_COL = 1;
    private static final int STAFF_NAME_COL = 2;
    private static final int EXTERNAL_ID_COL = 3;
    private static final int ACTIVE_COL = 4;
    private static final int ACTIVATION_DATE_COL = 5;
    private static final int MEETING_START_DATE_COL = 6;
    private static final int IS_REPEATING_COL = 7;
    private static final int FREQUENCY_COL = 8;
    private static final int INTERVAL_COL = 9;
    private static final int REPEATS_ON_DAY_COL = 10;
    private static final int STATUS_COL = 11;
    private static final int CENTER_ID_COL = 12;
    private static final int FAILURE_COL = 13;
    @SuppressWarnings("CPD-END")
    
    private final RestClient restClient;
    private final Workbook workbook;
    
    private List<Center> centers;
	private List<Meeting> meetings;
   
	//private Sheet centerSheet;

    public CenterDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
        centers = new ArrayList<Center>();
		meetings = new ArrayList<Meeting>();
    }
	
	@Override
	public Result parse() {
		 Result result = new Result();
	        Sheet centersSheet = workbook.getSheet("Centers");
	        Integer noOfEntries = getNumberOfRows(centersSheet, 0);
	        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
	            Row row;
	            try {
	                row = centersSheet.getRow(rowIndex);
	                if(isNotImported(row, STATUS_COL)) {
	                    centers.add(parseAsCenter(row));
	                    meetings.add(parseAsMeeting(row));
	                }
	            } catch (Exception e) {
	                logger.error("row = " + rowIndex, e);
	                result.addError("Row = " + rowIndex + " , " + e.getMessage());
	            }
	        }
	        return result;
	}

	private Center parseAsCenter(Row row) {
		String status = readAsString(STATUS_COL, row);
    	String officeName = readAsString(OFFICE_NAME_COL, row);
        String officeId = getIdByName(workbook.getSheet("Offices"), officeName).toString();
        String staffName = readAsString(STAFF_NAME_COL, row);
        String staffId = getIdByName(workbook.getSheet("Staff"), staffName).toString();
        String externalId = readAsString(EXTERNAL_ID_COL, row);
        String activationDate = readAsDate(ACTIVATION_DATE_COL, row);
        String active = readAsBoolean(ACTIVE_COL, row).toString();
        String centerName = readAsString(NAME_COL, row);
        if(StringUtils.isBlank(centerName)) {
           	throw new IllegalArgumentException("Name is blank");
        }       
        return new Center(centerName, activationDate, active, externalId, officeId, staffId, row.getRowNum(), status);
	}	

	private Meeting parseAsMeeting(Row row) {
		String meetingStartDate = readAsDate(MEETING_START_DATE_COL, row);
    	String isRepeating = readAsBoolean(IS_REPEATING_COL, row).toString();
    	String frequency = readAsString(FREQUENCY_COL, row);
    	frequency = getFrequencyId(frequency);
    	String interval = readAsString(INTERVAL_COL, row);
    	String repeatsOnDay = readAsString(REPEATS_ON_DAY_COL, row);
    	repeatsOnDay = getRepeatsOnDayId(repeatsOnDay);
    	if(meetingStartDate.equals(""))
    		return null;
    	else {
    		if(repeatsOnDay.equals(""))
    			return new Meeting(meetingStartDate, isRepeating, frequency, interval, row.getRowNum());
    		else
    			return new WeeklyMeeting(meetingStartDate, isRepeating, frequency, interval, repeatsOnDay, row.getRowNum());
    	}
	}

	@Override
    public Result upload() {
        Result result = new Result();
        Sheet centerSheet = workbook.getSheet("Centers");
        int progressLevel = 0;
        String centerId = "";
        restClient.createAuthToken();
        for (int i = 0; i < centers.size(); i++) {
        	Row row = centerSheet.getRow(centers.get(i).getRowIndex());
        	Cell errorReportCell = row.createCell(FAILURE_COL);
        	Cell statusCell = row.createCell(STATUS_COL);
            try {
                String response = "";
                String status = centers.get(i).getStatus();
                progressLevel = getProgressLevel(status);
                
                if(progressLevel == 0)
                {
                   response = uploadCenter(i);
                   centerId = getCenterId(response);
                   progressLevel = 1;
                } else 
                	  centerId = readAsInt(CENTER_ID_COL, centerSheet.getRow(centers.get(i).getRowIndex()));
                
                if(meetings.get(i) != null)
                	progressLevel = uploadCenterMeeting(centerId, i);
                
                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (RuntimeException e) {
            	System.out.println(e);
            	String message = parseStatus(e.getMessage());
            	String status = "";
            	
            	if(progressLevel == 0)
            		status = "Creation";
            	else if(progressLevel == 1)
            		status = "Meeting";
            	statusCell.setCellValue(status + " failed.");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));
                
                if(progressLevel>0)
                	row.createCell(CENTER_ID_COL).setCellValue(Integer.parseInt(centerId));
               
            	errorReportCell.setCellValue(message);
                result.addError("Row = " + centers.get(i).getRowIndex() + " ," + message);
            }
        }
        setReportHeaders(centerSheet);
        return result;
    }
    
	private int getProgressLevel(String status) {
		   
        if(status.equals("") || status.equals("Creation failed."))
        	return 0;
        else if(status.equals("Meeting failed."))
        	return 1;
        return 0;
    }
	
	private String uploadCenter(int rowIndex) {
        String payload = new Gson().toJson(centers.get(rowIndex));
        logger.info(payload);
        String response = restClient.post("centers", payload);
    	return response;
    }
	
	private String getCenterId(String response) {
		JsonParser parser = new JsonParser();
        JsonObject obj = parser.parse(response).getAsJsonObject();
        return obj.get("groupId").getAsString();
    }	

	private Integer uploadCenterMeeting(String centerId, int rowIndex) {
		Meeting meeting = meetings.get(rowIndex);
    	meeting.setCenterId(centerId);
    	meeting.setTitle("centers_" + centerId + "_CollectionMeeting");
    	String payload = new Gson().toJson(meeting);
        logger.info(payload);
        restClient.post("centers/" + centerId + "/calendars", payload);
       
        return 2;
	}
	
	private void setReportHeaders(Sheet sheet) {
    	writeString(STATUS_COL, sheet.getRow(0), "Status");
    	writeString(CENTER_ID_COL, sheet.getRow(0), "Center Id");
    	writeString(FAILURE_COL, sheet.getRow(0), "Failure Report");
	}

	private String getFrequencyId(String frequency) {
		if(frequency.equalsIgnoreCase("Daily"))
    		frequency = "1";
        else if(frequency.equalsIgnoreCase("Weekly"))
        	frequency = "2";
        else if(frequency.equalsIgnoreCase("Monthly"))
        	frequency = "3";
        else if(frequency.equalsIgnoreCase("Yearly"))
        	frequency = "4";
    	return frequency;
	}	    

	private String getRepeatsOnDayId(String repeatsOnDay) {
		if(repeatsOnDay.equalsIgnoreCase("Mon"))
    		repeatsOnDay = "1";
        else if(repeatsOnDay.equalsIgnoreCase("Tue"))
        	repeatsOnDay = "2";
        else if(repeatsOnDay.equalsIgnoreCase("Wed"))
        	repeatsOnDay = "3";
        else if(repeatsOnDay.equalsIgnoreCase("Thu"))
        	repeatsOnDay = "4";
        else if(repeatsOnDay.equalsIgnoreCase("Fri"))
        	repeatsOnDay = "5";
		
        else if(repeatsOnDay.equalsIgnoreCase("Sat"))
        	repeatsOnDay = "6";
        else if(repeatsOnDay.equalsIgnoreCase("Sun"))
        	repeatsOnDay = "7";
    	return repeatsOnDay;
	}

	 public List<Center> getCenters() {
	        return centers;
     }
	    
	 public List<Meeting> getMeetings() {
	        return meetings;
	 }

}
