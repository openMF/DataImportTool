package org.openmf.mifos.dataimport.handler.client;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.client.Group;
import org.openmf.mifos.dataimport.dto.client.Meeting;
import org.openmf.mifos.dataimport.dto.client.WeeklyMeeting;
import org.openmf.mifos.dataimport.handler.AbstractDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.openmf.mifos.dataimport.utils.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class GroupDataImportHandler extends AbstractDataImportHandler {

	private static final Logger logger = LoggerFactory.getLogger(GroupDataImportHandler.class);
	
	@SuppressWarnings("CPD-START")
	private static final int NAME_COL = 0;
    private static final int OFFICE_NAME_COL = 1;
    private static final int STAFF_NAME_COL = 2;
    private static final int EXTERNAL_ID_COL = 3;
    private static final int ACTIVE_COL = 4;
    private static final int ACTIVATION_DATE_COL = 5;
    private static final int MEETING_START_DATE_COL = 6;
    private static final int IS_REPEATING_COL = 7;
    private static final int REPEATS_COL = 8;
    private static final int REPEATS_EVERY_COL = 9;
    private static final int REPEATS_ON_DAY_COL = 10;
    private static final int STATUS_COL = 11;
    private static final int GROUP_ID_COL = 12;
    private static final int FAILURE_COL = 13;
    private static final int CLIENT_NAMES_STARTING_COL = 14;
    private static final int CLIENT_NAMES_ENDING_COL = 109;
    @SuppressWarnings("CPD-END")
    
    private final RestClient restClient;
    
    private final Workbook workbook;
    
    private List<Group> groups;
    private List<Meeting> meetings;

    public GroupDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
        groups = new ArrayList<Group>();
        meetings = new ArrayList<Meeting>();
    }
    
    @Override
    public Result parse() {
        Result result = new Result();
        Sheet groupsSheet = workbook.getSheet("Groups");
        Integer noOfEntries = getNumberOfRows(groupsSheet, 0);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = groupsSheet.getRow(rowIndex);
                if(isNotImported(row, STATUS_COL)) {
                    groups.add(parseAsGroup(row));
                    meetings.add(parseAsMeeting(row));
                }
            } catch (Exception e) {
                logger.error("row = " + rowIndex, e);
                result.addError("Row = " + rowIndex + " , " + e.getMessage());
            }
        }
        return result;
    }

    private Group parseAsGroup(Row row) {
    	String status = readAsString(STATUS_COL, row);
    	String officeName = readAsString(OFFICE_NAME_COL, row);
        String officeId = getIdByName(workbook.getSheet("Offices"), officeName).toString();
        String staffName = readAsString(STAFF_NAME_COL, row);
        String staffId = getIdByName(workbook.getSheet("Staff"), staffName).toString();
        String externalId = readAsString(EXTERNAL_ID_COL, row);
        String activationDate = readAsDate(ACTIVATION_DATE_COL, row);
        String active = readAsBoolean(ACTIVE_COL, row).toString();
        String groupName = readAsString(NAME_COL, row);
        if(StringUtils.isBlank(groupName)) {
           	throw new IllegalArgumentException("Name is blank");
        }
        ArrayList<String> clientMemberIds = new ArrayList<String>();
        for(int cellNo = CLIENT_NAMES_STARTING_COL; cellNo < CLIENT_NAMES_ENDING_COL; cellNo++) {
        	String clientName = readAsString(cellNo, row);
        	if(clientName.equals(""))
        		break;
        	String clientId = getIdByName(workbook.getSheet("Clients"), clientName).toString();
        	if(!clientMemberIds.contains(clientId))
        		clientMemberIds.add(clientId);
        }
        return new Group(groupName, clientMemberIds, activationDate, active, externalId, officeId, staffId, row.getRowNum(), status);
	}
    
    private Meeting parseAsMeeting(Row row) {
    	String meetingStartDate = readAsDate(MEETING_START_DATE_COL, row);
    	String isRepeating = readAsBoolean(IS_REPEATING_COL, row).toString();
    	String repeats = readAsString(REPEATS_COL, row);
    	String repeatsEvery = readAsString(REPEATS_EVERY_COL, row);
    	String repeatsOnDay = readAsString(REPEATS_ON_DAY_COL, row);
    	if(!repeatsOnDay.equals(""))
    		repeatsOnDay = repeatsOnDay.substring(0, 2).toUpperCase(Locale.ENGLISH);
    	if(meetingStartDate.equals(""))
    		return null;
    	else {
    		if(repeatsOnDay.equals(""))
    			return new Meeting(meetingStartDate, isRepeating, repeats, repeatsEvery, row.getRowNum());
    		else
    			return new WeeklyMeeting(meetingStartDate, isRepeating, repeats, repeatsEvery, repeatsOnDay, row.getRowNum());
    	}
    }
    
    @Override
    public Result upload() {
        Result result = new Result();
        Sheet groupSheet = workbook.getSheet("Groups");
        int progressLevel = 0;
        String groupId = "";
        restClient.createAuthToken();
        for (int i = 0; i < groups.size(); i++) {
        	Row row = groupSheet.getRow(groups.get(i).getRowIndex());
        	Cell errorReportCell = row.createCell(FAILURE_COL);
        	Cell statusCell = row.createCell(STATUS_COL);
            try {
                String response = "";
                String status = groups.get(i).getStatus();
                progressLevel = getProgressLevel(status);
                
                if(progressLevel == 0)
                {
                   response = uploadGroup(i);
                   groupId = getGroupId(response);
                   progressLevel = 1;
                } else 
                	  groupId = readAsInt(GROUP_ID_COL, groupSheet.getRow(groups.get(i).getRowIndex()));
                
                if(meetings.get(i) != null)
                	progressLevel = uploadGroupMeeting(groupId, i);
                
                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (RuntimeException e) {
            	String message = parseStatus(e.getMessage());
            	String status = "";
            	
            	if(progressLevel == 0)
            		status = "Creation";
            	else if(progressLevel == 1)
            		status = "Meeting";
            	statusCell.setCellValue(status + " failed.");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));
                
                if(progressLevel>0)
                	row.createCell(GROUP_ID_COL).setCellValue(Integer.parseInt(groupId));
               
            	errorReportCell.setCellValue(message);
                result.addError("Row = " + groups.get(i).getRowIndex() + " ," + message);
            }
        }
        setReportHeaders(groupSheet);
        return result;
    }
    
    private int getProgressLevel(String status) {
        if(status.equals("") || status.equals("Creation failed."))
        	return 0;
        else if(status.equals("Meeting failed."))
        	return 1;
        return 0;
    }
    
    private String uploadGroup(int rowIndex) {
        String payload = new Gson().toJson(groups.get(rowIndex));
        logger.info(payload);
        String response = restClient.post("groups", payload);
    	return response;
    }
    
    private String getGroupId(String response) {
        JsonParser parser = new JsonParser();
        JsonObject obj = parser.parse(response).getAsJsonObject();
        return obj.get("groupId").getAsString();
    }
    
    private Integer uploadGroupMeeting(String groupId, int rowIndex) {
    	Meeting meeting = meetings.get(rowIndex);
    	meeting.setGroupId(groupId);
    	meeting.setTitle("groups_" + groupId + "_CollectionMeeting");
    	String payload = new Gson().toJson(meeting);
        logger.info(payload);
        restClient.post("groups/" + groupId + "/calendars", payload);
        return 2;
   }
    
    private void setReportHeaders(Sheet sheet) {
    	writeString(STATUS_COL, sheet.getRow(0), "Status");
    	writeString(GROUP_ID_COL, sheet.getRow(0), "Group Id");
    	writeString(FAILURE_COL, sheet.getRow(0), "Failure Report");
    }
    
    public List<Group> getGroups() {
        return groups;
    }
    
    public List<Meeting> getMeetings() {
        return meetings;
    }
}
