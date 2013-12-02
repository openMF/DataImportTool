package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.client.CompactGroup;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class GroupSheetPopulator extends AbstractWorkbookPopulator {
	
private static final Logger logger = LoggerFactory.getLogger(GroupSheetPopulator.class);
	
    private final RestClient restClient;

    private String content;
    
    private List<CompactGroup> groups;
    private ArrayList<String> officeNames;
    
    private Map<String, ArrayList<String>> officeToGroups;
    private Map<String, Integer> groupNameToGroupId;
    private Map<Integer, Integer[]> officeNameToBeginEndIndexesOfGroups;
    
    private static final int OFFICE_NAME_COL = 0;
    private static final int GROUP_NAME_COL = 1;
    private static final int GROUP_ID_COL = 2;
    
    public GroupSheetPopulator(RestClient restClient) {
    	this.restClient = restClient;
    }
    
    @Override
    public Result downloadAndParse() {
    	Result result = new Result();
    	try {
        	restClient.createAuthToken();
        	groups = new ArrayList<CompactGroup>();
            content = restClient.get("groups?limit=-1");
            parseGroups();
            content = restClient.get("offices?limit=-1");
            parseOfficeNames();
        } catch (Exception e) {
            result.addError(e.getMessage());
            logger.error(e.getMessage());
        }
    	return result;
    }

    @Override
    public Result populate(Workbook workbook) {
    	Result result = new Result();
    	Sheet groupSheet = workbook.createSheet("Groups");
    	setLayout(groupSheet);
    	try{
    	     setOfficeToGroupsMap();
    	     populateGroupsByOfficeName(groupSheet);
    	     groupSheet.protectSheet("");
    	} catch (Exception e) {
    		result.addError(e.getMessage());
    		logger.error(e.getMessage());
    	}
        return result;
    }
    
    private void parseGroups() {
    	Gson gson = new Gson();
        JsonParser parser = new JsonParser();
        JsonObject obj = parser.parse(content).getAsJsonObject();
        JsonArray array = obj.getAsJsonArray("pageItems");
        Iterator<JsonElement> iterator = array.iterator();
        groupNameToGroupId = new HashMap<String, Integer>();
        while(iterator.hasNext()) {
        	JsonElement json = iterator.next();
        	CompactGroup group = gson.fromJson(json, CompactGroup.class);
        	if(group.isActive())
        	  groups.add(group);
        	groupNameToGroupId.put(group.getName().trim(), group.getId());
        }
    }
    
    private void parseOfficeNames() {
    	JsonElement json = new JsonParser().parse(content);
    	JsonArray array = json.getAsJsonArray();
        Iterator<JsonElement> iterator = array.iterator();
        officeNames = new ArrayList<String>();
        while(iterator.hasNext()) {
        	String officeName = iterator.next().getAsJsonObject().get("name").toString();
        	officeName = officeName.substring(1, officeName.length()-1).trim().replaceAll("[ )(]", "_");
         officeNames.add(officeName);
        }
    }
    
    private void setOfficeToGroupsMap() {
    	officeToGroups = new HashMap<String, ArrayList<String>>();
    	for(CompactGroup group : groups) {
    		add(group.getOfficeName().trim().replaceAll("[ )(]", "_"), group.getName().trim());
    	}
    }
    
    //Guava Multi-map can reduce this.
    private void add(String key, String value) {
        ArrayList<String> values = officeToGroups.get(key);
        if (values == null) {
            values = new ArrayList<String>();
        }
        values.add(value);
        officeToGroups.put(key, values);
    }
    
    private void populateGroupsByOfficeName(Sheet groupSheet) {
    	int rowIndex = 1, officeIndex = 0, startIndex = 1;
    	officeNameToBeginEndIndexesOfGroups = new HashMap<Integer, Integer[]>();
    	Row row = groupSheet.createRow(rowIndex);
		for(String officeName : officeNames) {
			startIndex = rowIndex+1;
       	    writeString(OFFICE_NAME_COL, row, officeName);
       	    ArrayList<String> groupsList = new ArrayList<String>();
       	    
       	    if(officeToGroups.containsKey(officeName))
       	    	groupsList = officeToGroups.get(officeName);
       	    
       	 if(!groupsList.isEmpty()) {
     		   for(String groupName : groupsList) {
     		       writeString(GROUP_NAME_COL, row, groupName);
     		       writeInt(GROUP_ID_COL, row, groupNameToGroupId.get(groupName));
     		       row = groupSheet.createRow(++rowIndex);
     		   }
     		  officeNameToBeginEndIndexesOfGroups.put(officeIndex++, new Integer[]{startIndex, rowIndex});
     	    }
     	    else {
     	    	officeNameToBeginEndIndexesOfGroups.put(officeIndex++, new Integer[]{startIndex, rowIndex+1});
     	    }
		}
    }
    
    private void setLayout(Sheet worksheet) {
    	Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        for(int colIndex = 0; colIndex<=10; colIndex++)
           worksheet.setColumnWidth(colIndex, 6000);
        writeString(OFFICE_NAME_COL, rowHeader, "Office Names");
        writeString(GROUP_NAME_COL, rowHeader, "Group Names");
        writeString(GROUP_ID_COL, rowHeader, "Group ID");
    }
    
    public Integer getGroupsSize() {
    	return groups.size();
    }
    
    public List<CompactGroup> getGroups() {
    	return groups;
    }
    
    public Map<Integer, Integer[]> getOfficeNameToBeginEndIndexesOfGroups() {
    	return officeNameToBeginEndIndexesOfGroups;
    }
    
    public Map<String, Integer> getGroupNameToGroupId() {
    	return groupNameToGroupId;
    }
    
    public Map<String, ArrayList<String>> getOfficeToGroups() {
    	return officeToGroups;
    }

}
