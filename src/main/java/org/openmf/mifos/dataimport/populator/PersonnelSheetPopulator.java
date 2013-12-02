package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Office;
import org.openmf.mifos.dataimport.dto.Personnel;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class PersonnelSheetPopulator extends AbstractWorkbookPopulator {

    private static final Logger logger = LoggerFactory.getLogger(PersonnelSheetPopulator.class);
	
	private final RestClient client;
	private final Boolean onlyLoanOfficers;
	
	private String content;
	
	private List<Personnel> personnel;
	private List<Office> offices;
	
	//Maintaining the one to many relationship
	private Map<String, ArrayList<String>> officeToPersonnel;
	private Map<String, Integer> staffNameToStaffId;
	
	private Map<Integer, Integer[]> officeNameToBeginEndIndexesOfStaff;
	private Map<Integer,String> officeIdToOfficeName;
	
	private static final int OFFICE_NAME_COL = 0;
	private static final int STAFF_NAME_COL = 1;
	private static final int STAFF_ID_COL = 2;
	
	public PersonnelSheetPopulator(Boolean onlyLoanOfficers, RestClient client) {
		this.onlyLoanOfficers = onlyLoanOfficers;
        this.client = client;
    }
	
	 @Override
	    public Result downloadAndParse() {
	        Result result = new Result();
	        try {
	        	client.createAuthToken();
	        	personnel = new ArrayList<Personnel>();
	            content = client.get("staff?limit=-1");
	            parseStaff();
	            content = client.get("offices?limit=-1");
	            parseOffices();
	        } catch (RuntimeException re) {
	            result.addError(re.getMessage());
	            logger.error(re.getMessage());
	        }
	        return result;
	    }

	    @Override
	    public Result populate(Workbook workbook) {
	    	Result result = new Result();
	    	Sheet staffSheet = workbook.createSheet("Staff");
	        setLayout(staffSheet);
	    	try{
	        setOfficeToPersonnelMap();
	        populateStaffByOfficeName(staffSheet);
	        staffSheet.protectSheet("");
	    	} catch (RuntimeException re) {
	    		result.addError(re.getMessage());
	    		logger.error(re.getMessage());
	    	}
	        return result;
	    }
	    
	    private void parseStaff() {
	    	Gson gson = new Gson();
            JsonElement json = new JsonParser().parse(content);
            JsonArray array = json.getAsJsonArray();
            Iterator<JsonElement> iterator = array.iterator();
            staffNameToStaffId = new HashMap<String, Integer>();
            while(iterator.hasNext()) {
            	json = iterator.next();
            	Personnel person = gson.fromJson(json, Personnel.class);
            	if(!onlyLoanOfficers)
            	    personnel.add(person);
            	else{
            	   if(person.isLoanOfficer())
            		   personnel.add(person);
            	}
            	staffNameToStaffId.put(person.getName(), person.getId());
            }
	    }
	    
	    private void parseOffices() {
	    	offices = new ArrayList<Office>();
            JsonElement json = new JsonParser().parse(content);
            JsonArray array = json.getAsJsonArray();
            Iterator<JsonElement> iterator = array.iterator();
            officeIdToOfficeName = new HashMap<Integer,String>();
            while(iterator.hasNext()) {
            	json = iterator.next();
            	Office office = new Gson().fromJson(json, Office.class);
            	officeIdToOfficeName.put(office.getId(), office.getName().trim().replaceAll("[ )(]", "_"));
            	offices.add(office);
            }
	    }
	    
	    private void populateStaffByOfficeName(Sheet staffSheet) {
	    	int rowIndex = 1, startIndex = 1, officeIndex = 0;
	    	officeNameToBeginEndIndexesOfStaff = new HashMap<Integer, Integer[]>();
	    	Row row = staffSheet.createRow(rowIndex);
	        for(Office office : offices) {
	        	startIndex = rowIndex+1;
	        	writeString(OFFICE_NAME_COL, row, office.getName().trim().replaceAll("[ )(]", "_"));
	        	
	        	ArrayList<String> fullStaffList = getStaffList(office.getHierarchy());
	        	
	        	if(!fullStaffList.isEmpty()) {
	        		for(String staffName : fullStaffList) {
	        			int staffId = staffNameToStaffId.get(staffName);
	        		    writeString(STAFF_NAME_COL, row, staffName);
	        		    writeInt(STAFF_ID_COL, row, staffId);
	        		    row = staffSheet.createRow(++rowIndex);	
	        		}
	        		officeNameToBeginEndIndexesOfStaff.put(officeIndex++, new Integer[]{startIndex, rowIndex});
	        	} else 
	        		  officeIndex++;
	        }
	    }
	    
	    private void setOfficeToPersonnelMap() {
	    	officeToPersonnel = new HashMap<String, ArrayList<String>>();
	    	for(Personnel person : personnel) {
	    		add(person.getOfficeName().trim().replaceAll("[ )(]", "_"), person.getName().trim());
	    	}
	    }
	    
	    //Guava Multi-map can reduce this.
	    private void add(String key, String value) {
	        ArrayList<String> values = officeToPersonnel.get(key);
	        if (values == null) {
	            values = new ArrayList<String>();
	        }
	        values.add(value);
	        officeToPersonnel.put(key, values);
	    }
	    
	    private ArrayList<String> getStaffList(String hierarchy) {
	    	ArrayList<String> fullStaffList = new ArrayList<String>();
	    	Integer hierarchyLength = hierarchy.length();
			String[] officeIds = hierarchy.substring(1, hierarchyLength).split("\\.");
			String headOffice = offices.get(0).getName().trim().replaceAll("[ )(]", "_");
			if(officeToPersonnel.containsKey(headOffice))
			    fullStaffList.addAll(officeToPersonnel.get(headOffice));
			if(officeIds[0].isEmpty())
				return fullStaffList;
			for(int i=0; i<officeIds.length; i++) {
				String officeName = getOfficeNameFromOfficeId(Integer.parseInt(officeIds[i]));
				if(officeToPersonnel.containsKey(officeName))
	    	        fullStaffList.addAll(officeToPersonnel.get(officeName));
			}
	    	return fullStaffList;
	    }
	    
	    private String getOfficeNameFromOfficeId(Integer officeId) {
	    	return officeIdToOfficeName.get(officeId);
	    }
	    
	    
	    private void setLayout(Sheet worksheet) {
	    	for(Integer i=0; i<3; i++)
	    		worksheet.setColumnWidth(i, 6000);
	        Row rowHeader = worksheet.createRow(0);
	        rowHeader.setHeight((short)500);
	        writeString(OFFICE_NAME_COL, rowHeader, "Office Name");
	        writeString(STAFF_NAME_COL, rowHeader, "Staff List");
	        writeString(STAFF_ID_COL, rowHeader, "Staff ID");
	    }
	    
	    public List<Personnel> getPersonnel() {
	        return personnel;
	    }
	    
	    public List<Office> getOffices() {
	    	return offices;
	    }
	    
	    public Map<String, ArrayList<String>> getOfficeToPersonnel() {
	    	return officeToPersonnel;
	    }
	    
	    public Map<Integer, Integer[]> getOfficeNameToBeginEndIndexesOfStaff() {
	    	return officeNameToBeginEndIndexesOfStaff;
	    }
	    
	    public Map<Integer,String> getOfficeIdToOfficeName() {
	    	return officeIdToOfficeName;
	    }
	    
	    public Map<String, Integer> getStaffNameToStaffId() {
	    	return staffNameToStaffId;
	    }
}
