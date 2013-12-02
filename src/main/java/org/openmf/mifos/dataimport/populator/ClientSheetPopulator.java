package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.client.CompactClient;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class ClientSheetPopulator extends AbstractWorkbookPopulator {
	
	private static final Logger logger = LoggerFactory.getLogger(ClientSheetPopulator.class);
	
    private final RestClient restClient;

    private String content;
    
    private List<CompactClient> clients;
    private ArrayList<String> officeNames;
    
    private Map<String, ArrayList<String>> officeToClients;
    private Map<Integer, Integer[]> officeNameToBeginEndIndexesOfClients;
    private Map<String, Integer> clientNameToClientId;
    
    private static final int OFFICE_NAME_COL = 0;
    private static final int CLIENT_NAME_COL = 1;
    private static final int CLIENT_ID_COL = 2;
	
	public ClientSheetPopulator(RestClient restClient) {
    	this.restClient = restClient;
    }
	
	@Override
    public Result downloadAndParse() {
    	Result result = new Result();
    	try {
        	restClient.createAuthToken();
        	clients = new ArrayList<CompactClient>();
            content = restClient.get("clients?limit=-1");
            parseClients();
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
    	Sheet clientSheet = workbook.createSheet("Clients");
    	setLayout(clientSheet);
    	try{
    	     setOfficeToClientsMap();
    	     populateClientsByOfficeName(clientSheet);
    		 clientSheet.protectSheet("");
    	} catch (Exception e) {
    		result.addError(e.getMessage());
    		logger.error(e.getMessage());
    	}
        return result;
    }
    
    private void parseClients() {
    	Gson gson = new Gson();
        JsonParser parser = new JsonParser();
        JsonObject obj = parser.parse(content).getAsJsonObject();
        JsonArray array = obj.getAsJsonArray("pageItems");
        Iterator<JsonElement> iterator = array.iterator();
        clientNameToClientId = new HashMap<String, Integer>();
        while(iterator.hasNext()) {
        	JsonElement json = iterator.next();
        	CompactClient client = gson.fromJson(json, CompactClient.class);
        	if(client.isActive()) {
        	  clients.add(client);
        	}
        	clientNameToClientId.put(client.getDisplayName().trim(), client.getId());
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
    
    private void populateClientsByOfficeName(Sheet clientSheet) {
    	int rowIndex = 1, startIndex = 1, officeIndex = 0;
    	officeNameToBeginEndIndexesOfClients = new HashMap<Integer, Integer[]>();
	    Row row = clientSheet.createRow(rowIndex);
		for(String officeName : officeNames) {
			startIndex = rowIndex+1;
       	    writeString(OFFICE_NAME_COL, row, officeName);
       	    ArrayList<String> clientList = new ArrayList<String>();
       	    if(officeToClients.containsKey(officeName))
       	         clientList = officeToClients.get(officeName);
       	    
       	    if(!clientList.isEmpty()) {
       		   for(String clientName : clientList) {
       		       writeString(CLIENT_NAME_COL, row, clientName);
       		       writeInt(CLIENT_ID_COL, row, clientNameToClientId.get(clientName));
       		       row = clientSheet.createRow(++rowIndex);
       		   }
       		officeNameToBeginEndIndexesOfClients.put(officeIndex++, new Integer[]{startIndex, rowIndex});
       	    }
       	    else 
       	    	officeIndex++;
		}
    }
    
    private void setOfficeToClientsMap() {
    	officeToClients = new HashMap<String, ArrayList<String>>();
    	for(CompactClient person : clients) 
    		add(person.getOfficeName().trim().replaceAll("[ )(]", "_"), person.getDisplayName().trim());
    }
    
    //Guava Multi-map can reduce this.
    private void add(String key, String value) {
        ArrayList<String> values = officeToClients.get(key);
        if (values == null) {
            values = new ArrayList<String>();
        }
        values.add(value);
        officeToClients.put(key, values);
    }
    
    private void setLayout(Sheet worksheet) {
    	Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        worksheet.setColumnWidth(OFFICE_NAME_COL, 6000);
        for(int colIndex = 1; colIndex<=10; colIndex++)
           worksheet.setColumnWidth(colIndex, 6000);
        writeString(OFFICE_NAME_COL, rowHeader, "Office Names");
        writeString(CLIENT_NAME_COL, rowHeader, "Client Names");
        writeString(CLIENT_ID_COL, rowHeader, "Client ID");
    }
    
    public List<CompactClient> getClients() {
        return clients;
    }
    
    public String[] getOfficeNames() {
        return officeNames.toArray(new String[officeNames.size()]);
    }

    public Integer getClientsSize() {
    	return clients.size();
    }
    
    public Map<Integer, Integer[]> getOfficeNameToBeginEndIndexesOfClients() {
    	return officeNameToBeginEndIndexesOfClients;
    }
    
    public Map<String, ArrayList<String>> getOfficeToClients() {
    	return officeToClients;
    }
    
    public Map<String, Integer> getClientNameToClientId() {
    	return clientNameToClientId;
    }
}
