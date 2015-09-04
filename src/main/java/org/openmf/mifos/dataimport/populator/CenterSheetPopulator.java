package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.client.CompactCenter;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class CenterSheetPopulator extends AbstractWorkbookPopulator {

	 

	private final RestClient restClient;
	private String content;

	private List<CompactCenter> centers;
	private ArrayList<String> officeNames;

	private Map<String, ArrayList<String>> officeToCenters;
	private Map<String, Integer> centerNameToCenterId;
	private Map<Integer, Integer[]> officeNameToBeginEndIndexesOfCenters;

	private static final int OFFICE_NAME_COL = 0;
	private static final int CENTER_NAME_COL = 1;
	private static final int CENTER_ID_COL = 2;

	public CenterSheetPopulator(RestClient restClient) {
		this.restClient = restClient;
	}

	@Override
	public Result downloadAndParse() {
		Result result = new Result();
		try {
			restClient.createAuthToken();
			centers = new ArrayList<CompactCenter>();
			content = restClient.get("centers?paged=true&limit=-1");
			parseCenters();
			content = restClient.get("offices?limit=-1");
			parseOfficeNames();
		} catch (Exception e) {
			result.addError(e.getMessage());
		}
		return result;
	}

	@Override
	public Result populate(Workbook workbook) {
		Result result = new Result();
		Sheet centerSheet = workbook.createSheet("Center");
		setLayout(centerSheet);
		try {
			setOfficeToCentersMap();
			populateCentersByOfficeName(centerSheet);
			centerSheet.protectSheet("");
		} catch (Exception e) {
			result.addError(e.getMessage());
		}
		return result;
	}
	
	private void parseCenters() {
		Gson gson = new Gson();
		JsonParser parser = new JsonParser();
		JsonObject obj = parser.parse(content).getAsJsonObject();
		JsonArray array = obj.getAsJsonArray("pageItems");
		Iterator<JsonElement> iterator = array.iterator();
		centerNameToCenterId = new HashMap<String, Integer>();
		while (iterator.hasNext()) {
			JsonElement json = iterator.next();
			CompactCenter center = gson.fromJson(json, CompactCenter.class);
			if (center.isActive())
				centers.add(center);
			centerNameToCenterId.put(center.getName().trim(), center.getId());
		}
	}
	
	private void parseOfficeNames() {
		JsonElement json = new JsonParser().parse(content);
		JsonArray array = json.getAsJsonArray();
		Iterator<JsonElement> iterator = array.iterator();
		officeNames = new ArrayList<String>();
		while (iterator.hasNext()) {
			String officeName = iterator.next().getAsJsonObject().get("name")
					.toString();
			officeName = officeName.substring(1, officeName.length() - 1)
					.trim().replaceAll("[ )(]", "_");
			officeNames.add(officeName);
		}

	}

	private void setOfficeToCentersMap() {
		officeToCenters = new HashMap<String, ArrayList<String>>();
		for (CompactCenter center : centers) {
			add(center.getOfficeName().trim().replaceAll("[ )(]", "_"), center
					.getName().trim());
		}
	}	
	
	// Guava Multi-map can reduce this.
	private void add(String key, String value) {
			ArrayList<String> values = officeToCenters.get(key);
			if (values == null) {
				values = new ArrayList<String>();
			}
			values.add(value);
			officeToCenters.put(key, values);
	}	
	
	private void populateCentersByOfficeName(Sheet centerSheet) {
		int rowIndex = 1, officeIndex = 0, startIndex = 1;
		officeNameToBeginEndIndexesOfCenters = new HashMap<Integer, Integer[]>();
		Row row = centerSheet.createRow(rowIndex);
		for (String officeName : officeNames) {
			startIndex = rowIndex + 1;
			writeString(OFFICE_NAME_COL, row, officeName);
			ArrayList<String> centersList = new ArrayList<String>();

			if (officeToCenters.containsKey(officeName))
				centersList = officeToCenters.get(officeName);

			if (!centersList.isEmpty()) {
				for (String centerName : centersList) {
					writeString(CENTER_NAME_COL, row, centerName);
					writeInt(CENTER_ID_COL, row,
							centerNameToCenterId.get(centerName));
					row = centerSheet.createRow(++rowIndex);
				}
				officeNameToBeginEndIndexesOfCenters.put(officeIndex++,
						new Integer[] { startIndex, rowIndex });
			} else {
				officeNameToBeginEndIndexesOfCenters.put(officeIndex++,
						new Integer[] { startIndex, rowIndex + 1 });
			}
		}
	}
	
	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
		rowHeader.setHeight((short) 500);
		for (int colIndex = 0; colIndex <= 10; colIndex++)
			worksheet.setColumnWidth(colIndex, 6000);
		writeString(OFFICE_NAME_COL, rowHeader, "Office Names");
		writeString(CENTER_NAME_COL, rowHeader, "Center Names");
		writeString(CENTER_ID_COL, rowHeader, "Center ID");
	}

	public Integer getCentersSize() {
		return centers.size();
	}

	public List<CompactCenter> getCenters() {
		return centers;
	}

	public Map<Integer, Integer[]> getOfficeNameToBeginEndIndexesOfCenters() {
		return officeNameToBeginEndIndexesOfCenters;
	}

	public Map<String, Integer> getCenterNameToCenterId() {
		return centerNameToCenterId;
	}

	public Map<String, ArrayList<String>> getOfficeToCenters() {
		return officeToCenters;
	}

}
