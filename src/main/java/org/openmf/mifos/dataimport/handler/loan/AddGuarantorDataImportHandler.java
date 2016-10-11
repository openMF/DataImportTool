package org.openmf.mifos.dataimport.handler.loan;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.AddGuarantor;
import org.openmf.mifos.dataimport.handler.AbstractDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;

import com.google.gson.Gson;

public class AddGuarantorDataImportHandler extends AbstractDataImportHandler  {
	 
	
	 private final RestClient restClient;
	    private final Workbook workbook;
	    private List<AddGuarantor> guarantors;
	    
	    private String loanAccountId = "";
	    
	    private static final int LOAN_ACCOUNT_NO_COL = 2;
	    private static final int GUARANTO_TYPE_COL =3;
	    private static final int CLIENT_RELATIONSHIP_TYPE_COL =4;
	    
	    private static final int ENTITY_ID_COL = 6;
	    
	    //private static final int ENTITY_ID_COL = 5;
	    private static final int FIRST_NAME_COL = 7;
	    private static final int LAST_NAME_COL = 8;
	    private static final int ADDRESS_LINE_1_COL= 9;
	    private static final int ADDRESS_LINE_2_COL = 10;
	    private static final int CITY_COL = 11;
	    private static final int DOB_COL = 12;
	    private static final int ZIP_COL = 13;
	    private static final int SAVINGS_ID=14;
	    private static final int AMOUNT=15;
	    private static final int STATUS_COL = 18;
	    
	    public AddGuarantorDataImportHandler(Workbook workbook, RestClient client) {
	        this.workbook = workbook;
	        this.restClient = client;
	        guarantors = new ArrayList<AddGuarantor>();
	    }

	    @Override
	    public Result parse() {
	        Result result = new Result();
	        Sheet addGuarantorSheet = workbook.getSheet("guarantor");
	        Integer noOfEntries = getNumberOfRows(addGuarantorSheet, LOAN_ACCOUNT_NO_COL);
	        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
	            Row row;
	            try {
	                row = addGuarantorSheet.getRow(rowIndex);
	                if (isNotImported(row, STATUS_COL))
	                	guarantors.add(parseAsGuarantor(row));
	                    
	            } catch (Exception e) {
	                result.addError("Row = " + rowIndex + " , " + e.getMessage());
	            }
	        }
	        return result;
	    }

	
	
	
	private AddGuarantor parseAsGuarantor(Row row) {
   	 String loanAccountIdCheck = readAsInt(LOAN_ACCOUNT_NO_COL, row);
        if(!loanAccountIdCheck.equals(""))
           loanAccountId = loanAccountIdCheck;
        String guarantorType = readAsInt(GUARANTO_TYPE_COL, row);
        
        String guarantorTypeId = "";
        if (guarantorType.equalsIgnoreCase("Internal"))
        	guarantorTypeId = "1";
        else if(guarantorType.equalsIgnoreCase("External"))
        	guarantorTypeId = "3";
        String clientName = readAsString(ENTITY_ID_COL, row);
        String entityId = getIdByName(workbook.getSheet("Clients"), clientName).toString();
        String clientRelationshipTypeId = readAsInt(CLIENT_RELATIONSHIP_TYPE_COL, row);
        //String entityId = readAsInt(ENTITY_ID_COL, row);
        String firstname = readAsString(FIRST_NAME_COL, row);
        String lastname = readAsString(LAST_NAME_COL, row);
        String addressLine1 = readAsString(ADDRESS_LINE_1_COL, row);
        String addressLine2 = readAsString(ADDRESS_LINE_2_COL, row);
        String city = readAsString(CITY_COL, row);
        String dob = readAsDate(DOB_COL, row);
        String zip = readAsLong(ZIP_COL, row);
        String savingsId = readAsInt(SAVINGS_ID, row);
        String amount = readAsDouble(AMOUNT, row).toString();
        
        
		return new AddGuarantor(guarantorTypeId,clientRelationshipTypeId,entityId,firstname,lastname,addressLine1,addressLine2,city,dob,zip,savingsId,amount, row.getRowNum(),
				
			Integer.parseInt(loanAccountId));
	}

	@Override
    public Result upload() {
        Result result = new Result();
        Sheet addGuarantorSheet = workbook.getSheet("guarantor");
        restClient.createAuthToken();
        for (AddGuarantor addGuarantor : guarantors) {
            try {
                Gson gson = new Gson();
                String payload = gson.toJson(addGuarantor);
                restClient.post("loans/" + addGuarantor.getAccountId() + "/guarantors", payload);
                Cell statusCell = addGuarantorSheet.getRow(addGuarantor.getRowIndex()).createCell(STATUS_COL);
                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (Exception e) {
            	Cell loanAccountIdCell = addGuarantorSheet.getRow(addGuarantor.getRowIndex()).createCell(LOAN_ACCOUNT_NO_COL);
                loanAccountIdCell.setCellValue(addGuarantor.getAccountId());
            	String message = parseStatus(e.getMessage());
            	Cell statusCell = addGuarantorSheet.getRow(addGuarantor.getRowIndex()).createCell(STATUS_COL);
            	statusCell.setCellValue(message);
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));
                result.addError("Row = " + addGuarantor.getRowIndex() + " ," + message);
            }
        }
        addGuarantorSheet.setColumnWidth(STATUS_COL, 15000);
    	writeString(STATUS_COL, addGuarantorSheet.getRow(0), "Status");
        return result;
    }
	public List<AddGuarantor> getGuatantor() {
    	return guarantors;
    }
		
}
