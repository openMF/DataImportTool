package org.openmf.mifos.dataimport.populator.client;

import java.util.ArrayList;
import java.util.Arrays;

import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.populator.AbstractWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.OfficeSheetPopulator;
import org.openmf.mifos.dataimport.populator.PersonnelSheetPopulator;

public class CenterWorkbookPopulator extends AbstractWorkbookPopulator {

	@SuppressWarnings("CPD-START")
    private static final int CENTER_NAME_COL = 0;
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
    private static final int LOOKUP_OFFICE_NAME_COL = 251;
    private static final int LOOKUP_OFFICE_OPENING_DATE_COL = 252;
    private static final int LOOKUP_REPEAT_NORMAL_COL = 253;
    private static final int LOOKUP_REPEAT_MONTHLY_COL = 254;
    private static final int LOOKUP_IF_REPEAT_WEEKLY_COL = 255;
    @SuppressWarnings("CPD-END")
    
    private OfficeSheetPopulator officeSheetPopulator;
	private PersonnelSheetPopulator personnelSheetPopulator;
	
	
	public CenterWorkbookPopulator(OfficeSheetPopulator officeSheetPopulator, PersonnelSheetPopulator personnelSheetPopulator) {
    	this.officeSheetPopulator = officeSheetPopulator;
    	this.personnelSheetPopulator = personnelSheetPopulator;
    	
    }
	
	@Override
	public Result downloadAndParse() {
	    	Result result = officeSheetPopulator.downloadAndParse();
	    	if(result.isSuccess()) {
	    	   result = personnelSheetPopulator.downloadAndParse();
	    	}
	    	    	return result;
	}

	

	private void setLookupTable(Sheet centerSheet) {
		setOfficeDateLookupTable(centerSheet, officeSheetPopulator.getOffices(), LOOKUP_OFFICE_NAME_COL, LOOKUP_OFFICE_OPENING_DATE_COL);
    	int rowIndex;
    	for(rowIndex = 1; rowIndex <= 11; rowIndex++) {
    		Row row = centerSheet.getRow(rowIndex);
    		if(row == null)
    			row = centerSheet.createRow(rowIndex);
    		writeInt(LOOKUP_REPEAT_MONTHLY_COL, row, rowIndex);
    	}
    	for(rowIndex = 1; rowIndex <= 3; rowIndex++) 
    		writeInt(LOOKUP_REPEAT_NORMAL_COL, centerSheet.getRow(rowIndex), rowIndex);
    	String[] days = new String[]{"Mon","Tue","Wed","Thu","Fri","Sat","Sun"};
    	for(rowIndex = 1; rowIndex <= 7; rowIndex++) 
    		writeString(LOOKUP_IF_REPEAT_WEEKLY_COL, centerSheet.getRow(rowIndex), days[rowIndex-1]);
		
	}

	@Override
	public Result populate(Workbook workbook) {
		Sheet centerSheet = workbook.createSheet("Centers");
    	Result result = personnelSheetPopulator.populate(workbook);
    	if(result.isSuccess())
    	   result = officeSheetPopulator.populate(workbook);
    	
        setLayout(centerSheet);
        setLookupTable(centerSheet);
        if(result.isSuccess())
           result = setRules(centerSheet);
        return result;
	}
	
	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        worksheet.setColumnWidth(CENTER_NAME_COL, 4000);
        worksheet.setColumnWidth(OFFICE_NAME_COL, 5000);
        worksheet.setColumnWidth(STAFF_NAME_COL, 5000);
        worksheet.setColumnWidth(EXTERNAL_ID_COL, 2500);
        worksheet.setColumnWidth(ACTIVE_COL, 2000);
        worksheet.setColumnWidth(ACTIVATION_DATE_COL, 3500);
        worksheet.setColumnWidth(MEETING_START_DATE_COL, 3500);
        worksheet.setColumnWidth(IS_REPEATING_COL, 2000);
        worksheet.setColumnWidth(FREQUENCY_COL, 3000);
        worksheet.setColumnWidth(INTERVAL_COL, 2000);
        worksheet.setColumnWidth(REPEATS_ON_DAY_COL, 2500);
        worksheet.setColumnWidth(STATUS_COL, 2000);
        worksheet.setColumnWidth(CENTER_ID_COL, 2000);
        worksheet.setColumnWidth(FAILURE_COL, 2000);
       
        worksheet.setColumnWidth(LOOKUP_OFFICE_NAME_COL, 6000);
        worksheet.setColumnWidth(LOOKUP_OFFICE_OPENING_DATE_COL, 4000);
        worksheet.setColumnWidth(LOOKUP_REPEAT_NORMAL_COL, 3000);
        worksheet.setColumnWidth(LOOKUP_REPEAT_MONTHLY_COL, 3000);
        worksheet.setColumnWidth(LOOKUP_IF_REPEAT_WEEKLY_COL, 3000);
        
        writeString(CENTER_NAME_COL, rowHeader, "Center Name*");
        writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
        writeString(STAFF_NAME_COL, rowHeader, "Staff Name*");
        writeString(EXTERNAL_ID_COL, rowHeader, "External ID");
        writeString(ACTIVE_COL, rowHeader, "Active*");
        writeString(ACTIVATION_DATE_COL, rowHeader, "Activation Date*");
        writeString(MEETING_START_DATE_COL, rowHeader, "Meeting Start Date* (On or After)");
        writeString(IS_REPEATING_COL, rowHeader, "Repeat*");
        writeString(FREQUENCY_COL, rowHeader, "Frequency*");
        writeString(INTERVAL_COL, rowHeader, "Interval*");
        writeString(REPEATS_ON_DAY_COL, rowHeader, "Repeats On*");
     
        writeString(LOOKUP_OFFICE_NAME_COL, rowHeader, "Office Name");
        writeString(LOOKUP_OFFICE_OPENING_DATE_COL, rowHeader, "Opening Date");
        writeString(LOOKUP_REPEAT_NORMAL_COL, rowHeader, "Repeat Normal Range");
        writeString(LOOKUP_REPEAT_MONTHLY_COL, rowHeader, "Repeat Monthly Range");
        writeString(LOOKUP_IF_REPEAT_WEEKLY_COL, rowHeader, "If Repeat Weekly Range");

	}
	
	private Result setRules(Sheet worksheet) {
		Result result = new Result();
    	try {
    	CellRangeAddressList officeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), OFFICE_NAME_COL, OFFICE_NAME_COL);
    	CellRangeAddressList staffNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), STAFF_NAME_COL, STAFF_NAME_COL);
    	CellRangeAddressList dateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), ACTIVATION_DATE_COL, ACTIVATION_DATE_COL);
    	CellRangeAddressList activeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), ACTIVE_COL, ACTIVE_COL);
    	CellRangeAddressList meetingStartDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), MEETING_START_DATE_COL, MEETING_START_DATE_COL);
    	CellRangeAddressList isRepeatRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), IS_REPEATING_COL, IS_REPEATING_COL);
    	CellRangeAddressList repeatsRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), FREQUENCY_COL, FREQUENCY_COL);
    	CellRangeAddressList repeatsEveryRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), INTERVAL_COL, INTERVAL_COL);
    	CellRangeAddressList repeatsOnRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPEATS_ON_DAY_COL, REPEATS_ON_DAY_COL);
    	
    	
    	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
    	ArrayList<String> officeNames = new ArrayList<String>(Arrays.asList(officeSheetPopulator.getOfficeNames()));
    	setNames(worksheet, officeNames);
    	

    	DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
    	DataValidationConstraint staffNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Staff_\",$B1))");
    	DataValidationConstraint activationDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=VLOOKUP($B1,$IR$2:$IS" + (officeNames.size() + 1)+",2,FALSE)", "=TODAY()", "dd/mm/yy");
    	DataValidationConstraint booleanConstraint = validationHelper.createExplicitListConstraint(new String[]{"True", "False"});
    	DataValidationConstraint meetingStartDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=$F1", "=TODAY()", "dd/mm/yy");
    	DataValidationConstraint repeatsConstraint = validationHelper.createExplicitListConstraint(new String[]{"Daily", "Weekly", "Monthly", "Yearly"});
    	DataValidationConstraint repeatsEveryConstraint = validationHelper.createFormulaListConstraint("INDIRECT($I1)");
    	DataValidationConstraint repeatsOnConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE($I1,\"_DAYS\"))");


    	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
    	DataValidation staffValidation = validationHelper.createValidation(staffNameConstraint, staffNameRange);
    	DataValidation activationDateValidation = validationHelper.createValidation(activationDateConstraint, dateRange);
    	DataValidation activeValidation = validationHelper.createValidation(booleanConstraint, activeRange);
    	DataValidation meetingStartDateValidation = validationHelper.createValidation(meetingStartDateConstraint, meetingStartDateRange);
    	DataValidation isRepeatValidation = validationHelper.createValidation(booleanConstraint, isRepeatRange);
    	DataValidation repeatsValidation = validationHelper.createValidation(repeatsConstraint, repeatsRange);
    	DataValidation repeatsEveryValidation = validationHelper.createValidation(repeatsEveryConstraint, repeatsEveryRange);
    	DataValidation repeatsOnValidation = validationHelper.createValidation(repeatsOnConstraint, repeatsOnRange);
    	

    	worksheet.addValidationData(activeValidation);
        worksheet.addValidationData(officeValidation);
        worksheet.addValidationData(staffValidation);
        worksheet.addValidationData(activationDateValidation);
        worksheet.addValidationData(meetingStartDateValidation);
        worksheet.addValidationData(isRepeatValidation);
        worksheet.addValidationData(repeatsValidation);
        worksheet.addValidationData(repeatsEveryValidation);
        worksheet.addValidationData(repeatsOnValidation);
        
    	} catch (RuntimeException re) {
    		result.addError(re.getMessage());
    	}
       return result;
	}
	
	private void setNames(Sheet worksheet, ArrayList<String> officeNames) {
    	Workbook centerWorkbook = worksheet.getWorkbook();
    	Name officeCenter = centerWorkbook.createName();
    	officeCenter.setNameName("Office");
    	officeCenter.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));
    	
    	//Repeat constraint names
    	Name repeatsDaily = centerWorkbook.createName();
    	repeatsDaily.setNameName("Daily");
    	repeatsDaily.setRefersToFormula("Centers!$IT$2:$IT$4");
    	Name repeatsWeekly = centerWorkbook.createName();
    	repeatsWeekly.setNameName("Weekly");
    	repeatsWeekly.setRefersToFormula("Centers!$IT$2:$IT$4");
    	Name repeatYearly = centerWorkbook.createName();
    	repeatYearly.setNameName("Yearly");
    	repeatYearly.setRefersToFormula("Centers!$IT$2:$IT$4");
    	Name repeatsMonthly = centerWorkbook.createName();
    	repeatsMonthly.setNameName("Monthly");
    	repeatsMonthly.setRefersToFormula("Centers!$IU$2:$IU$12");
    	Name repeatsOnWeekly = centerWorkbook.createName();
    	repeatsOnWeekly.setNameName("Weekly_Days");
    	repeatsOnWeekly.setRefersToFormula("Centers!$IV$2:$IV$8");
    	
    	
    	//Staff Names for each office
    	for(Integer i = 0; i < officeNames.size(); i++) {
    		//Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator.getOfficeNameToBeginEndIndexesOfClients().get(i);
    		Integer[] officeNameToBeginEndIndexesOfStaff = personnelSheetPopulator.getOfficeNameToBeginEndIndexesOfStaff().get(i);
    		//Name clientName = groupWorkbook.createName();
    		Name loanOfficerName = centerWorkbook.createName();
    		 /*if(officeNameToBeginEndIndexesOfClients != null) {
    	          clientName.setNameName("Client_" + officeNames.get(i));
    	          clientName.setRefersToFormula("Clients!$B$" + officeNameToBeginEndIndexesOfClients[0] + ":$B$" + officeNameToBeginEndIndexesOfClients[1]);
    		 }*/
    		 if(officeNameToBeginEndIndexesOfStaff != null) {
    	        loanOfficerName.setNameName("Staff_" + officeNames.get(i));
    	        loanOfficerName.setRefersToFormula("Staff!$B$" + officeNameToBeginEndIndexesOfStaff[0] + ":$B$" + officeNameToBeginEndIndexesOfStaff[1]);
    		 }
    	}
		
	}
	
	
}
