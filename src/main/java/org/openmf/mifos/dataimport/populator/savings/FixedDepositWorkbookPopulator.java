package org.openmf.mifos.dataimport.populator.savings;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.openmf.mifos.dataimport.dto.savings.FixedDepositProduct;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.populator.AbstractWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.ClientSheetPopulator;
import org.openmf.mifos.dataimport.populator.OfficeSheetPopulator;
import org.openmf.mifos.dataimport.populator.PersonnelSheetPopulator;

public class FixedDepositWorkbookPopulator extends AbstractWorkbookPopulator {
	
    private OfficeSheetPopulator officeSheetPopulator;
    private ClientSheetPopulator clientSheetPopulator;
    private PersonnelSheetPopulator personnelSheetPopulator;
    private FixedDepositProductSheetPopulator productSheetPopulator;
    
    @SuppressWarnings("CPD-START")
    private static final int OFFICE_NAME_COL = 0;
    private static final int CLIENT_NAME_COL = 1;
    private static final int PRODUCT_COL = 2;
    private static final int FIELD_OFFICER_NAME_COL = 3;
    private static final int SUBMITTED_ON_DATE_COL = 4;
    private static final int APPROVED_DATE_COL = 5;
    private static final int ACTIVATION_DATE_COL = 6;
    private static final int INTEREST_COMPOUNDING_PERIOD_COL = 7;
    private static final int INTEREST_POSTING_PERIOD_COL = 8;
    private static final int INTEREST_CALCULATION_COL = 9;
    private static final int INTEREST_CALCULATION_DAYS_IN_YEAR_COL = 10;
    private static final int LOCKIN_PERIOD_COL = 11;
    private static final int LOCKIN_PERIOD_FREQUENCY_COL = 12;
    private static final int DEPOSIT_AMOUNT_COL = 13;
    private static final int DEPOSIT_PERIOD_COL = 14;
    private static final int DEPOSIT_PERIOD_FREQUENCY_COL = 15;
    private static final int EXTERNAL_ID_COL = 16;
    private static final int CHARGE_ID_1 = 18;
    private static final int CHARGE_AMOUNT_1 = 19;
    private static final int CHARGE_DUE_DATE_1 = 20;
    private static final int CHARGE_ID_2 = 21;
    private static final int CHARGE_AMOUNT_2 = 22;
    private static final int CHARGE_DUE_DATE_2 = 23;
    private static final int CLOSED_ON_DATE = 24;
    private static final int ON_ACCOUNT_CLOSURE_ID = 25;
    private static final int TO_SAVINGS_ACCOUNT_ID = 26;
    
    private static final int LOOKUP_CLIENT_NAME_COL = 31;
    private static final int LOOKUP_ACTIVATION_DATE_COL = 32;
    


    @SuppressWarnings("CPD-END")
    
    public FixedDepositWorkbookPopulator(OfficeSheetPopulator officeSheetPopulator, ClientSheetPopulator clientSheetPopulator,
            PersonnelSheetPopulator personnelSheetPopulator,
            FixedDepositProductSheetPopulator productSheetPopulator) {
        this.officeSheetPopulator = officeSheetPopulator;
        this.clientSheetPopulator = clientSheetPopulator;
        this.personnelSheetPopulator = personnelSheetPopulator;
        this.productSheetPopulator = productSheetPopulator;
    }

	@Override
	public Result downloadAndParse() {
		Result result = officeSheetPopulator.downloadAndParse();
        if (result.isSuccess()) result = clientSheetPopulator.downloadAndParse();
        if (result.isSuccess()) result = personnelSheetPopulator.downloadAndParse();
        if (result.isSuccess()) result = productSheetPopulator.downloadAndParse();
        return result;
	}

	@Override
	public Result populate(Workbook workbook) {
		Sheet fixedDepositSheet = workbook.createSheet("FixedDeposit");
        Result result = officeSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = clientSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = personnelSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = productSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = setRules(fixedDepositSheet);
        if (result.isSuccess()) result = setDefaults(fixedDepositSheet);
        if (result.isSuccess())
            setClientAndGroupDateLookupTable(fixedDepositSheet, clientSheetPopulator.getClients(), null,
                    LOOKUP_CLIENT_NAME_COL, LOOKUP_ACTIVATION_DATE_COL);
        setLayout(fixedDepositSheet);
        return result;
	}

	private Result setRules(Sheet worksheet) {
		Result result = new Result();
        try {
            CellRangeAddressList officeNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    OFFICE_NAME_COL, OFFICE_NAME_COL);
            CellRangeAddressList clientNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    CLIENT_NAME_COL, CLIENT_NAME_COL);
            CellRangeAddressList productNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), PRODUCT_COL,
                    PRODUCT_COL);
            CellRangeAddressList fieldOfficerRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    FIELD_OFFICER_NAME_COL, FIELD_OFFICER_NAME_COL);
            CellRangeAddressList submittedDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    SUBMITTED_ON_DATE_COL, SUBMITTED_ON_DATE_COL);
            CellRangeAddressList approvedDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    APPROVED_DATE_COL, APPROVED_DATE_COL);
            CellRangeAddressList activationDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    ACTIVATION_DATE_COL, ACTIVATION_DATE_COL);
            CellRangeAddressList interestCompudingPeriodRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    INTEREST_COMPOUNDING_PERIOD_COL, INTEREST_COMPOUNDING_PERIOD_COL);
            CellRangeAddressList interestPostingPeriodRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    INTEREST_POSTING_PERIOD_COL, INTEREST_POSTING_PERIOD_COL);
            CellRangeAddressList interestCalculationRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    INTEREST_CALCULATION_COL, INTEREST_CALCULATION_COL);
            CellRangeAddressList interestCalculationDaysInYearRange = new CellRangeAddressList(1,
                    SpreadsheetVersion.EXCEL97.getLastRowIndex(), INTEREST_CALCULATION_DAYS_IN_YEAR_COL,
                    INTEREST_CALCULATION_DAYS_IN_YEAR_COL);
            CellRangeAddressList lockinPeriodFrequencyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    LOCKIN_PERIOD_FREQUENCY_COL, LOCKIN_PERIOD_FREQUENCY_COL);
            CellRangeAddressList depositAmountRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		DEPOSIT_AMOUNT_COL, DEPOSIT_AMOUNT_COL);
            CellRangeAddressList depositPeriodTypeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		DEPOSIT_PERIOD_FREQUENCY_COL, DEPOSIT_PERIOD_FREQUENCY_COL);
            
            DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet) worksheet);

            setNames(worksheet);

            DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
            DataValidationConstraint clientNameConstraint = validationHelper
                    .createFormulaListConstraint("INDIRECT(CONCATENATE(\"Client_\",$A1))");
            DataValidationConstraint productNameConstraint = validationHelper.createFormulaListConstraint("Products");
            DataValidationConstraint fieldOfficerNameConstraint = validationHelper
                    .createFormulaListConstraint("INDIRECT(CONCATENATE(\"Staff_\",$A1))");
            DataValidationConstraint submittedDateConstraint = validationHelper.createDateConstraint(
                    DataValidationConstraint.OperatorType.BETWEEN, "=VLOOKUP($B1,$AF$2:$AG$"
                            + (clientSheetPopulator.getClientsSize() + 1) + ",2,FALSE)", "=TODAY()",
                    "dd/mm/yy");
            DataValidationConstraint approvalDateConstraint = validationHelper.createDateConstraint(
                    DataValidationConstraint.OperatorType.BETWEEN, "=$E1", "=TODAY()", "dd/mm/yy");
            DataValidationConstraint activationDateConstraint = validationHelper.createDateConstraint(
                    DataValidationConstraint.OperatorType.BETWEEN, "=$F1", "=TODAY()", "dd/mm/yy");
            DataValidationConstraint interestCompudingPeriodConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "Daily", "Monthly", "Quarterly", "Semi-Annual", "Annually" });
            DataValidationConstraint interestPostingPeriodConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "Monthly", "Quarterly", "BiAnnual", "Annually" });
            DataValidationConstraint interestCalculationConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "Daily Balance", "Average Daily Balance" });
            DataValidationConstraint interestCalculationDaysInYearConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "360 Days", "365 Days" });
            DataValidationConstraint frequency = validationHelper.createExplicitListConstraint(new String[] { "Days",
                    "Weeks", "Months", "Years" });
            DataValidationConstraint depositConstraint = validationHelper.createDecimalConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "=INDIRECT(CONCATENATE(\"Min_Deposit_\",$C1))", null);

            DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
            DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
            DataValidation productNameValidation = validationHelper.createValidation(productNameConstraint, productNameRange);
            DataValidation fieldOfficerValidation = validationHelper.createValidation(fieldOfficerNameConstraint, fieldOfficerRange);
            DataValidation interestCompudingPeriodValidation = validationHelper.createValidation(interestCompudingPeriodConstraint,
                    interestCompudingPeriodRange);
            DataValidation interestPostingPeriodValidation = validationHelper.createValidation(interestPostingPeriodConstraint,
                    interestPostingPeriodRange);
            DataValidation interestCalculationValidation = validationHelper.createValidation(interestCalculationConstraint,
                    interestCalculationRange);
            DataValidation interestCalculationDaysInYearValidation = validationHelper.createValidation(
                    interestCalculationDaysInYearConstraint, interestCalculationDaysInYearRange);
            DataValidation lockinPeriodFrequencyValidation = validationHelper.createValidation(frequency,
                    lockinPeriodFrequencyRange);
            DataValidation depositPeriodTypeValidation = validationHelper.createValidation(frequency,
            		depositPeriodTypeRange);
            DataValidation submittedDateValidation = validationHelper.createValidation(submittedDateConstraint, submittedDateRange);
            DataValidation approvalDateValidation = validationHelper.createValidation(approvalDateConstraint, approvedDateRange);
            DataValidation activationDateValidation = validationHelper.createValidation(activationDateConstraint, activationDateRange);
            DataValidation  depositAmountValidation = validationHelper.createValidation(depositConstraint, depositAmountRange);
            
            
            worksheet.addValidationData(officeValidation);
            worksheet.addValidationData(clientValidation);
            worksheet.addValidationData(productNameValidation);
            worksheet.addValidationData(fieldOfficerValidation);
            worksheet.addValidationData(submittedDateValidation);
            worksheet.addValidationData(approvalDateValidation);
            worksheet.addValidationData(activationDateValidation);
            worksheet.addValidationData(interestCompudingPeriodValidation);
            worksheet.addValidationData(interestPostingPeriodValidation);
            worksheet.addValidationData(interestCalculationValidation);
            worksheet.addValidationData(interestCalculationDaysInYearValidation);
            worksheet.addValidationData(lockinPeriodFrequencyValidation);
            worksheet.addValidationData(depositPeriodTypeValidation);
            worksheet.addValidationData(depositAmountValidation);

        } catch (RuntimeException re) {
        	re.printStackTrace();
            result.addError(re.getMessage());
            re.printStackTrace();
        }
        return result;
	}

	private void setNames(Sheet worksheet) {
		Workbook savingsWorkbook = worksheet.getWorkbook();
        ArrayList<String> officeNames = new ArrayList<String>(Arrays.asList(officeSheetPopulator.getOfficeNames()));
        List<FixedDepositProduct> products = productSheetPopulator.getProducts();

        // Office Names
        Name officeGroup = savingsWorkbook.createName();
        officeGroup.setNameName("Office");
        officeGroup.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));

        // Client and Loan Officer Names for each office
        for (Integer i = 0; i < officeNames.size(); i++) {
            Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator.getOfficeNameToBeginEndIndexesOfClients().get(i);
            Integer[] officeNameToBeginEndIndexesOfStaff = personnelSheetPopulator.getOfficeNameToBeginEndIndexesOfStaff().get(i);
            Name clientName = savingsWorkbook.createName();
            Name fieldOfficerName = savingsWorkbook.createName();
            if (officeNameToBeginEndIndexesOfStaff != null) {
                fieldOfficerName.setNameName("Staff_" + officeNames.get(i));
                fieldOfficerName.setRefersToFormula("Staff!$B$" + officeNameToBeginEndIndexesOfStaff[0] + ":$B$"
                        + officeNameToBeginEndIndexesOfStaff[1]);
            }
            if (officeNameToBeginEndIndexesOfClients != null) {
                clientName.setNameName("Client_" + officeNames.get(i));
                clientName.setRefersToFormula("Clients!$B$" + officeNameToBeginEndIndexesOfClients[0] + ":$B$"
                        + officeNameToBeginEndIndexesOfClients[1]);
            }
        }

        // Product Name
        Name productGroup = savingsWorkbook.createName();
        productGroup.setNameName("Products");
        productGroup.setRefersToFormula("Products!$B$2:$B$" + (productSheetPopulator.getProductsSize() + 1));

        // Default Interest Rate, Interest Compounding Period, Interest Posting
        // Period, Interest Calculation, Interest Calculation Days In Year,
        // Minimum Deposit, Lockin Period, Lockin Period Frequency
        // Names for each product
        for (Integer i = 0; i < products.size(); i++) {
            Name interestCompoundingPeriodName = savingsWorkbook.createName();
            Name interestPostingPeriodName = savingsWorkbook.createName();
            Name interestCalculationName = savingsWorkbook.createName();
            Name daysInYearName = savingsWorkbook.createName();
            Name lockinPeriodName = savingsWorkbook.createName();
            Name lockinPeriodFrequencyName = savingsWorkbook.createName();
            Name depositName = savingsWorkbook.createName();
    		Name minDepositName = savingsWorkbook.createName();
    		Name maxDepositName = savingsWorkbook.createName();
    		Name minDepositTermTypeName = savingsWorkbook.createName();
            
            FixedDepositProduct product = products.get(i);
            String productName = product.getName().replaceAll("[ ]", "_");
            
            interestCompoundingPeriodName.setNameName("Interest_Compouding_" + productName);
            interestPostingPeriodName.setNameName("Interest_Posting_" + productName);
            interestCalculationName.setNameName("Interest_Calculation_" + productName);
            daysInYearName.setNameName("Days_In_Year_" + productName);
            minDepositName.setNameName("Min_Deposit_" + productName);
            maxDepositName.setNameName("Max_Deposit_" + productName);
            depositName.setNameName("Deposit_" + productName);
            interestCompoundingPeriodName.setRefersToFormula("Products!$E$" + (i + 2));
            interestPostingPeriodName.setRefersToFormula("Products!$F$" + (i + 2));
            interestCalculationName.setRefersToFormula("Products!$G$" + (i + 2));
            daysInYearName.setRefersToFormula("Products!$H$" + (i + 2));
            depositName.setRefersToFormula("Products!$N$" + (i + 2));
            minDepositName.setRefersToFormula("Products!$L$" + (i + 2));
            maxDepositName.setRefersToFormula("Products!$M$" + (i + 2));
            
            if(product.getMinDepositTermType() != null) {
            	minDepositTermTypeName.setNameName("Term_Type_" + productName);
            	minDepositTermTypeName.setRefersToFormula("Products!$P$" + (i + 2));
            }
            if (product.getLockinPeriodFrequency() != null) {
                lockinPeriodName.setNameName("Lockin_Period_" + productName);
                lockinPeriodName.setRefersToFormula("Products!$I$" + (i + 2));
            }
            if (product.getLockinPeriodFrequencyType() != null) {
                lockinPeriodFrequencyName.setNameName("Lockin_Frequency_" + productName);
                lockinPeriodFrequencyName.setRefersToFormula("Products!$J$" + (i + 2));
            }
        }
	}

	private Result setDefaults(Sheet worksheet) {
		Result result = new Result();
        Workbook workbook = worksheet.getWorkbook();
        CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd-mmm");
        dateCellStyle.setDataFormat(df);
        try {
            for (Integer rowNo = 1; rowNo < 1000; rowNo++) {
                Row row = worksheet.createRow(rowNo);
                writeFormula(INTEREST_COMPOUNDING_PERIOD_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Interest_Compouding_\",$C"
                        + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"Interest_Compouding_\",$C" + (rowNo + 1) + ")))");
                writeFormula(INTEREST_POSTING_PERIOD_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Interest_Posting_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Interest_Posting_\",$C" + (rowNo + 1) + ")))");
                writeFormula(INTEREST_CALCULATION_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Interest_Calculation_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Interest_Calculation_\",$C" + (rowNo + 1) + ")))");
                writeFormula(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Days_In_Year_\",$C"
                        + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"Days_In_Year_\",$C" + (rowNo + 1) + ")))");
                writeFormula(LOCKIN_PERIOD_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Lockin_Period_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Lockin_Period_\",$C" + (rowNo + 1) + ")))");
                writeFormula(LOCKIN_PERIOD_FREQUENCY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Lockin_Frequency_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Lockin_Frequency_\",$C" + (rowNo + 1) + ")))");
                writeFormula(DEPOSIT_AMOUNT_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Deposit_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Deposit_\",$C" + (rowNo + 1) + ")))");
                writeFormula(DEPOSIT_PERIOD_FREQUENCY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Term_Type_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Term_Type_\",$C" + (rowNo + 1) + ")))");
            }
        } catch (RuntimeException re) {
        	re.printStackTrace();
            result.addError(re.getMessage());
            re.printStackTrace();
        }
        return result;
	}

	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short) 500);
        worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
        worksheet.setColumnWidth(CLIENT_NAME_COL, 4000);
        worksheet.setColumnWidth(PRODUCT_COL, 4000);
        worksheet.setColumnWidth(FIELD_OFFICER_NAME_COL, 4000);
        worksheet.setColumnWidth(SUBMITTED_ON_DATE_COL, 3200);
        worksheet.setColumnWidth(APPROVED_DATE_COL, 3200);
        worksheet.setColumnWidth(ACTIVATION_DATE_COL, 3700);
        worksheet.setColumnWidth(INTEREST_COMPOUNDING_PERIOD_COL, 3000);
        worksheet.setColumnWidth(INTEREST_POSTING_PERIOD_COL, 3000);
        worksheet.setColumnWidth(INTEREST_CALCULATION_COL, 4000);
        worksheet.setColumnWidth(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, 3000);
        worksheet.setColumnWidth(LOCKIN_PERIOD_COL, 3000);
        worksheet.setColumnWidth(LOCKIN_PERIOD_FREQUENCY_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_AMOUNT_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_PERIOD_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_PERIOD_FREQUENCY_COL, 3000);
        worksheet.setColumnWidth(EXTERNAL_ID_COL, 3000);
        
        worksheet.setColumnWidth(CHARGE_ID_1, 6000);
        worksheet.setColumnWidth(CHARGE_AMOUNT_1, 6000);
        worksheet.setColumnWidth(CHARGE_DUE_DATE_1, 6000);
        worksheet.setColumnWidth(CHARGE_ID_2, 6000);
        worksheet.setColumnWidth(CHARGE_AMOUNT_2, 6000);
        worksheet.setColumnWidth(CHARGE_DUE_DATE_2, 6000);

        worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 6000);
        worksheet.setColumnWidth(LOOKUP_ACTIVATION_DATE_COL, 6000);

        writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
        writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
        writeString(PRODUCT_COL, rowHeader, "Product*");
        writeString(FIELD_OFFICER_NAME_COL, rowHeader, "Field Officer*");
        writeString(SUBMITTED_ON_DATE_COL, rowHeader, "Submitted On*");
        writeString(APPROVED_DATE_COL, rowHeader, "Approved On*");
        writeString(ACTIVATION_DATE_COL, rowHeader, "Activation Date*");
        writeString(INTEREST_COMPOUNDING_PERIOD_COL, rowHeader, "Interest Compounding Period*");
        writeString(INTEREST_POSTING_PERIOD_COL, rowHeader, "Interest Posting Period*");
        writeString(INTEREST_CALCULATION_COL, rowHeader, "Interest Calculated*");
        writeString(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, rowHeader, "# Days in Year*");
        writeString(LOCKIN_PERIOD_COL, rowHeader, "Locked In For");
        writeString(DEPOSIT_AMOUNT_COL, rowHeader, "Deposit Amount");
        writeString(DEPOSIT_PERIOD_COL, rowHeader, "Deposit Period");
        writeString(EXTERNAL_ID_COL, rowHeader, "External Id");
        
        writeString(CHARGE_ID_1,rowHeader,"Charge Id");
        writeString(CHARGE_AMOUNT_1, rowHeader, "Charged Amount");
        writeString(CHARGE_DUE_DATE_1, rowHeader, "Charged On Date");
        writeString(CHARGE_ID_2,rowHeader,"Charge Id");
        writeString(CHARGE_AMOUNT_2, rowHeader, "Charged Amount");
        writeString(CHARGE_DUE_DATE_2, rowHeader, "Charged On Date");
        writeString(CLOSED_ON_DATE, rowHeader, "Close on Date");
        writeString(ON_ACCOUNT_CLOSURE_ID,rowHeader,"Action(Account Transfer(200) or cash(100) ");
        writeString(TO_SAVINGS_ACCOUNT_ID,rowHeader, "Transfered Account No.");
        writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Client Name");
        writeString(LOOKUP_ACTIVATION_DATE_COL, rowHeader, "Client Activation Date");
	}

}
