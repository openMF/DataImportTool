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
import org.openmf.mifos.dataimport.dto.savings.SavingsProduct;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.populator.AbstractWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.ClientSheetPopulator;
import org.openmf.mifos.dataimport.populator.GroupSheetPopulator;
import org.openmf.mifos.dataimport.populator.OfficeSheetPopulator;
import org.openmf.mifos.dataimport.populator.PersonnelSheetPopulator;

public class SavingsWorkbookPopulator extends AbstractWorkbookPopulator {

     

    private OfficeSheetPopulator officeSheetPopulator;
    private ClientSheetPopulator clientSheetPopulator;
    private GroupSheetPopulator groupSheetPopulator;
    private PersonnelSheetPopulator personnelSheetPopulator;
    private SavingsProductSheetPopulator productSheetPopulator;

    @SuppressWarnings("CPD-START")
    private static final int OFFICE_NAME_COL = 0;
    private static final int SAVINGS_TYPE_COL = 1;
    private static final int CLIENT_NAME_COL = 2;
    private static final int PRODUCT_COL = 3;
    private static final int FIELD_OFFICER_NAME_COL = 4;
    private static final int SUBMITTED_ON_DATE_COL = 5;
    private static final int APPROVED_DATE_COL = 6;
    private static final int ACTIVATION_DATE_COL = 7;
    private static final int CURRENCY_COL = 8;
    private static final int DECIMAL_PLACES_COL = 9;
    private static final int IN_MULTIPLES_OF_COL = 10;
    private static final int NOMINAL_ANNUAL_INTEREST_RATE_COL = 11;
    private static final int INTEREST_COMPOUNDING_PERIOD_COL = 12;
    private static final int INTEREST_POSTING_PERIOD_COL = 13;
    private static final int INTEREST_CALCULATION_COL = 14;
    private static final int INTEREST_CALCULATION_DAYS_IN_YEAR_COL = 15;
    private static final int MIN_OPENING_BALANCE_COL = 16;
    private static final int LOCKIN_PERIOD_COL = 17;
    private static final int LOCKIN_PERIOD_FREQUENCY_COL = 18;
    private static final int APPLY_WITHDRAWAL_FEE_FOR_TRANSFERS = 19;
    private static final int ALLOW_OVER_DRAFT_COL = 20;
    private static final int OVER_DRAFT_LIMIT_COL= 21;
    private static final int EXTERNAL_ID_COL = 22;
    private static final int LOOKUP_CLIENT_NAME_COL = 31;
    private static final int LOOKUP_ACTIVATION_DATE_COL = 32;
    private static final int CHARGE_ID_1 = 34;
    private static final int CHARGE_AMOUNT_1 = 35;
    private static final int CHARGE_DUE_DATE_1 = 36;
    private static final int CHARGE_ID_2 = 37;
    private static final int CHARGE_AMOUNT_2 = 38;
    private static final int CHARGE_DUE_DATE_2 = 39;


    @SuppressWarnings("CPD-END")
    public SavingsWorkbookPopulator(OfficeSheetPopulator officeSheetPopulator, ClientSheetPopulator clientSheetPopulator,
            GroupSheetPopulator groupSheetPopulator, PersonnelSheetPopulator personnelSheetPopulator,
            SavingsProductSheetPopulator productSheetPopulator) {
        this.officeSheetPopulator = officeSheetPopulator;
        this.clientSheetPopulator = clientSheetPopulator;
        this.groupSheetPopulator = groupSheetPopulator;
        this.personnelSheetPopulator = personnelSheetPopulator;
        this.productSheetPopulator = productSheetPopulator;
    }

    @Override
    public Result downloadAndParse() {
        Result result = officeSheetPopulator.downloadAndParse();
        if (result.isSuccess()) result = clientSheetPopulator.downloadAndParse();
        if (result.isSuccess()) result = groupSheetPopulator.downloadAndParse();
        if (result.isSuccess()) result = personnelSheetPopulator.downloadAndParse();
        if (result.isSuccess()) result = productSheetPopulator.downloadAndParse();
        return result;
    }

    @Override
    public Result populate(Workbook workbook) {
        Sheet savingsSheet = workbook.createSheet("Savings");
        Result result = officeSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = clientSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = groupSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = personnelSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = productSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = setRules(savingsSheet);
        if (result.isSuccess()) result = setDefaults(savingsSheet);
        if (result.isSuccess())
            setClientAndGroupDateLookupTable(savingsSheet, clientSheetPopulator.getClients(), groupSheetPopulator.getGroups(),
                    LOOKUP_CLIENT_NAME_COL, LOOKUP_ACTIVATION_DATE_COL);
        setLayout(savingsSheet);
        return result;
    }

    private void setLayout(Sheet worksheet) {
        Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short) 500);
        worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
        worksheet.setColumnWidth(SAVINGS_TYPE_COL, 4000);
        worksheet.setColumnWidth(CLIENT_NAME_COL, 4000);
        worksheet.setColumnWidth(PRODUCT_COL, 4000);
        worksheet.setColumnWidth(FIELD_OFFICER_NAME_COL, 4000);
        worksheet.setColumnWidth(SUBMITTED_ON_DATE_COL, 3200);
        worksheet.setColumnWidth(APPROVED_DATE_COL, 3200);
        worksheet.setColumnWidth(ACTIVATION_DATE_COL, 3700);
        worksheet.setColumnWidth(CURRENCY_COL, 2500);
        worksheet.setColumnWidth(DECIMAL_PLACES_COL, 2500);
        worksheet.setColumnWidth(IN_MULTIPLES_OF_COL, 3000);

        worksheet.setColumnWidth(NOMINAL_ANNUAL_INTEREST_RATE_COL, 3000);
        worksheet.setColumnWidth(INTEREST_COMPOUNDING_PERIOD_COL, 3000);
        worksheet.setColumnWidth(INTEREST_POSTING_PERIOD_COL, 3000);
        worksheet.setColumnWidth(INTEREST_CALCULATION_COL, 4000);
        worksheet.setColumnWidth(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, 3000);
        worksheet.setColumnWidth(MIN_OPENING_BALANCE_COL, 4000);
        worksheet.setColumnWidth(LOCKIN_PERIOD_COL, 3000);
        worksheet.setColumnWidth(LOCKIN_PERIOD_FREQUENCY_COL, 3000);
        worksheet.setColumnWidth(APPLY_WITHDRAWAL_FEE_FOR_TRANSFERS, 4000);

        worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 6000);
        worksheet.setColumnWidth(LOOKUP_ACTIVATION_DATE_COL, 6000);
        worksheet.setColumnWidth(EXTERNAL_ID_COL, 6000);
        
        worksheet.setColumnWidth(ALLOW_OVER_DRAFT_COL, 6000);
        worksheet.setColumnWidth(OVER_DRAFT_LIMIT_COL, 6000);
        
        worksheet.setColumnWidth(CHARGE_ID_1, 6000);
        worksheet.setColumnWidth(CHARGE_AMOUNT_1, 6000);
        worksheet.setColumnWidth(CHARGE_DUE_DATE_1, 6000);
        worksheet.setColumnWidth(CHARGE_ID_2, 6000);
        worksheet.setColumnWidth(CHARGE_AMOUNT_2, 6000);
        worksheet.setColumnWidth(CHARGE_DUE_DATE_2, 6000);

        writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
        writeString(SAVINGS_TYPE_COL, rowHeader, "Individual/Group*");
        writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
        writeString(PRODUCT_COL, rowHeader, "Product*");
        writeString(FIELD_OFFICER_NAME_COL, rowHeader, "Field Officer*");
        writeString(SUBMITTED_ON_DATE_COL, rowHeader, "Submitted On*");
        writeString(APPROVED_DATE_COL, rowHeader, "Approved On*");
        writeString(ACTIVATION_DATE_COL, rowHeader, "Activation Date*");
        writeString(CURRENCY_COL, rowHeader, "Currency");
        writeString(DECIMAL_PLACES_COL, rowHeader, "Decimal Places");
        writeString(IN_MULTIPLES_OF_COL, rowHeader, "In Multiples Of");
        writeString(NOMINAL_ANNUAL_INTEREST_RATE_COL, rowHeader, "Interest Rate %*");
        writeString(INTEREST_COMPOUNDING_PERIOD_COL, rowHeader, "Interest Compounding Period*");
        writeString(INTEREST_POSTING_PERIOD_COL, rowHeader, "Interest Posting Period*");
        writeString(INTEREST_CALCULATION_COL, rowHeader, "Interest Calculated*");
        writeString(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, rowHeader, "# Days in Year*");
        writeString(MIN_OPENING_BALANCE_COL, rowHeader, "Min Opening Balance");
        writeString(LOCKIN_PERIOD_COL, rowHeader, "Locked In For");
        writeString(APPLY_WITHDRAWAL_FEE_FOR_TRANSFERS, rowHeader, "Apply Withdrawal Fee For Transfers");

        writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Client Name");
        writeString(LOOKUP_ACTIVATION_DATE_COL, rowHeader, "Client Activation Date");
        writeString(EXTERNAL_ID_COL, rowHeader, "External Id");
        
        writeString(ALLOW_OVER_DRAFT_COL, rowHeader, "Is Overdraft Allowed ");
        writeString(OVER_DRAFT_LIMIT_COL, rowHeader,"  Maximum Overdraft Amount Limit ");
        
        writeString(CHARGE_ID_1,rowHeader,"Charge Id");
        writeString(CHARGE_AMOUNT_1, rowHeader, "Charged Amount");
        writeString(CHARGE_DUE_DATE_1, rowHeader, "Charged On Date");
        writeString(CHARGE_ID_2,rowHeader,"Charge Id");
        writeString(CHARGE_AMOUNT_2, rowHeader, "Charged Amount");
        writeString(CHARGE_DUE_DATE_2, rowHeader, "Charged On Date");
    }

    private Result setRules(Sheet worksheet) {
        Result result = new Result();
        try {
            CellRangeAddressList officeNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    OFFICE_NAME_COL, OFFICE_NAME_COL);
            CellRangeAddressList savingsTypeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
                    SAVINGS_TYPE_COL, SAVINGS_TYPE_COL);
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
            CellRangeAddressList applyWithdrawalFeeForTransfersRange = new CellRangeAddressList(1,
                    SpreadsheetVersion.EXCEL97.getLastRowIndex(), APPLY_WITHDRAWAL_FEE_FOR_TRANSFERS, APPLY_WITHDRAWAL_FEE_FOR_TRANSFERS);
            CellRangeAddressList allowOverdraftRange = new CellRangeAddressList(1,SpreadsheetVersion.EXCEL97.getLastRowIndex(),ALLOW_OVER_DRAFT_COL,ALLOW_OVER_DRAFT_COL);
           
             

            DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet) worksheet);

            setNames(worksheet);

            DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
            DataValidationConstraint savingsTypeConstraint = validationHelper.createExplicitListConstraint(new String[] { "Individual",
                    "Group" });
            DataValidationConstraint clientNameConstraint = validationHelper
                    .createFormulaListConstraint("IF($B1=\"Individual\",INDIRECT(CONCATENATE(\"Client_\",$A1)),INDIRECT(CONCATENATE(\"Group_\",$A1)))");
            DataValidationConstraint productNameConstraint = validationHelper.createFormulaListConstraint("Products");
            DataValidationConstraint fieldOfficerNameConstraint = validationHelper
                    .createFormulaListConstraint("INDIRECT(CONCATENATE(\"Staff_\",$A1))");
            DataValidationConstraint submittedDateConstraint = validationHelper.createDateConstraint(
                    DataValidationConstraint.OperatorType.BETWEEN, "=VLOOKUP($C1,$AF$2:$AG$"
                            + (clientSheetPopulator.getClientsSize() + groupSheetPopulator.getGroupsSize() + 1) + ",2,FALSE)", "=TODAY()",
                    "dd/mm/yy");
            DataValidationConstraint approvalDateConstraint = validationHelper.createDateConstraint(
                    DataValidationConstraint.OperatorType.BETWEEN, "=$F1", "=TODAY()", "dd/mm/yy");
            DataValidationConstraint activationDateConstraint = validationHelper.createDateConstraint(
                    DataValidationConstraint.OperatorType.BETWEEN, "=$G1", "=TODAY()", "dd/mm/yy");
            DataValidationConstraint interestCompudingPeriodConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "Daily", "Monthly", "Semi-Annual" });
            DataValidationConstraint interestPostingPeriodConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "Monthly", "Quarterly", "Annually", "BiAnnual" });
            DataValidationConstraint interestCalculationConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "Daily Balance", "Average Daily Balance" });
            DataValidationConstraint interestCalculationDaysInYearConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "360 Days", "365 Days" });
            DataValidationConstraint lockinPeriodFrequencyConstraint = validationHelper.createExplicitListConstraint(new String[] { "Days",
                    "Weeks", "Months", "Years" });
            DataValidationConstraint applyWithdrawalFeeForTransferConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "True", "False" });
            DataValidationConstraint allowOverdraftConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "True", "False" });

            DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
            DataValidation savingsTypeValidation = validationHelper.createValidation(savingsTypeConstraint, savingsTypeRange);
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
            DataValidation lockinPeriodFrequencyValidation = validationHelper.createValidation(lockinPeriodFrequencyConstraint,
                    lockinPeriodFrequencyRange);
            DataValidation applyWithdrawalFeeForTransferValidation = validationHelper.createValidation(
                    applyWithdrawalFeeForTransferConstraint, applyWithdrawalFeeForTransfersRange);
            DataValidation submittedDateValidation = validationHelper.createValidation(submittedDateConstraint, submittedDateRange);
            DataValidation approvalDateValidation = validationHelper.createValidation(approvalDateConstraint, approvedDateRange);
            DataValidation activationDateValidation = validationHelper.createValidation(activationDateConstraint, activationDateRange);
            DataValidation allowOverdraftValidation = validationHelper.createValidation(
                    allowOverdraftConstraint, allowOverdraftRange);
            
            worksheet.addValidationData(officeValidation);
            worksheet.addValidationData(savingsTypeValidation);
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
            worksheet.addValidationData(applyWithdrawalFeeForTransferValidation);
            worksheet.addValidationData(allowOverdraftValidation);

        } catch (RuntimeException re) {
            result.addError(re.getMessage());
        }
        return result;
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
                writeFormula(CURRENCY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Currency_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Currency_\",$D" + (rowNo + 1) + ")))");
                writeFormula(DECIMAL_PLACES_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Decimal_Places_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Decimal_Places_\",$D" + (rowNo + 1) + ")))");
                writeFormula(IN_MULTIPLES_OF_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"In_Multiples_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"In_Multiples_\",$D" + (rowNo + 1) + ")))");
                writeFormula(NOMINAL_ANNUAL_INTEREST_RATE_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Interest_Rate_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Interest_Rate_\",$D" + (rowNo + 1) + ")))");
                writeFormula(INTEREST_COMPOUNDING_PERIOD_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Interest_Compouding_\",$D"
                        + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"Interest_Compouding_\",$D" + (rowNo + 1) + ")))");
                writeFormula(INTEREST_POSTING_PERIOD_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Interest_Posting_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Interest_Posting_\",$D" + (rowNo + 1) + ")))");
                writeFormula(INTEREST_CALCULATION_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Interest_Calculation_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Interest_Calculation_\",$D" + (rowNo + 1) + ")))");
                writeFormula(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Days_In_Year_\",$D"
                        + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"Days_In_Year_\",$D" + (rowNo + 1) + ")))");
                writeFormula(MIN_OPENING_BALANCE_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Min_Balance_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Min_Balance_\",$D" + (rowNo + 1) + ")))");
                writeFormula(LOCKIN_PERIOD_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Lockin_Period_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Lockin_Period_\",$D" + (rowNo + 1) + ")))");
                writeFormula(LOCKIN_PERIOD_FREQUENCY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Lockin_Frequency_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Lockin_Frequency_\",$D" + (rowNo + 1) + ")))");
                writeFormula(APPLY_WITHDRAWAL_FEE_FOR_TRANSFERS, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Withdrawal_Fee_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Withdrawal_Fee_\",$D" + (rowNo + 1) + ")))");
                writeFormula(ALLOW_OVER_DRAFT_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Overdraft_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Overdraft_\",$D" + (rowNo + 1) + ")))");
                writeFormula(OVER_DRAFT_LIMIT_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Overdraft_Limit_\",$D" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Overdraft_Limit_\",$D" + (rowNo + 1) + ")))");
            }
        } catch (RuntimeException re) {
             
            result.addError(re.getMessage());
        }
        return result;
    }

    private void setNames(Sheet worksheet) {
        Workbook savingsWorkbook = worksheet.getWorkbook();
        ArrayList<String> officeNames = new ArrayList<String>(Arrays.asList(officeSheetPopulator.getOfficeNames()));
        List<SavingsProduct> products = productSheetPopulator.getProducts();

        // Office Names
        Name officeGroup = savingsWorkbook.createName();
        officeGroup.setNameName("Office");
        officeGroup.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));

        // Client and Loan Officer Names for each office
        for (Integer i = 0; i < officeNames.size(); i++) {
            Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator.getOfficeNameToBeginEndIndexesOfClients().get(i);
            Integer[] officeNameToBeginEndIndexesOfStaff = personnelSheetPopulator.getOfficeNameToBeginEndIndexesOfStaff().get(i);
            Integer[] officeNameToBeginEndIndexesOfGroups = groupSheetPopulator.getOfficeNameToBeginEndIndexesOfGroups().get(i);
            Name clientName = savingsWorkbook.createName();
            Name fieldOfficerName = savingsWorkbook.createName();
            Name groupName = savingsWorkbook.createName();
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
            if (officeNameToBeginEndIndexesOfGroups != null) {
                groupName.setNameName("Group_" + officeNames.get(i));
                groupName.setRefersToFormula("Groups!$B$" + officeNameToBeginEndIndexesOfGroups[0] + ":$B$"
                        + officeNameToBeginEndIndexesOfGroups[1]);
            }
        }

        // Product Name
        Name productGroup = savingsWorkbook.createName();
        productGroup.setNameName("Products");
        productGroup.setRefersToFormula("Products!$B$2:$B$" + (productSheetPopulator.getProductsSize() + 1));

        // Default Interest Rate, Interest Compounding Period, Interest Posting
        // Period, Interest Calculation, Interest Calculation Days In Year,
        // Minimum Opening Balance, Lockin Period, Lockin Period Frequency,
        // Withdrawal Fee Amount, Withdrawal Fee Type, Annual Fee, Annual Fee on
        // Date
        // Names for each product
        for (Integer i = 0; i < products.size(); i++) {
            Name interestRateName = savingsWorkbook.createName();
            Name interestCompoundingPeriodName = savingsWorkbook.createName();
            Name interestPostingPeriodName = savingsWorkbook.createName();
            Name interestCalculationName = savingsWorkbook.createName();
            Name daysInYearName = savingsWorkbook.createName();
            Name minOpeningBalanceName = savingsWorkbook.createName();
            Name lockinPeriodName = savingsWorkbook.createName();
            Name lockinPeriodFrequencyName = savingsWorkbook.createName();
            Name currencyName = savingsWorkbook.createName();
            Name decimalPlacesName = savingsWorkbook.createName();
            Name inMultiplesOfName = savingsWorkbook.createName();
            Name withdrawalFeeName = savingsWorkbook.createName();
            Name allowOverdraftName = savingsWorkbook.createName();
            Name overdraftLimitName = savingsWorkbook.createName();
            SavingsProduct product = products.get(i);
            String productName = product.getName().replaceAll("[ ]", "_");
            if (product.getNominalAnnualInterestRate() != null) {
                interestRateName.setNameName("Interest_Rate_" + productName);
                interestRateName.setRefersToFormula("Products!$C$" + (i + 2));
            }
            interestCompoundingPeriodName.setNameName("Interest_Compouding_" + productName);
            interestPostingPeriodName.setNameName("Interest_Posting_" + productName);
            interestCalculationName.setNameName("Interest_Calculation_" + productName);
            daysInYearName.setNameName("Days_In_Year_" + productName);
            currencyName.setNameName("Currency_" + productName);
            decimalPlacesName.setNameName("Decimal_Places_" + productName);
            withdrawalFeeName.setNameName("Withdrawal_Fee_" + productName);
            allowOverdraftName.setNameName("Overdraft_" + productName);
            
            interestCompoundingPeriodName.setRefersToFormula("Products!$D$" + (i + 2));
            interestPostingPeriodName.setRefersToFormula("Products!$E$" + (i + 2));
            interestCalculationName.setRefersToFormula("Products!$F$" + (i + 2));
            daysInYearName.setRefersToFormula("Products!$G$" + (i + 2));
            currencyName.setRefersToFormula("Products!$K$" + (i + 2));
            decimalPlacesName.setRefersToFormula("Products!$L$" + (i + 2));
            withdrawalFeeName.setRefersToFormula("Products!$N$" + (i + 2));
            allowOverdraftName.setRefersToFormula("Products!$O$" + (i + 2));
            if (product.getOverdraftLimit() != null) {
            	overdraftLimitName.setNameName("Overdraft_Limit_" + productName);
            	overdraftLimitName.setRefersToFormula("Products!$P$" + (i + 2));
            }
            if (product.getMinRequiredOpeningBalance() != null) {
                minOpeningBalanceName.setNameName("Min_Balance_" + productName);
                minOpeningBalanceName.setRefersToFormula("Products!$H$" + (i + 2));
            }
            if (product.getLockinPeriodFrequency() != null) {
                lockinPeriodName.setNameName("Lockin_Period_" + productName);
                lockinPeriodName.setRefersToFormula("Products!$I$" + (i + 2));
            }
            if (product.getLockinPeriodFrequencyType() != null) {
                lockinPeriodFrequencyName.setNameName("Lockin_Frequency_" + productName);
                lockinPeriodFrequencyName.setRefersToFormula("Products!$J$" + (i + 2));
            }
            if (product.getCurrency().getInMultiplesOf() != null) {
                inMultiplesOfName.setNameName("In_Multiples_" + productName);
                inMultiplesOfName.setRefersToFormula("Products!$M$" + (i + 2));
            }
        }
    }

}
