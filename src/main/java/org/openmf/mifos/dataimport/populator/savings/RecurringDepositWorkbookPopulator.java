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
import org.openmf.mifos.dataimport.dto.savings.RecurringDepositProduct;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.populator.AbstractWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.ClientSheetPopulator;
import org.openmf.mifos.dataimport.populator.OfficeSheetPopulator;
import org.openmf.mifos.dataimport.populator.PersonnelSheetPopulator;

public class RecurringDepositWorkbookPopulator extends AbstractWorkbookPopulator {
	
	private OfficeSheetPopulator officeSheetPopulator;
    private ClientSheetPopulator clientSheetPopulator;
    private PersonnelSheetPopulator personnelSheetPopulator;
    private RecurringDepositProductSheetPopulator productSheetPopulator;
    
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
    private static final int RECURRING_DEPOSIT_AMOUNT_COL = 13;
    private static final int DEPOSIT_PERIOD_COL = 14;
    private static final int DEPOSIT_PERIOD_FREQUENCY_COL = 15;
    private static final int DEPOSIT_FREQUENCY_COL = 16;
    private static final int DEPOSIT_FREQUENCY_TYPE_COL = 17;
    private static final int DEPOSIT_START_DATE_COL = 18;
    private static final int IS_MANDATORY_DEPOSIT_COL = 19;
    private static final int ALLOW_WITHDRAWAL_COL = 20;
    private static final int FREQ_SAME_AS_GROUP_CENTER_COL = 21;
    private static final int ADJUST_ADVANCE_PAYMENTS_COL = 22;
    private static final int EXTERNAL_ID_COL = 23;
    private static final int CHARGE_ID_1 = 28;
    private static final int CHARGE_AMOUNT_1 = 29;
    private static final int CHARGE_DUE_DATE_1 = 30;
    private static final int CHARGE_ID_2 = 33;
    private static final int CHARGE_AMOUNT_2 = 34;
    private static final int CHARGE_DUE_DATE_2 = 35;
    
   
    private static final int LOOKUP_CLIENT_NAME_COL = 31;
    private static final int LOOKUP_ACTIVATION_DATE_COL = 32;
    


    @SuppressWarnings("CPD-END")
    
    public RecurringDepositWorkbookPopulator(OfficeSheetPopulator officeSheetPopulator, ClientSheetPopulator clientSheetPopulator,
            PersonnelSheetPopulator personnelSheetPopulator, RecurringDepositProductSheetPopulator productSheetPopulator) {
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
		Sheet recurringDepositSheet = workbook.createSheet("RecurringDeposit");
        Result result = officeSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = clientSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = personnelSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = productSheetPopulator.populate(workbook);
        if (result.isSuccess()) result = setRules(recurringDepositSheet);
        if (result.isSuccess()) result = setDefaults(recurringDepositSheet);
        if (result.isSuccess())
            setClientAndGroupDateLookupTable(recurringDepositSheet, clientSheetPopulator.getClients(), null,
                    LOOKUP_CLIENT_NAME_COL, LOOKUP_ACTIVATION_DATE_COL);
        setLayout(recurringDepositSheet);
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
            		RECURRING_DEPOSIT_AMOUNT_COL, RECURRING_DEPOSIT_AMOUNT_COL);
            CellRangeAddressList depositPeriodTypeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		DEPOSIT_PERIOD_FREQUENCY_COL, DEPOSIT_PERIOD_FREQUENCY_COL);
            CellRangeAddressList depositFrequencyTypeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		DEPOSIT_FREQUENCY_TYPE_COL, DEPOSIT_FREQUENCY_TYPE_COL);
            CellRangeAddressList isMandatoryDepositRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		IS_MANDATORY_DEPOSIT_COL, IS_MANDATORY_DEPOSIT_COL);
            CellRangeAddressList allowWithdrawalRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		ALLOW_WITHDRAWAL_COL, ALLOW_WITHDRAWAL_COL);
            CellRangeAddressList adjustAdvancePaymentRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		ADJUST_ADVANCE_PAYMENTS_COL, ADJUST_ADVANCE_PAYMENTS_COL);
            CellRangeAddressList sameFreqAsGroupRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		FREQ_SAME_AS_GROUP_CENTER_COL, FREQ_SAME_AS_GROUP_CENTER_COL);
            CellRangeAddressList depositStartDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
            		DEPOSIT_START_DATE_COL, DEPOSIT_START_DATE_COL);
            
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
            DataValidationConstraint booleanConstraint = validationHelper.createExplicitListConstraint(new String[] {
                    "True", "False" });
            DataValidationConstraint depositStartDateConstraint = validationHelper.createDateConstraint(
                    DataValidationConstraint.OperatorType.BETWEEN, "=$G1", "=TODAY()", "dd/mm/yy");
            
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
            DataValidation depositFrequencyTypeValidation = validationHelper.createValidation(frequency,
            		depositFrequencyTypeRange);
            DataValidation submittedDateValidation = validationHelper.createValidation(submittedDateConstraint, submittedDateRange);
            DataValidation approvalDateValidation = validationHelper.createValidation(approvalDateConstraint, approvedDateRange);
            DataValidation activationDateValidation = validationHelper.createValidation(activationDateConstraint, activationDateRange);
            DataValidation  depositAmountValidation = validationHelper.createValidation(depositConstraint, depositAmountRange);
            DataValidation isMandatoryDepositValidation = validationHelper.createValidation(
                    booleanConstraint, isMandatoryDepositRange);
            DataValidation allowWithdrawalValidation = validationHelper.createValidation(
                    booleanConstraint, allowWithdrawalRange);
            DataValidation adjustAdvancePaymentValidation = validationHelper.createValidation(
                    booleanConstraint, adjustAdvancePaymentRange);
            DataValidation sameFreqAsGroupValidation = validationHelper.createValidation(
                    booleanConstraint, sameFreqAsGroupRange);
            DataValidation depositStartDateValidation = validationHelper.createValidation(
            		depositStartDateConstraint, depositStartDateRange);
            
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
            worksheet.addValidationData(depositFrequencyTypeValidation);
            worksheet.addValidationData(isMandatoryDepositValidation);
            worksheet.addValidationData(allowWithdrawalValidation);
            worksheet.addValidationData(adjustAdvancePaymentValidation);
            worksheet.addValidationData(sameFreqAsGroupValidation);
            worksheet.addValidationData(depositStartDateValidation);

        } catch (RuntimeException re) {
        	re.printStackTrace();
            result.addError(re.getMessage());
        }
        return result;
	}

	private void setNames(Sheet worksheet) {
		Workbook savingsWorkbook = worksheet.getWorkbook();
        ArrayList<String> officeNames = new ArrayList<String>(Arrays.asList(officeSheetPopulator.getOfficeNames()));
        List<RecurringDepositProduct> products = productSheetPopulator.getProducts();

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
    		Name allowWithdrawalName = savingsWorkbook.createName();
    		Name mandatoryDepositName = savingsWorkbook.createName();
    		Name adjustAdvancePaymentsName = savingsWorkbook.createName();
            
            RecurringDepositProduct product = products.get(i);
            String productName = product.getName().replaceAll("[ ]", "_");
            
            interestCompoundingPeriodName.setNameName("Interest_Compouding_" + productName);
            interestPostingPeriodName.setNameName("Interest_Posting_" + productName);
            interestCalculationName.setNameName("Interest_Calculation_" + productName);
            daysInYearName.setNameName("Days_In_Year_" + productName);
            minDepositName.setNameName("Min_Deposit_" + productName);
            maxDepositName.setNameName("Max_Deposit_" + productName);
            depositName.setNameName("Deposit_" + productName);
            allowWithdrawalName.setNameName("Allow_Withdrawal_" + productName);
            mandatoryDepositName.setNameName("Mandatory_Deposit_" + productName);
            adjustAdvancePaymentsName.setNameName("Adjust_Advance_" + productName);
            interestCompoundingPeriodName.setRefersToFormula("Products!$E$" + (i + 2));
            interestPostingPeriodName.setRefersToFormula("Products!$F$" + (i + 2));
            interestCalculationName.setRefersToFormula("Products!$G$" + (i + 2));
            daysInYearName.setRefersToFormula("Products!$H$" + (i + 2));
            depositName.setRefersToFormula("Products!$N$" + (i + 2));
            minDepositName.setRefersToFormula("Products!$L$" + (i + 2));
            maxDepositName.setRefersToFormula("Products!$M$" + (i + 2));
            allowWithdrawalName.setRefersToFormula("Products!$Y$" + (i + 2));
            mandatoryDepositName.setRefersToFormula("Products!$X$" + (i + 2));
            adjustAdvancePaymentsName.setRefersToFormula("Products!$Z$" + (i + 2));
            
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
                writeFormula(RECURRING_DEPOSIT_AMOUNT_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Deposit_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Deposit_\",$C" + (rowNo + 1) + ")))");
                writeFormula(DEPOSIT_PERIOD_FREQUENCY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Term_Type_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Term_Type_\",$C" + (rowNo + 1) + ")))");
                writeFormula(IS_MANDATORY_DEPOSIT_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Mandatory_Deposit_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Mandatory_Deposit_\",$C" + (rowNo + 1) + ")))");
                writeFormula(ALLOW_WITHDRAWAL_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Allow_Withdrawal_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Allow_Withdrawal_\",$C" + (rowNo + 1) + ")))");
                writeFormula(ADJUST_ADVANCE_PAYMENTS_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"Adjust_Advance_\",$C" + (rowNo + 1)
                        + "))),\"\",INDIRECT(CONCATENATE(\"Adjust_Advance_\",$C" + (rowNo + 1) + ")))");
            }
        } catch (RuntimeException re) {
        	re.printStackTrace();
            result.addError(re.getMessage());
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
        worksheet.setColumnWidth(RECURRING_DEPOSIT_AMOUNT_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_PERIOD_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_PERIOD_FREQUENCY_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_FREQUENCY_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_FREQUENCY_TYPE_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_START_DATE_COL, 3000);
        worksheet.setColumnWidth(IS_MANDATORY_DEPOSIT_COL, 3000);
        worksheet.setColumnWidth(ALLOW_WITHDRAWAL_COL, 3000);
        worksheet.setColumnWidth(ADJUST_ADVANCE_PAYMENTS_COL, 3000);
        worksheet.setColumnWidth(FREQ_SAME_AS_GROUP_CENTER_COL, 3000);
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
        writeString(RECURRING_DEPOSIT_AMOUNT_COL, rowHeader, "Recurring Deposit Amount");
        writeString(DEPOSIT_PERIOD_COL, rowHeader, "Deposit Period");
        writeString(DEPOSIT_FREQUENCY_COL, rowHeader, "Deposit Frequency");
        writeString(DEPOSIT_START_DATE_COL, rowHeader, "Deposit Start Date");
        writeString(IS_MANDATORY_DEPOSIT_COL, rowHeader, "Is Mandatory Deposit?");
        writeString(ALLOW_WITHDRAWAL_COL, rowHeader, "Allow Withdrawal?");
        writeString(ADJUST_ADVANCE_PAYMENTS_COL, rowHeader, "Adjust Advance Payments Toward Future Installments ");
        writeString(FREQ_SAME_AS_GROUP_CENTER_COL, rowHeader, "Deposit Frequency Same as Group/Center meeting");
        writeString(EXTERNAL_ID_COL, rowHeader, "External Id");
        
        writeString(CHARGE_ID_1,rowHeader,"Charge Id");
        writeString(CHARGE_AMOUNT_1, rowHeader, "Charged Amount");
        writeString(CHARGE_DUE_DATE_1, rowHeader, "Charged On Date");
        writeString(CHARGE_ID_2,rowHeader,"Charge Id");
        writeString(CHARGE_AMOUNT_2, rowHeader, "Charged Amount");
        writeString(CHARGE_DUE_DATE_2, rowHeader, "Charged On Date");

        writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Client Name");
        writeString(LOOKUP_ACTIVATION_DATE_COL, rowHeader, "Client Activation Date");
	}


}
