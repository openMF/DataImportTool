package org.openmf.mifos.dataimport.populator.loan;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.openmf.mifos.dataimport.dto.loan.LoanProduct;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.populator.AbstractWorkbookPopulator;
import org.openmf.mifos.dataimport.populator.ClientSheetPopulator;
import org.openmf.mifos.dataimport.populator.ExtrasSheetPopulator;
import org.openmf.mifos.dataimport.populator.GroupSheetPopulator;
import org.openmf.mifos.dataimport.populator.OfficeSheetPopulator;
import org.openmf.mifos.dataimport.populator.PersonnelSheetPopulator;

public class LoanWorkbookPopulator extends AbstractWorkbookPopulator {
	
	private OfficeSheetPopulator officeSheetPopulator;
	private ClientSheetPopulator clientSheetPopulator;
	private GroupSheetPopulator groupSheetPopulator;
	private PersonnelSheetPopulator personnelSheetPopulator;
	private LoanProductSheetPopulator productSheetPopulator;
	private ExtrasSheetPopulator extrasSheetPopulator;
	
	@SuppressWarnings("CPD-START")
	private static final int OFFICE_NAME_COL = 0;
	private static final int LOAN_TYPE_COL = 1;
    private static final int CLIENT_NAME_COL = 2;
    private static final int PRODUCT_COL = 3;
    private static final int LOAN_OFFICER_NAME_COL = 4;
    private static final int SUBMITTED_ON_DATE_COL = 5;
    private static final int APPROVED_DATE_COL = 6;
    private static final int DISBURSED_DATE_COL = 7;
    private static final int DISBURSED_PAYMENT_TYPE_COL = 8;
    private static final int FUND_NAME_COL = 9;   
    private static final int PRINCIPAL_COL = 10;
    private static final int NO_OF_REPAYMENTS_COL = 11;
    private static final int REPAID_EVERY_COL = 12;
    private static final int REPAID_EVERY_FREQUENCY_COL = 13;
    private static final int LOAN_TERM_COL = 14;
    private static final int LOAN_TERM_FREQUENCY_COL = 15;
    private static final int NOMINAL_INTEREST_RATE_COL = 16;
    private static final int NOMINAL_INTEREST_RATE_FREQUENCY_COL = 17;
    private static final int AMORTIZATION_COL = 18;
    private static final int INTEREST_METHOD_COL = 19;
    private static final int INTEREST_CALCULATION_PERIOD_COL = 20;
    private static final int ARREARS_TOLERANCE_COL = 21;
    private static final int REPAYMENT_STRATEGY_COL = 22;
    private static final int GRACE_ON_PRINCIPAL_PAYMENT_COL = 23;
    private static final int GRACE_ON_INTEREST_PAYMENT_COL = 24;
    private static final int GRACE_ON_INTEREST_CHARGED_COL = 25;
    private static final int INTEREST_CHARGED_FROM_COL = 26;
    private static final int FIRST_REPAYMENT_COL = 27;
    private static final int TOTAL_AMOUNT_REPAID_COL = 28;
    private static final int LAST_REPAYMENT_DATE_COL = 29;
    private static final int REPAYMENT_TYPE_COL = 30;
    private static final int LOOKUP_CLIENT_NAME_COL = 42;
    private static final int LOOKUP_ACTIVATION_DATE_COL = 43;
    @SuppressWarnings("CPD-END")
	
	public LoanWorkbookPopulator(OfficeSheetPopulator officeSheetPopulator, ClientSheetPopulator clientSheetPopulator,
			GroupSheetPopulator groupSheetPopulator, PersonnelSheetPopulator personnelSheetPopulator,
			LoanProductSheetPopulator productSheetPopulator, ExtrasSheetPopulator extrasSheetPopulator) {
    	this.officeSheetPopulator = officeSheetPopulator;
    	this.clientSheetPopulator = clientSheetPopulator;
    	this.groupSheetPopulator = groupSheetPopulator;
		this.personnelSheetPopulator = personnelSheetPopulator;
    	this.productSheetPopulator = productSheetPopulator;
    	this.extrasSheetPopulator = extrasSheetPopulator;
    }
	
	    @Override
	    public Result downloadAndParse() {
	    	Result result =  officeSheetPopulator.downloadAndParse();
	    	if(result.isSuccess())
	    		result = clientSheetPopulator.downloadAndParse();
	    	if(result.isSuccess())
	    		result = groupSheetPopulator.downloadAndParse();
	    	if(result.isSuccess())
	    		result = personnelSheetPopulator.downloadAndParse();
	    	if(result.isSuccess())
	    		result = productSheetPopulator.downloadAndParse();
	    	if(result.isSuccess())
	    		result = extrasSheetPopulator.downloadAndParse();
	    	return result;
	    }

	    @Override
	    public Result populate(Workbook workbook) {
	    	Sheet loanSheet = workbook.createSheet("Loans");
	    	Result result = officeSheetPopulator.populate(workbook);
	    	if(result.isSuccess())
	    		result = clientSheetPopulator.populate(workbook);
	    	if(result.isSuccess())
	    		result = groupSheetPopulator.populate(workbook);
	    	if(result.isSuccess()) 
	    		result = personnelSheetPopulator.populate(workbook);
	    	if(result.isSuccess())
	    		result = productSheetPopulator.populate(workbook);
	    	if(result.isSuccess())
	    		result = extrasSheetPopulator.populate(workbook);
	    	setLayout(loanSheet);
	    	if(result.isSuccess())
	            result = setRules(loanSheet);
	    	if(result.isSuccess()) 
	            result = setDefaults(loanSheet);
	    	setClientAndGroupDateLookupTable(loanSheet, clientSheetPopulator.getClients(), groupSheetPopulator.getGroups(),
	    			LOOKUP_CLIENT_NAME_COL, LOOKUP_ACTIVATION_DATE_COL);
	        return result;
	    }
	    
	    private void setLayout(Sheet worksheet) {
	    	Row rowHeader = worksheet.createRow(0);
	        rowHeader.setHeight((short)500);
	        worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
	        worksheet.setColumnWidth(LOAN_TYPE_COL, 3500);
            worksheet.setColumnWidth(CLIENT_NAME_COL, 4000);
            worksheet.setColumnWidth(PRODUCT_COL, 4000);
            worksheet.setColumnWidth(LOAN_OFFICER_NAME_COL, 4000);
            worksheet.setColumnWidth(SUBMITTED_ON_DATE_COL, 3200);
            worksheet.setColumnWidth(APPROVED_DATE_COL, 3200);
            worksheet.setColumnWidth(DISBURSED_DATE_COL, 3700);
            worksheet.setColumnWidth(DISBURSED_PAYMENT_TYPE_COL, 4000);
            worksheet.setColumnWidth(FUND_NAME_COL, 3000);
            worksheet.setColumnWidth(PRINCIPAL_COL, 3000);
            worksheet.setColumnWidth(LOAN_TERM_COL, 2000);
            worksheet.setColumnWidth(LOAN_TERM_FREQUENCY_COL, 2500);
            worksheet.setColumnWidth(NO_OF_REPAYMENTS_COL, 3800);
            worksheet.setColumnWidth(REPAID_EVERY_COL, 2000);
            worksheet.setColumnWidth(REPAID_EVERY_FREQUENCY_COL, 2000);
            worksheet.setColumnWidth(NOMINAL_INTEREST_RATE_COL, 2000);
            worksheet.setColumnWidth(NOMINAL_INTEREST_RATE_FREQUENCY_COL, 3000);
            worksheet.setColumnWidth(AMORTIZATION_COL, 6000);
            worksheet.setColumnWidth(INTEREST_METHOD_COL, 4000);
            worksheet.setColumnWidth(INTEREST_CALCULATION_PERIOD_COL, 4000);
            worksheet.setColumnWidth(ARREARS_TOLERANCE_COL, 4000);
            worksheet.setColumnWidth(REPAYMENT_STRATEGY_COL, 4700);
            worksheet.setColumnWidth(GRACE_ON_PRINCIPAL_PAYMENT_COL, 3500);
            worksheet.setColumnWidth(GRACE_ON_INTEREST_PAYMENT_COL, 3500);
            worksheet.setColumnWidth(GRACE_ON_INTEREST_CHARGED_COL, 3500);
            worksheet.setColumnWidth(INTEREST_CHARGED_FROM_COL, 4000);
            worksheet.setColumnWidth(FIRST_REPAYMENT_COL, 4700);
            worksheet.setColumnWidth(TOTAL_AMOUNT_REPAID_COL, 3500);
            worksheet.setColumnWidth(LAST_REPAYMENT_DATE_COL, 3000);
            worksheet.setColumnWidth(REPAYMENT_TYPE_COL, 4300);
            worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 6000);
            worksheet.setColumnWidth(LOOKUP_ACTIVATION_DATE_COL, 6000);
            writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
            writeString(LOAN_TYPE_COL, rowHeader, "Loan Type*");
            writeString(CLIENT_NAME_COL, rowHeader, "Client/Group Name*");
            writeString(PRODUCT_COL, rowHeader, "Product*");
            writeString(LOAN_OFFICER_NAME_COL, rowHeader, "Loan Officer*");
            writeString(SUBMITTED_ON_DATE_COL, rowHeader, "Submitted On*");
            writeString(APPROVED_DATE_COL, rowHeader, "Approved On*");
            writeString(DISBURSED_DATE_COL, rowHeader, "Disbursed Date*");
            writeString(DISBURSED_PAYMENT_TYPE_COL, rowHeader, "Payment Type*");
            writeString(FUND_NAME_COL, rowHeader, "Fund Name");
            writeString(PRINCIPAL_COL, rowHeader, "Principal*");
            writeString(LOAN_TERM_COL, rowHeader, "Loan Term*");
            writeString(NO_OF_REPAYMENTS_COL, rowHeader, "# of Repayments*");
            writeString(REPAID_EVERY_COL, rowHeader, "Repaid Every*");
            writeString(NOMINAL_INTEREST_RATE_COL, rowHeader, "Nominal Interest %*");
            writeString(AMORTIZATION_COL, rowHeader, "Amortization*");
            writeString(INTEREST_METHOD_COL, rowHeader, "Interest Method*");
            writeString(INTEREST_CALCULATION_PERIOD_COL, rowHeader, "Interest Calculation Period*");
            writeString(ARREARS_TOLERANCE_COL, rowHeader, "Arrears Tolerance");
            writeString(REPAYMENT_STRATEGY_COL, rowHeader, "Repayment Strategy*");
            writeString(GRACE_ON_PRINCIPAL_PAYMENT_COL, rowHeader, "Grace-Principal Payment");
            writeString(GRACE_ON_INTEREST_PAYMENT_COL, rowHeader, "Grace-Interest Payment");
            writeString(GRACE_ON_INTEREST_CHARGED_COL, rowHeader, "Interest-Free Period(s)");
            writeString(INTEREST_CHARGED_FROM_COL, rowHeader, "Interest Charged From");
            writeString(FIRST_REPAYMENT_COL, rowHeader, "First Repayment On");
            writeString(TOTAL_AMOUNT_REPAID_COL, rowHeader, "Amount Repaid");
            writeString(LAST_REPAYMENT_DATE_COL, rowHeader, "Date-Last Repayment");
            writeString(REPAYMENT_TYPE_COL, rowHeader, "Repayment Type");
            writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Client Name");
            writeString(LOOKUP_ACTIVATION_DATE_COL, rowHeader, "Client Activation Date");
            CellStyle borderStyle = worksheet.getWorkbook().createCellStyle();
            CellStyle doubleBorderStyle = worksheet.getWorkbook().createCellStyle();
            borderStyle.setBorderBottom(CellStyle.BORDER_THIN);
            doubleBorderStyle.setBorderBottom(CellStyle.BORDER_THIN);
            doubleBorderStyle.setBorderRight(CellStyle.BORDER_THICK);
            for(int colNo = 0; colNo < 30; colNo++) {
            	Cell cell = rowHeader.getCell(colNo);
            	if(cell == null)
            		rowHeader.createCell(colNo);
            	rowHeader.getCell(colNo).setCellStyle(borderStyle);
            }
            rowHeader.getCell(FIRST_REPAYMENT_COL).setCellStyle(doubleBorderStyle);
            rowHeader.getCell(REPAYMENT_TYPE_COL).setCellStyle(doubleBorderStyle);
	    }
	    
	    private Result setRules(Sheet worksheet) {
	    	Result result = new Result();
	    	try {
	    		CellRangeAddressList officeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), OFFICE_NAME_COL, OFFICE_NAME_COL);
	    		CellRangeAddressList loanTypeRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_TYPE_COL, LOAN_TYPE_COL);
	        	CellRangeAddressList clientNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), CLIENT_NAME_COL, CLIENT_NAME_COL);
	        	CellRangeAddressList productNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), PRODUCT_COL, PRODUCT_COL);
	        	CellRangeAddressList loanOfficerRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_OFFICER_NAME_COL, LOAN_OFFICER_NAME_COL);
	        	CellRangeAddressList submittedDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), SUBMITTED_ON_DATE_COL, SUBMITTED_ON_DATE_COL);
	        	CellRangeAddressList fundNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), FUND_NAME_COL, FUND_NAME_COL);
	        	CellRangeAddressList principalRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), PRINCIPAL_COL, PRINCIPAL_COL);
	        	CellRangeAddressList noOfRepaymentsRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), NO_OF_REPAYMENTS_COL, NO_OF_REPAYMENTS_COL);
	        	CellRangeAddressList repaidFrequencyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAID_EVERY_FREQUENCY_COL, REPAID_EVERY_FREQUENCY_COL);
	        	CellRangeAddressList loanTermRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_TERM_COL, LOAN_TERM_COL);
	        	CellRangeAddressList loanTermFrequencyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_TERM_FREQUENCY_COL, LOAN_TERM_FREQUENCY_COL);
	        	CellRangeAddressList interestFrequencyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), NOMINAL_INTEREST_RATE_FREQUENCY_COL, NOMINAL_INTEREST_RATE_FREQUENCY_COL);
	        	CellRangeAddressList interestRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), NOMINAL_INTEREST_RATE_COL, NOMINAL_INTEREST_RATE_COL);
	        	CellRangeAddressList amortizationRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), AMORTIZATION_COL, AMORTIZATION_COL);
	        	CellRangeAddressList interestMethodRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), INTEREST_METHOD_COL, INTEREST_METHOD_COL);
	        	CellRangeAddressList intrestCalculationPeriodRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), INTEREST_CALCULATION_PERIOD_COL, INTEREST_CALCULATION_PERIOD_COL);
	        	CellRangeAddressList repaymentStrategyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAYMENT_STRATEGY_COL, REPAYMENT_STRATEGY_COL);
	        	CellRangeAddressList arrearsToleranceRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), ARREARS_TOLERANCE_COL, ARREARS_TOLERANCE_COL);
	        	CellRangeAddressList graceOnPrincipalPaymentRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), GRACE_ON_PRINCIPAL_PAYMENT_COL, GRACE_ON_PRINCIPAL_PAYMENT_COL);
	        	CellRangeAddressList graceOnInterestPaymentRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), GRACE_ON_INTEREST_PAYMENT_COL, GRACE_ON_INTEREST_PAYMENT_COL);
	        	CellRangeAddressList graceOnInterestChargedRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), GRACE_ON_INTEREST_CHARGED_COL, GRACE_ON_INTEREST_CHARGED_COL);
	        	CellRangeAddressList approvedDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), APPROVED_DATE_COL, APPROVED_DATE_COL);
	        	CellRangeAddressList disbursedDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), DISBURSED_DATE_COL, DISBURSED_DATE_COL);
	        	CellRangeAddressList paymentTypeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), DISBURSED_PAYMENT_TYPE_COL, DISBURSED_PAYMENT_TYPE_COL);
	        	CellRangeAddressList repaymentTypeRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAYMENT_TYPE_COL, REPAYMENT_TYPE_COL);
	        	CellRangeAddressList lastrepaymentDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LAST_REPAYMENT_DATE_COL, LAST_REPAYMENT_DATE_COL);
	        	
	        	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
	        	
	        	setNames(worksheet);
	        	
	        	DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
	        	DataValidationConstraint loanTypeConstraint = validationHelper.createExplicitListConstraint(new String[] {"Individual","Group"});
	        	DataValidationConstraint clientNameConstraint = validationHelper.createFormulaListConstraint("IF($B1=\"Individual\",INDIRECT(CONCATENATE(\"Client_\",$A1)),INDIRECT(CONCATENATE(\"Group_\",$A1)))");
	        	DataValidationConstraint productNameConstraint = validationHelper.createFormulaListConstraint("Products");
	        	DataValidationConstraint loanOfficerNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Staff_\",$A1))");
	        	DataValidationConstraint submittedDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=IF(INDIRECT(CONCATENATE(\"START_DATE_\",$D1))>VLOOKUP($C1,$AQ$2:$AR$" + (clientSheetPopulator.getClientsSize() + groupSheetPopulator.getGroupsSize() + 1) + ",2,FALSE),INDIRECT(CONCATENATE(\"START_DATE_\",$D1)),VLOOKUP($C1,$AQ$2:$AR$" + (clientSheetPopulator.getClientsSize() + groupSheetPopulator.getGroupsSize() + 1) + ",2,FALSE))", "=TODAY()", "dd/mm/yy");
	        	DataValidationConstraint approvalDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=$F1", "=TODAY()", "dd/mm/yy");
	        	DataValidationConstraint disbursedDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=$G1", "=TODAY()", "dd/mm/yy");
	        	DataValidationConstraint paymentTypeConstraint = validationHelper.createFormulaListConstraint("PaymentTypes");
	        	DataValidationConstraint fundNameConstraint = validationHelper.createFormulaListConstraint("Funds");
	        	DataValidationConstraint principalConstraint = validationHelper.createDecimalConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=INDIRECT(CONCATENATE(\"MIN_PRINCIPAL_\",$D1))", "=INDIRECT(CONCATENATE(\"MAX_PRINCIPAL_\",$D1))");
	        	DataValidationConstraint noOfRepaymentsConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=INDIRECT(CONCATENATE(\"MIN_REPAYMENT_\",$D1))", "=INDIRECT(CONCATENATE(\"MAX_REPAYMENT_\",$D1))");
	        	DataValidationConstraint frequencyConstraint = validationHelper.createExplicitListConstraint(new String[] {"Days","Weeks","Months"});
	        	DataValidationConstraint loanTermConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "=$L1*$M1", null);
	        	DataValidationConstraint interestFrequencyConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"INTEREST_FREQUENCY_\",$D1))");
	        	DataValidationConstraint interestConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=INDIRECT(CONCATENATE(\"MIN_INTEREST_\",$D1))", "=INDIRECT(CONCATENATE(\"MAX_INTEREST_\",$D1))");
	        	DataValidationConstraint amortizationConstraint = validationHelper.createExplicitListConstraint(new String[] {"Equal principal payments","Equal installments"});
	        	DataValidationConstraint interestMethodConstraint = validationHelper.createExplicitListConstraint(new String[] {"Flat","Declining Balance"});
	        	DataValidationConstraint interestCalculationPeriodConstraint = validationHelper.createExplicitListConstraint(new String[] {"Daily","Same as repayment period"});
	        	DataValidationConstraint repaymentStrategyConstraint = validationHelper.createExplicitListConstraint(new String[] {"Mifos style","Heavensfamily","Creocore","RBI (India)","Principal Interest Penalties Fees Order","Interest Principal Penalties Fees Order"});
	        	DataValidationConstraint arrearsToleranceConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null);
	        	DataValidationConstraint graceOnPrincipalPaymentConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null);
	        	DataValidationConstraint graceOnInterestPaymentConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null);
	        	DataValidationConstraint graceOnInterestChargedConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null);
	        	DataValidationConstraint lastRepaymentDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=$H1", "=TODAY()", "dd/mm/yy");
	        	
	        	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
	        	DataValidation loanTypeValidation = validationHelper.createValidation(loanTypeConstraint, loanTypeRange);
	        	DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
	        	DataValidation productNameValidation = validationHelper.createValidation(productNameConstraint, productNameRange);
	        	DataValidation loanOfficerValidation = validationHelper.createValidation(loanOfficerNameConstraint, loanOfficerRange);
	        	DataValidation fundNameValidation = validationHelper.createValidation(fundNameConstraint, fundNameRange);
	        	DataValidation repaidFrequencyValidation = validationHelper.createValidation(frequencyConstraint, repaidFrequencyRange);
	        	DataValidation loanTermFrequencyValidation = validationHelper.createValidation(frequencyConstraint, loanTermFrequencyRange);
	        	DataValidation amortizationValidation = validationHelper.createValidation(amortizationConstraint, amortizationRange);
	        	DataValidation interestMethodValidation = validationHelper.createValidation(interestMethodConstraint, interestMethodRange);
	        	DataValidation interestCalculationPeriodValidation = validationHelper.createValidation(interestCalculationPeriodConstraint, intrestCalculationPeriodRange);
	        	DataValidation repaymentStrategyValidation = validationHelper.createValidation(repaymentStrategyConstraint, repaymentStrategyRange);
	        	DataValidation paymentTypeValidation = validationHelper.createValidation(paymentTypeConstraint, paymentTypeRange);
	        	DataValidation repaymentTypeValidation = validationHelper.createValidation(paymentTypeConstraint, repaymentTypeRange);
	        	DataValidation submittedDateValidation = validationHelper.createValidation(submittedDateConstraint, submittedDateRange);
	        	DataValidation approvalDateValidation = validationHelper.createValidation(approvalDateConstraint, approvedDateRange);
	        	DataValidation disbursedDateValidation = validationHelper.createValidation(disbursedDateConstraint, disbursedDateRange);
	        	DataValidation lastRepaymentDateValidation = validationHelper.createValidation(lastRepaymentDateConstraint, lastrepaymentDateRange);
	        	DataValidation principalValidation = validationHelper.createValidation(principalConstraint, principalRange);
	        	DataValidation loanTermValidation = validationHelper.createValidation(loanTermConstraint, loanTermRange);
	        	DataValidation noOfRepaymentsValidation = validationHelper.createValidation(noOfRepaymentsConstraint, noOfRepaymentsRange);
	        	DataValidation interestValidation = validationHelper.createValidation(interestConstraint, interestRange);
	        	DataValidation arrearsToleranceValidation = validationHelper.createValidation(arrearsToleranceConstraint, arrearsToleranceRange);
	        	DataValidation graceOnPrincipalPaymentValidation = validationHelper.createValidation(graceOnPrincipalPaymentConstraint, graceOnPrincipalPaymentRange);
	        	DataValidation graceOnInterestPaymentValidation = validationHelper.createValidation(graceOnInterestPaymentConstraint, graceOnInterestPaymentRange);
	        	DataValidation graceOnInterestChargedValidation = validationHelper.createValidation(graceOnInterestChargedConstraint, graceOnInterestChargedRange);
	        	DataValidation interestFrequencyValidation = validationHelper.createValidation(interestFrequencyConstraint, interestFrequencyRange);
	        	interestFrequencyValidation.setSuppressDropDownArrow(true);
	        	
	        	
	        	
	        	worksheet.addValidationData(officeValidation);
	        	worksheet.addValidationData(loanTypeValidation);
	            worksheet.addValidationData(clientValidation);
	            worksheet.addValidationData(productNameValidation);
	            worksheet.addValidationData(loanOfficerValidation);
	            worksheet.addValidationData(submittedDateValidation);
	            worksheet.addValidationData(approvalDateValidation);
	            worksheet.addValidationData(disbursedDateValidation);
	            worksheet.addValidationData(paymentTypeValidation);
	            worksheet.addValidationData(fundNameValidation);
	            worksheet.addValidationData(principalValidation);
	            worksheet.addValidationData(repaidFrequencyValidation);
	            worksheet.addValidationData(loanTermFrequencyValidation);
	            worksheet.addValidationData(noOfRepaymentsValidation);
	            worksheet.addValidationData(loanTermValidation);
	            worksheet.addValidationData(interestValidation);
	            worksheet.addValidationData(interestFrequencyValidation);
	            worksheet.addValidationData(amortizationValidation);
	            worksheet.addValidationData(interestMethodValidation);
	            worksheet.addValidationData(interestCalculationPeriodValidation);
	            worksheet.addValidationData(repaymentStrategyValidation);
	            worksheet.addValidationData(arrearsToleranceValidation);
	            worksheet.addValidationData(graceOnPrincipalPaymentValidation);
	            worksheet.addValidationData(graceOnInterestPaymentValidation);
	            worksheet.addValidationData(graceOnInterestChargedValidation);
	            worksheet.addValidationData(lastRepaymentDateValidation);
	            worksheet.addValidationData(repaymentTypeValidation);
	    	} catch (RuntimeException re) {
	    		result.addError(re.getMessage());
	    	}
	       return result;
	    }
	    
	    private Result setDefaults(Sheet worksheet) {
	    	Result result = new Result();
	    	try {
	    		for(Integer rowNo = 1; rowNo < 1000; rowNo++)
	    		{
	    			Row row = worksheet.createRow(rowNo);
	    			writeFormula(FUND_NAME_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"FUND_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"FUND_\",$D"+ (rowNo + 1) + ")))");
	    			writeFormula(PRINCIPAL_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"PRINCIPAL_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"PRINCIPAL_\",$D"+ (rowNo + 1) + ")))");
	    			writeFormula(REPAID_EVERY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"REPAYMENT_EVERY_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"REPAYMENT_EVERY_\",$D"+ (rowNo + 1) + ")))");
	    			writeFormula(REPAID_EVERY_FREQUENCY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"REPAYMENT_FREQUENCY_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"REPAYMENT_FREQUENCY_\",$D"+ (rowNo + 1) + ")))");
	    			writeFormula(NO_OF_REPAYMENTS_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"NO_REPAYMENT_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"NO_REPAYMENT_\",$D"+ (rowNo + 1) + ")))");
	    			writeFormula(LOAN_TERM_COL, row, "IF(ISERROR($L" + (rowNo + 1) + "*$M" + (rowNo + 1) + "),\"\",$L" + (rowNo + 1) + "*$M" + (rowNo + 1) + ")");
	    			writeFormula(LOAN_TERM_FREQUENCY_COL, row, "$N" + (rowNo + 1));
	    			writeFormula(NOMINAL_INTEREST_RATE_FREQUENCY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"INTEREST_FREQUENCY_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"INTEREST_FREQUENCY_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(NOMINAL_INTEREST_RATE_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"INTEREST_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"INTEREST_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(AMORTIZATION_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"AMORTIZATION_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"AMORTIZATION_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(INTEREST_METHOD_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"INTEREST_TYPE_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"INTEREST_TYPE_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(INTEREST_CALCULATION_PERIOD_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"INTEREST_CALCULATION_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"INTEREST_CALCULATION_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(ARREARS_TOLERANCE_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"ARREARS_TOLERANCE_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"ARREARS_TOLERANCE_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(REPAYMENT_STRATEGY_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"STRATEGY_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"STRATEGY_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(GRACE_ON_PRINCIPAL_PAYMENT_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"GRACE_PRINCIPAL_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"GRACE_PRINCIPAL_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(GRACE_ON_INTEREST_PAYMENT_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"GRACE_INTEREST_PAYMENT_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"GRACE_INTEREST_PAYMENT_\",$D" + (rowNo + 1) + ")))");
	    			writeFormula(GRACE_ON_INTEREST_CHARGED_COL, row, "IF(ISERROR(INDIRECT(CONCATENATE(\"GRACE_INTEREST_CHARGED_\",$D" + (rowNo + 1) + "))),\"\",INDIRECT(CONCATENATE(\"GRACE_INTEREST_CHARGED_\",$D" + (rowNo + 1) + ")))");
	    		}
	    	} catch (RuntimeException re) {
	    		result.addError(re.getMessage());
	    	}
	       return result;
	    }
	    
	    private void setNames(Sheet worksheet) {
	    	Workbook loanWorkbook = worksheet.getWorkbook();
        	ArrayList<String> officeNames = new ArrayList<String>(Arrays.asList(officeSheetPopulator.getOfficeNames()));
        	List<LoanProduct> products = productSheetPopulator.getProducts();
        	
        	//Office Names
        	Name officeGroup = loanWorkbook.createName();
        	officeGroup.setNameName("Office");
        	officeGroup.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));
        	
        	//Client and Loan Officer Names for each office
        	for(Integer i = 0; i < officeNames.size(); i++) {
        		Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator.getOfficeNameToBeginEndIndexesOfClients().get(i);
        		Integer[] officeNameToBeginEndIndexesOfStaff = personnelSheetPopulator.getOfficeNameToBeginEndIndexesOfStaff().get(i);
        		Integer[] officeNameToBeginEndIndexesOfGroups = groupSheetPopulator.getOfficeNameToBeginEndIndexesOfGroups().get(i);
        		Name clientName = loanWorkbook.createName();
        		Name loanOfficerName = loanWorkbook.createName();
        		Name groupName = loanWorkbook.createName();
        	    
        		if(officeNameToBeginEndIndexesOfStaff != null) {
        	       loanOfficerName.setNameName("Staff_" + officeNames.get(i));
        	       loanOfficerName.setRefersToFormula("Staff!$B$" + officeNameToBeginEndIndexesOfStaff[0] + ":$B$" + officeNameToBeginEndIndexesOfStaff[1]);
        		}
        	    if(officeNameToBeginEndIndexesOfClients != null) {
        	    	clientName.setNameName("Client_" + officeNames.get(i));
            	    clientName.setRefersToFormula("Clients!$B$" + officeNameToBeginEndIndexesOfClients[0] + ":$B$" + officeNameToBeginEndIndexesOfClients[1]);
        	    }
        	    if(officeNameToBeginEndIndexesOfGroups != null) {
        	    	groupName.setNameName("Group_" + officeNames.get(i));
            	    groupName.setRefersToFormula("Groups!$B$" + officeNameToBeginEndIndexesOfGroups[0] + ":$B$" + officeNameToBeginEndIndexesOfGroups[1]);
        	    }
        	    
        	}
        	
        	//Product Name
        	Name productGroup = loanWorkbook.createName();
        	productGroup.setNameName("Products");
        	productGroup.setRefersToFormula("Products!$B$2:$B$" + (productSheetPopulator.getProductsSize() + 1));
        	
        	//Fund Name
        	Name fundGroup = loanWorkbook.createName();
        	fundGroup.setNameName("Funds");
        	fundGroup.setRefersToFormula("Extras!$B$2:$B$" + (extrasSheetPopulator.getFundsSize() + 1));
        	
        	//Payment Type Name
        	Name paymentTypeGroup = loanWorkbook.createName();
        	paymentTypeGroup.setNameName("PaymentTypes");
        	paymentTypeGroup.setRefersToFormula("Extras!$D$2:$D$" + (extrasSheetPopulator.getPaymentTypesSize() + 1));
        	
        	//Default Fund, Default Principal, Min Principal, Max Principal, Default No. of Repayments, Min Repayments, Max Repayments, Repayment Every,
        	//Repayment Every Frequency, Interest Rate, Min Interest Rate, Max Interest Rate, Interest Frequency, Amortization, Interest Type,
        	//Interest Calculation Period, Transaction Processing Strategy, Arrears Tolerance, GraceOnPrincipalPayment, GraceOnInterestPayment, 
        	//GraceOnInterestCharged, StartDate Names for each loan product
        	for(Integer i = 0; i < products.size(); i++) {
        		Name fundName = loanWorkbook.createName();
        		Name principalName = loanWorkbook.createName();
        		Name minPrincipalName = loanWorkbook.createName();
        		Name maxPrincipalName = loanWorkbook.createName();
        		Name noOfRepaymentName = loanWorkbook.createName();
        		Name minNoOfRepayment = loanWorkbook.createName();
        		Name maxNoOfRepaymentName = loanWorkbook.createName();
        		Name repaymentEveryName = loanWorkbook.createName();
        		Name repaymentFrequencyName = loanWorkbook.createName();
        		Name interestName = loanWorkbook.createName();
        		Name minInterestName = loanWorkbook.createName();
        		Name maxInterestName = loanWorkbook.createName();
        		Name interestFrequencyName = loanWorkbook.createName();
        		Name amortizationName = loanWorkbook.createName();
        		Name interestTypeName = loanWorkbook.createName();
        		Name interestCalculationPeriodName = loanWorkbook.createName();
        		Name transactionProcessingStrategyName = loanWorkbook.createName();
        		Name arrearsToleranceName = loanWorkbook.createName();
        		Name graceOnPrincipalPaymentName = loanWorkbook.createName();
        		Name graceOnInterestPaymentName = loanWorkbook.createName();
        		Name graceOnInterestChargedName = loanWorkbook.createName();
        		Name startDateName = loanWorkbook.createName();
        		String productName = products.get(i).getName().replaceAll("[ ]", "_");
        	    fundName.setNameName("FUND_" + productName);
        	    principalName.setNameName("PRINCIPAL_" + productName);
        	    minPrincipalName.setNameName("MIN_PRINCIPAL_" + productName);
        	    maxPrincipalName.setNameName("MAX_PRINCIPAL_" + productName);
        	    noOfRepaymentName.setNameName("NO_REPAYMENT_" + productName);
        	    minNoOfRepayment.setNameName("MIN_REPAYMENT_" + productName);
        	    maxNoOfRepaymentName.setNameName("MAX_REPAYMENT_" + productName);
        	    repaymentEveryName.setNameName("REPAYMENT_EVERY_" + productName);
        	    repaymentFrequencyName.setNameName("REPAYMENT_FREQUENCY_" + productName);
        	    interestName.setNameName("INTEREST_" + productName);
        	    minInterestName.setNameName("MIN_INTEREST_" + productName);
        	    maxInterestName.setNameName("MAX_INTEREST_" + productName);
        	    interestFrequencyName .setNameName("INTEREST_FREQUENCY_" + productName);
        	    amortizationName.setNameName("AMORTIZATION_" + productName);
        	    interestTypeName.setNameName("INTEREST_TYPE_" + productName);
        	    interestCalculationPeriodName.setNameName("INTEREST_CALCULATION_" + productName);
        	    transactionProcessingStrategyName.setNameName("STRATEGY_" + productName);
        	    arrearsToleranceName.setNameName("ARREARS_TOLERANCE_" + productName);
        	    graceOnPrincipalPaymentName.setNameName("GRACE_PRINCIPAL_" + productName);
        	    graceOnInterestPaymentName.setNameName("GRACE_INTEREST_PAYMENT_" + productName);
        	    graceOnInterestChargedName.setNameName("GRACE_INTEREST_CHARGED_" + productName);
        	    startDateName.setNameName("START_DATE_" + productName);
        	    if(products.get(i).getFundName() != null)
        	        fundName.setRefersToFormula("Products!$C$" + (i + 2));
        	    principalName.setRefersToFormula("Products!$D$" + (i + 2));
        	    minPrincipalName.setRefersToFormula("Products!$E$" + (i + 2));
        	    maxPrincipalName.setRefersToFormula("Products!$F$" + (i + 2));
        	    noOfRepaymentName.setRefersToFormula("Products!$G$" + (i + 2));
        	    minNoOfRepayment.setRefersToFormula("Products!$H$" + (i + 2));
        	    maxNoOfRepaymentName.setRefersToFormula("Products!$I$" + (i + 2));
        	    repaymentEveryName.setRefersToFormula("Products!$J$" + (i + 2));
        	    repaymentFrequencyName.setRefersToFormula("Products!$K$" + (i + 2));
        	    interestName.setRefersToFormula("Products!$L$" + (i + 2));
        	    minInterestName.setRefersToFormula("Products!$M$" + (i + 2));
        	    maxInterestName.setRefersToFormula("Products!$N$" + (i + 2));
        	    interestFrequencyName .setRefersToFormula("Products!$O$" + (i + 2));
        	    amortizationName.setRefersToFormula("Products!$P$" + (i + 2));
        	    interestTypeName.setRefersToFormula("Products!$Q$" + (i + 2));
        	    interestCalculationPeriodName.setRefersToFormula("Products!$R$" + (i + 2));
        	    transactionProcessingStrategyName.setRefersToFormula("Products!$T$" + (i + 2));
        	    arrearsToleranceName.setRefersToFormula("Products!$S$" + (i + 2));
        	    graceOnPrincipalPaymentName.setRefersToFormula("Products!$U$" + (i + 2));
        	    graceOnInterestPaymentName.setRefersToFormula("Products!$V$" + (i + 2));
        	    graceOnInterestChargedName.setRefersToFormula("Products!$W$" + (i + 2));
        	    startDateName.setRefersToFormula("Products!$X$" + (i + 2));
        	}
	    }
	    
}
