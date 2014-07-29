package org.openmf.mifos.dataimport.populator.accounting;

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
import org.openmf.mifos.dataimport.populator.ExtrasSheetPopulator;
import org.openmf.mifos.dataimport.populator.GlAccountSheetPopulator;
import org.openmf.mifos.dataimport.populator.OfficeSheetPopulator;

public class AddJournalEntriesWorkbookPopulator extends
		AbstractWorkbookPopulator {

	private OfficeSheetPopulator officeSheetPopulator;
	private GlAccountSheetPopulator glAccountSheetPopulator;
	private ExtrasSheetPopulator extrasSheetPopulator;

	private static final int OFFICE_NAME_COL = 0;

	private static final int TRANSACION_ON_DATE_COL = 1;

	private static final int CURRENCY_NAME_COL = 2;

	private static final int PAYMENT_TYPE_ID_COL = 3;

	private static final int TRANSACTION_ID_COL = 4;

	private static final int GL_ACCOUNT_ID_CREDIT_COL = 5;

	private static final int AMOUNT_CREDIT_COL = 6;

	private static final int GL_ACCOUNT_ID_DEBIT_COL = 7;

	private static final int AMOUNT_DEBIT_COL = 8;

	public AddJournalEntriesWorkbookPopulator(
			OfficeSheetPopulator officeSheetPopulator,
			GlAccountSheetPopulator glAccountSheetPopulator,
			ExtrasSheetPopulator extrasSheetPopulator) {
		this.officeSheetPopulator = officeSheetPopulator;
		this.glAccountSheetPopulator = glAccountSheetPopulator;
		this.extrasSheetPopulator = extrasSheetPopulator;
	}

	@Override
	public Result downloadAndParse() {
		Result result = officeSheetPopulator.downloadAndParse();
		if (result.isSuccess())
			result = glAccountSheetPopulator.downloadAndParse();
		if (result.isSuccess())
			result = extrasSheetPopulator.downloadAndParse();
		return result;
	}

	@Override
	public Result populate(Workbook workbook) {
		Sheet addJournalEntriesSheet = workbook
				.createSheet("AddJournalEntries");
		Result result = officeSheetPopulator.populate(workbook);
		if (result.isSuccess())
			result = glAccountSheetPopulator.populate(workbook);
		if (result.isSuccess())
			result = extrasSheetPopulator.populate(workbook);
		if (result.isSuccess())
			result = setRules(addJournalEntriesSheet);
		if (result.isSuccess())
			setLayout(addJournalEntriesSheet);

		return result;
	}

	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
		rowHeader.setHeight((short) 500);
		worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
		worksheet.setColumnWidth(TRANSACION_ON_DATE_COL, 4000);
		worksheet.setColumnWidth(CURRENCY_NAME_COL, 4000);
		worksheet.setColumnWidth(PAYMENT_TYPE_ID_COL, 4000);
		worksheet.setColumnWidth(TRANSACTION_ID_COL, 4000);
		worksheet.setColumnWidth(GL_ACCOUNT_ID_CREDIT_COL, 4000);
		worksheet.setColumnWidth(AMOUNT_CREDIT_COL, 4000);
		worksheet.setColumnWidth(GL_ACCOUNT_ID_DEBIT_COL, 4000);
		worksheet.setColumnWidth(AMOUNT_DEBIT_COL, 4000);

		writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
		writeString(TRANSACION_ON_DATE_COL, rowHeader, "Transaction On *");
		writeString(CURRENCY_NAME_COL, rowHeader, "Currecy Type*");
		writeString(PAYMENT_TYPE_ID_COL, rowHeader, "Payment Type*");
		writeString(TRANSACTION_ID_COL, rowHeader, "Transaction Id*");
		writeString(GL_ACCOUNT_ID_CREDIT_COL, rowHeader, "Credit Account Type*");
		writeString(AMOUNT_CREDIT_COL, rowHeader, "Amount*");
		writeString(GL_ACCOUNT_ID_DEBIT_COL, rowHeader, "Debit Account Type*");
		writeString(AMOUNT_DEBIT_COL, rowHeader, "Amount*");
	}

	private Result setRules(Sheet worksheet) {
		Result result = new Result();
		try {
			CellRangeAddressList officeNameRange = new CellRangeAddressList(1,
					SpreadsheetVersion.EXCEL97.getLastRowIndex(),
					OFFICE_NAME_COL, OFFICE_NAME_COL);

			CellRangeAddressList currencyCodeRange = new CellRangeAddressList(
					1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
					CURRENCY_NAME_COL, CURRENCY_NAME_COL);
			CellRangeAddressList paymenttypeRange = new CellRangeAddressList(1,
					SpreadsheetVersion.EXCEL97.getLastRowIndex(),
					PAYMENT_TYPE_ID_COL, PAYMENT_TYPE_ID_COL);

			CellRangeAddressList glaccountCreditRange = new CellRangeAddressList(
					1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
					GL_ACCOUNT_ID_CREDIT_COL, GL_ACCOUNT_ID_CREDIT_COL);
			CellRangeAddressList glaccountDebitRange = new CellRangeAddressList(
					1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
					GL_ACCOUNT_ID_DEBIT_COL, GL_ACCOUNT_ID_DEBIT_COL);

			DataValidationHelper validationHelper = new HSSFDataValidationHelper(
					(HSSFSheet) worksheet);

			setNames(worksheet);

			DataValidationConstraint officeNameConstraint = validationHelper
					.createFormulaListConstraint("Office");
			DataValidationConstraint currencyCodeConstraint = validationHelper
					.createFormulaListConstraint("Currency");
			DataValidationConstraint paymentTypeConstraint = validationHelper
					.createFormulaListConstraint("PaymentType");

			DataValidationConstraint glaccountConstraint = validationHelper
					.createFormulaListConstraint("GlAccounts");

			DataValidation officeValidation = validationHelper
					.createValidation(officeNameConstraint, officeNameRange);
			DataValidation currencyCodeValidation = validationHelper
					.createValidation(currencyCodeConstraint, currencyCodeRange);
			DataValidation paymentTypeValidation = validationHelper
					.createValidation(paymentTypeConstraint, paymenttypeRange);

			DataValidation glaccountCreditValidation = validationHelper
					.createValidation(glaccountConstraint, glaccountCreditRange);
			DataValidation glaccountDebitValidation = validationHelper
					.createValidation(glaccountConstraint, glaccountDebitRange);

			worksheet.addValidationData(officeValidation);
			worksheet.addValidationData(currencyCodeValidation);
			worksheet.addValidationData(paymentTypeValidation);

			worksheet.addValidationData(glaccountCreditValidation);
			worksheet.addValidationData(glaccountDebitValidation);

		} catch (RuntimeException re) {
			result.addError(re.getMessage());
			re.printStackTrace();
		}
		return result;
	}

	private void setNames(Sheet worksheet) {
		Workbook addJournalEntriesWorkbook = worksheet.getWorkbook();
		ArrayList<String> officeNames = new ArrayList<String>(
				Arrays.asList(officeSheetPopulator.getOfficeNames()));
		// Office Names
		Name officeGroup = addJournalEntriesWorkbook.createName();
		officeGroup.setNameName("Office");
		officeGroup.setRefersToFormula("Offices!$B$2:$B$"
				+ (officeNames.size() + 1));
		// Payment Type Name
		Name paymentTypeGroup = addJournalEntriesWorkbook.createName();
		paymentTypeGroup.setNameName("PaymentType");
		paymentTypeGroup.setRefersToFormula("Extras!$D$2:$D$"
				+ (extrasSheetPopulator.getPaymentTypesSize() + 1));
		// Currency Type Name
		Name currencyGroup = addJournalEntriesWorkbook.createName();
		currencyGroup.setNameName("Currency");
		currencyGroup.setRefersToFormula("Extras!$F$2:$F$"
				+ (extrasSheetPopulator.getCurrencysSize() + 1));

		// Account Name
		Name glaccountGroup = addJournalEntriesWorkbook.createName();
		glaccountGroup.setNameName("GlAccounts");
		glaccountGroup.setRefersToFormula("GlAccounts!$B$2:$B$"
				+ (glAccountSheetPopulator.getGlAccountNamesSize() + 1));

	}
}