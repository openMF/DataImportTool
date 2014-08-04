package org.openmf.mifos.dataimport.handler.accounting;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.CreditDebit;
import org.openmf.mifos.dataimport.dto.accounting.JournalEntries;
import org.openmf.mifos.dataimport.handler.AbstractDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.handler.savings.SavingsTransactionDataImportHandler;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;

public class JournalEntriesImportHandler extends AbstractDataImportHandler {
	private static final Logger logger = LoggerFactory
			.getLogger(SavingsTransactionDataImportHandler.class);

	private final RestClient restClient;
	private final Workbook workbook;

	private List<JournalEntries> gltransaction;

	List<CreditDebit> credits = new ArrayList<CreditDebit>();
	List<CreditDebit> debits = new ArrayList<CreditDebit>();

	private String transactionDate = "";

	private static final int OFFICE_NAME_COL = 0;

	private static final int TRANSACION_ON_DATE_COL = 1;

	private static final int CURRENCY_NAME_COL = 2;

	private static final int PAYMENT_TYPE_ID_COL = 3;

	private static final int TRANSACTION_ID_COL = 4;

	private static final int GL_ACCOUNT_ID_CREDIT_COL = 5;

	private static final int AMOUNT_CREDIT_COL = 6;

	private static final int GL_ACCOUNT_ID_DEBIT_COL = 7;

	private static final int AMOUNT_DEBIT_COL = 8;

	private static final int STATUS_COL = 9;

	public JournalEntriesImportHandler(Workbook workbook, RestClient client) {
		this.workbook = workbook;
		this.restClient = client;
		gltransaction = new ArrayList<JournalEntries>();
	}

	@Override
	public Result parse() {
		Result result = new Result();
		// boolean isNewTransaction = false;
		String currentTransactionId = "";
		String prevTransactionId = "";
		JournalEntries journalEntry = null;

		Sheet addJournalEntriesSheet = workbook.getSheet("AddJournalEntries");
		Integer noOfEntries = getNumberOfRows(addJournalEntriesSheet, 4);
		for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
			Row row;
			try {
				row = addJournalEntriesSheet.getRow(rowIndex);

				currentTransactionId = readAsString(TRANSACTION_ID_COL, row);

				if (currentTransactionId.equals(prevTransactionId)) {
					if (journalEntry != null) {

						String creditGLAcct = readAsString(
								GL_ACCOUNT_ID_CREDIT_COL, row);
						String glAccountIdCredit = getIdByName(
								workbook.getSheet("GlAccounts"), creditGLAcct)
								.toString();

						String debitGLAcct = readAsString(
								GL_ACCOUNT_ID_DEBIT_COL, row);

						String glAccountIdDebit = getIdByName(
								workbook.getSheet("GlAccounts"), debitGLAcct)
								.toString();

						String creditAmt = readAsString(AMOUNT_CREDIT_COL, row);
						String debitAmount = readAsString(AMOUNT_DEBIT_COL, row);

						if (!creditGLAcct.equalsIgnoreCase("")) {

							CreditDebit credit = new CreditDebit(
									glAccountIdCredit, creditAmt);
							journalEntry.addCredits(credit);
						}
						if (!debitGLAcct.equalsIgnoreCase("")) {
							CreditDebit debit = new CreditDebit(
									glAccountIdDebit, debitAmount);

							journalEntry.addDebits(debit);
						}
					}
				} else {

					if (journalEntry != null) {
						gltransaction.add(journalEntry);
						journalEntry = null;
					}

					journalEntry = parseAsaddJournalEntries(row);

					// if (isNotImported(row, STATUS_COL)) {
					//
					// }
				}
			} catch (RuntimeException re) {
				logger.error("row = " + rowIndex, re);
				result.addError("Row = " + rowIndex + " , " + re.getMessage());
			}

			prevTransactionId = currentTransactionId;
		}

		// Adding last JE
		gltransaction.add(journalEntry);

		return result;
	}

	private JournalEntries parseAsaddJournalEntries(Row row) {
		String transactionDateCheck = readAsDate(TRANSACION_ON_DATE_COL, row);
		if (!transactionDateCheck.equals(""))
			transactionDate = transactionDateCheck;

		String officeName = readAsString(OFFICE_NAME_COL, row);
		String officeId = getIdByName(workbook.getSheet("Offices"), officeName)
				.toString();
		// String transactionDate = readAsDate(TRANSACION_ON_DATE_COL, row);
		String paymentType = readAsString(PAYMENT_TYPE_ID_COL, row);
		String paymentTypeId = getIdByName(workbook.getSheet("Extras"),
				paymentType).toString();
		String currencyName = readAsString(CURRENCY_NAME_COL, row);
		String currencyCode = getCodeByName(workbook.getSheet("Extras"),
				currencyName).toString();
		// String transactionType = readAsString(TRANSACTION_ID_COL, row);
		String glAccountNameCredit = readAsString(GL_ACCOUNT_ID_CREDIT_COL, row);
		String glAccountIdCredit = getIdByName(workbook.getSheet("GlAccounts"),
				glAccountNameCredit).toString();
		String glAccountNameDebit = readAsString(GL_ACCOUNT_ID_DEBIT_COL, row);
		String glAccountIdDebit = getIdByName(workbook.getSheet("GlAccounts"),
				glAccountNameDebit).toString();

		credits = new ArrayList<CreditDebit>();
		debits = new ArrayList<CreditDebit>();

		String credit = readAsString(GL_ACCOUNT_ID_CREDIT_COL, row);
		String debit = readAsString(GL_ACCOUNT_ID_DEBIT_COL, row);

		if (!credit.equalsIgnoreCase("")) {
			credits.add(new CreditDebit(glAccountIdCredit, readAsString(
					AMOUNT_CREDIT_COL, row)));
		}

		if (!debit.equalsIgnoreCase("")) {
			debits.add(new CreditDebit(glAccountIdDebit, readAsString(
					AMOUNT_DEBIT_COL, row)));
		}

		return new JournalEntries(officeId, transactionDate, currencyCode,
				paymentTypeId, row.getRowNum(), credits, debits);

	}

	@Override
	public Result upload() {
		Result result = new Result();
		Sheet addJournalEntriesSheet = workbook.getSheet("AddJournalEntries");
		restClient.createAuthToken();
		for (JournalEntries transaction : gltransaction) {
			try {
				Gson gson = new Gson();
				String payload = gson.toJson(transaction);
				restClient.post("journalentries", payload);
				Cell statusCell = addJournalEntriesSheet.getRow(
						transaction.getRowIndex()).createCell(STATUS_COL);
				statusCell.setCellValue("Imported");
				statusCell.setCellStyle(getCellStyle(workbook,
						IndexedColors.LIGHT_GREEN));
			} catch (Exception e) {
				Cell transactionDateCell = addJournalEntriesSheet.getRow(
						transaction.getRowIndex()).createCell(
						TRANSACION_ON_DATE_COL);
				transactionDateCell.setCellValue(transaction
						.getTransactionDate());
				String message = parseStatus(e.getMessage());

				Cell statusCell = addJournalEntriesSheet.getRow(
						transaction.getRowIndex()).createCell(STATUS_COL);
				statusCell.setCellValue(message);
				statusCell.setCellStyle(getCellStyle(workbook,
						IndexedColors.RED));
				result.addError("Row = " + transaction.getRowIndex() + " ,"
						+ message);
			}
		}
		addJournalEntriesSheet.setColumnWidth(STATUS_COL, 15000);
		writeString(STATUS_COL, addJournalEntriesSheet.getRow(0), "Status");
		return result;

	}

}