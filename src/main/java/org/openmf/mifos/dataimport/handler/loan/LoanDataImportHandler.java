package org.openmf.mifos.dataimport.handler.loan;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Approval;
import org.openmf.mifos.dataimport.dto.Charge;
import org.openmf.mifos.dataimport.dto.Transaction;
import org.openmf.mifos.dataimport.dto.loan.DisbursalData;
import org.openmf.mifos.dataimport.dto.loan.GroupLoan;
import org.openmf.mifos.dataimport.dto.loan.Loan;
import org.openmf.mifos.dataimport.dto.loan.LoanDisbursal;
import org.openmf.mifos.dataimport.handler.AbstractDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class LoanDataImportHandler extends AbstractDataImportHandler {

    private static final Logger logger = LoggerFactory.getLogger(LoanDataImportHandler.class);

    @SuppressWarnings("CPD-START")
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
    private static final int STATUS_COL = 31;
    private static final int LOAN_ID_COL = 32;
    private static final int FAILURE_REPORT_COL = 33;
    private static final int EXTERNAL_ID_COL = 34;
    private static final int CHARGE_ID_1 = 35;
    private static final int CHARGE_AMOUNT_1 = 36;
    private static final int CHARGE_DUE_DATE_1 = 37;
    private static final int CHARGE_ID_2 = 38;
    private static final int CHARGE_AMOUNT_2 = 39;
    private static final int CHARGE_DUE_DATE_2 = 40;
    private static final int GROUP_ID = 41;
    private static final int LINK_ACCOUNT_ID = 44;

    @SuppressWarnings("CPD-END")
    private List<Loan> loans = new ArrayList<Loan>();
    private List<Approval> approvalDates = new ArrayList<Approval>();
    // private List<LoanDisbursal> disbursalDates = new
    // ArrayList<LoanDisbursal>();
    private List<Transaction> loanRepayments = new ArrayList<Transaction>();
    private List<DisbursalData> disbursalDates = new ArrayList<DisbursalData>();
    private final RestClient restClient;
    private final Workbook workbook;

    public LoanDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
    }

    @Override
    public Result parse() {
        Result result = new Result();
        Sheet loanSheet = workbook.getSheet("Loans");
        Integer noOfEntries = getNumberOfRows(loanSheet, 0);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = loanSheet.getRow(rowIndex);
                if (isNotImported(row, STATUS_COL)) {
                    loans.add(parseAsLoan(row));
                    approvalDates.add(parseAsLoanApproval(row));
                    disbursalDates.add(parseAsDisbursalData(row));
                    loanRepayments.add(parseAsLoanRepayment(row));
                }
            } catch (RuntimeException re) {
                logger.error("row = " + rowIndex, re);
                result.addError("Row = " + rowIndex + " , " + re.getMessage());
            }
        }

        return result;
    }

    private Loan parseAsLoan(Row row) {
        String externalId = readAsString(EXTERNAL_ID_COL, row);
        String status = readAsString(STATUS_COL, row);
        String productName = readAsString(PRODUCT_COL, row);
        String productId = getIdByName(workbook.getSheet("Products"), productName).toString();
        String loanOfficerName = readAsString(LOAN_OFFICER_NAME_COL, row);
        String loanOfficerId = getIdByName(workbook.getSheet("Staff"), loanOfficerName).toString();
        String submittedOnDate = readAsDate(SUBMITTED_ON_DATE_COL, row);
        String fundName = readAsString(FUND_NAME_COL, row);
        String fundId;
        if (fundName.equals(""))
            fundId = "";
        else
            fundId = getIdByName(workbook.getSheet("Extras"), fundName).toString();
        String principal = readAsDouble(PRINCIPAL_COL, row).toString();
        String numberOfRepayments = readAsString(NO_OF_REPAYMENTS_COL, row);
        String repaidEvery = readAsString(REPAID_EVERY_COL, row);
        String repaidEveryFrequency = readAsString(REPAID_EVERY_FREQUENCY_COL, row);
        String repaidEveryFrequencyId = "";
        if (repaidEveryFrequency.equalsIgnoreCase("Days"))
            repaidEveryFrequencyId = "0";
        else if (repaidEveryFrequency.equalsIgnoreCase("Weeks"))
            repaidEveryFrequencyId = "1";
        else if (repaidEveryFrequency.equalsIgnoreCase("Months")) repaidEveryFrequencyId = "2";
        String loanTerm = readAsString(LOAN_TERM_COL, row);
        String loanTermFrequency = readAsString(LOAN_TERM_FREQUENCY_COL, row);
        String loanTermFrequencyId = "";
        if (loanTermFrequency.equalsIgnoreCase("Days"))
            loanTermFrequencyId = "0";
        else if (loanTermFrequency.equalsIgnoreCase("Weeks"))
            loanTermFrequencyId = "1";
        else if (loanTermFrequency.equalsIgnoreCase("Months")) loanTermFrequencyId = "2";
        Double nominalInterestRate = readAsDouble(NOMINAL_INTEREST_RATE_COL, row);
        String amortization = readAsString(AMORTIZATION_COL, row);
        String amortizationId = "";
        if (amortization.equalsIgnoreCase("Equal principal payments"))
            amortizationId = "0";
        else if (amortization.equalsIgnoreCase("Equal installments")) amortizationId = "1";
        String interestMethod = readAsString(INTEREST_METHOD_COL, row);
        String interestMethodId = "";
        if (interestMethod.equalsIgnoreCase("Flat"))
            interestMethodId = "1";
        else if (interestMethod.equalsIgnoreCase("Declining Balance")) interestMethodId = "0";
        String interestCalculationPeriod = readAsString(INTEREST_CALCULATION_PERIOD_COL, row);
        String interestCalculationPeriodId = "";
        if (interestCalculationPeriod.equalsIgnoreCase("Daily"))
            interestCalculationPeriodId = "0";
        else if (interestCalculationPeriod.equalsIgnoreCase("Same as repayment period")) interestCalculationPeriodId = "1";
        String arrearsTolerance = readAsString(ARREARS_TOLERANCE_COL, row);
        String repaymentStrategy = readAsString(REPAYMENT_STRATEGY_COL, row);
        String repaymentStrategyId = "";
        if (repaymentStrategy.equalsIgnoreCase("Mifos style"))
            repaymentStrategyId = "1";
        else if (repaymentStrategy.equalsIgnoreCase("Heavensfamily"))
            repaymentStrategyId = "2";
        else if (repaymentStrategy.equalsIgnoreCase("Creocore"))
            repaymentStrategyId = "3";
        else if (repaymentStrategy.equalsIgnoreCase("RBI (India)"))
            repaymentStrategyId = "4";
        else if (repaymentStrategy.equalsIgnoreCase("Principal Interest Penalties Fees Order"))
            repaymentStrategyId = "5";
        else if (repaymentStrategy.equalsIgnoreCase("Interest Principal Penalties Fees Order")) repaymentStrategyId = "6";
        String graceOnPrincipalPayment = readAsString(GRACE_ON_PRINCIPAL_PAYMENT_COL, row);
        String graceOnInterestPayment = readAsString(GRACE_ON_INTEREST_PAYMENT_COL, row);
        String graceOnInterestCharged = readAsString(GRACE_ON_INTEREST_CHARGED_COL, row);
        String interestChargedFromDate = readAsDate(INTEREST_CHARGED_FROM_COL, row);
        String firstRepaymentOnDate = readAsDate(FIRST_REPAYMENT_COL, row);
        String loanType = readAsString(LOAN_TYPE_COL, row).toLowerCase(Locale.ENGLISH);
        String clientOrGroupName = readAsString(CLIENT_NAME_COL, row);

        List<Charge> charges = new ArrayList<Charge>();

        String charge1 = readAsString(CHARGE_ID_1, row);
        String charge2 = readAsString(CHARGE_ID_2, row);
        String groupId = readAsString(GROUP_ID, row);
        String linkAccountId = readAsString(LINK_ACCOUNT_ID, row);

        if (!charge1.equalsIgnoreCase("")) {
            charges.add(new Charge(readAsString(CHARGE_ID_1, row), readAsDouble(CHARGE_AMOUNT_1, row), readAsDate(CHARGE_DUE_DATE_1, row)));
        }

        if (!charge2.equalsIgnoreCase("")) {

            charges.add(new Charge(readAsString(CHARGE_ID_2, row), readAsDouble(CHARGE_AMOUNT_2, row), readAsDate(CHARGE_DUE_DATE_2, row)));
        }
        if (loanType.equals("individual")) {
            String clientId = getIdByName(workbook.getSheet("Clients"), clientOrGroupName).toString();
            return new Loan(loanType, clientId, productId, loanOfficerId, submittedOnDate, fundId, principal, numberOfRepayments,
                    repaidEvery, repaidEveryFrequencyId, loanTerm, loanTermFrequencyId, nominalInterestRate, submittedOnDate,
                    amortizationId, interestMethodId, interestCalculationPeriodId, arrearsTolerance, repaymentStrategyId,
                    graceOnPrincipalPayment, graceOnInterestPayment, graceOnInterestCharged, interestChargedFromDate, firstRepaymentOnDate,
                    row.getRowNum(), status, externalId, null, charges, linkAccountId);
        } else if (loanType.equals("jlg")) {
            String clientId = getIdByName(workbook.getSheet("Clients"), clientOrGroupName).toString();
            return new Loan(loanType, clientId, productId, loanOfficerId, submittedOnDate, fundId, principal, numberOfRepayments,
                    repaidEvery, repaidEveryFrequencyId, loanTerm, loanTermFrequencyId, nominalInterestRate, submittedOnDate,
                    amortizationId, interestMethodId, interestCalculationPeriodId, arrearsTolerance, repaymentStrategyId,
                    graceOnPrincipalPayment, graceOnInterestPayment, graceOnInterestCharged, interestChargedFromDate, firstRepaymentOnDate,
                    row.getRowNum(), status, externalId, groupId, charges, linkAccountId);
        } else {
            String groupIdforGroupLoan = getIdByName(workbook.getSheet("Groups"), clientOrGroupName).toString();
            return new GroupLoan(loanType, groupIdforGroupLoan, productId, loanOfficerId, submittedOnDate, fundId, principal,
                    numberOfRepayments, repaidEvery, repaidEveryFrequencyId, loanTerm, loanTermFrequencyId, nominalInterestRate,
                    submittedOnDate, amortizationId, interestMethodId, interestCalculationPeriodId, arrearsTolerance, repaymentStrategyId,
                    graceOnPrincipalPayment, graceOnInterestPayment, graceOnInterestCharged, interestChargedFromDate, firstRepaymentOnDate,
                    row.getRowNum(), status, externalId, linkAccountId);
        }
    }

    private Approval parseAsLoanApproval(Row row) {
        String approvedDate = readAsDate(APPROVED_DATE_COL, row);
        if (!approvedDate.equals(""))
            return new Approval(approvedDate, row.getRowNum());
        else
            return null;
    }

    private DisbursalData parseAsDisbursalData(Row row) {
        String disbursedDate = readAsDate(DISBURSED_DATE_COL, row);
        String paymentType = readAsString(DISBURSED_PAYMENT_TYPE_COL, row);
        String paymentTypeId = null;
        String linkAccountId = readAsString(LINK_ACCOUNT_ID, row);

        if (!paymentType.equals("")) {
            paymentTypeId = getIdByName(workbook.getSheet("Extras"), paymentType).toString();
        }
        if (!disbursedDate.equals("")) {
            LoanDisbursal loanDisbursal = new LoanDisbursal(disbursedDate, paymentTypeId, row.getRowNum());
            return new DisbursalData(linkAccountId, loanDisbursal);
        }
        return null;
    }

    private Transaction parseAsLoanRepayment(Row row) {
        String repaymentAmount = readAsDouble(TOTAL_AMOUNT_REPAID_COL, row).toString();
        String lastRepaymentDate = readAsDate(LAST_REPAYMENT_DATE_COL, row);
        String repaymentType = readAsString(REPAYMENT_TYPE_COL, row);
        String repaymentTypeId = getIdByName(workbook.getSheet("Extras"), repaymentType).toString();
        if (!repaymentAmount.equals("0.0"))
            return new Transaction(repaymentAmount, lastRepaymentDate, repaymentTypeId, row.getRowNum());
        else
            return null;
    }

    @Override
    public Result upload() {
        Result result = new Result();
        Sheet loanSheet = workbook.getSheet("Loans");
        restClient.createAuthToken();
        int progressLevel = 0;
        String loanId;
        for (int i = 0; i < loans.size(); i++) {
            Row row = loanSheet.getRow(loans.get(i).getRowIndex());
            Cell errorReportCell = row.createCell(FAILURE_REPORT_COL);
            Cell statusCell = row.createCell(STATUS_COL);
            loanId = "";
            try {
                String response = "";
                String status = loans.get(i).getStatus();
                progressLevel = getProgressLevel(status);

                if (progressLevel == 0) {
                    response = uploadLoan(i);
                    loanId = getLoanId(response);
                    progressLevel = 1;
                } else
                    loanId = readAsInt(LOAN_ID_COL, loanSheet.getRow(loans.get(i).getRowIndex()));

                if (progressLevel <= 1) progressLevel = uploadLoanApproval(loanId, i);

                if (progressLevel <= 2) progressLevel = uploadDisbursalData(loanId, i);

                if (loanRepayments.get(i) != null) progressLevel = uploadLoanRepayment(loanId, i);

                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (RuntimeException re) {
                String message = parseStatus(re.getMessage());
                String status = "";

                if (progressLevel == 0)
                    status = "Creation";
                else if (progressLevel == 1)
                    status = "Approval";
                else if (progressLevel == 2)
                    status = "Disbursal";
                else if (progressLevel == 3) status = "Repayment";
                statusCell.setCellValue(status + " failed.");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));

                if (progressLevel > 0) row.createCell(LOAN_ID_COL).setCellValue(Integer.parseInt(loanId));

                errorReportCell.setCellValue(message);
                result.addError("Row = " + loans.get(i).getRowIndex() + " ," + message);
            }
        }

        setReportHeaders(loanSheet);
        return result;
    }

    private int getProgressLevel(String status) {
        if (status.equals("") || status.equals("Creation failed."))
            return 0;
        else if (status.equals("Approval failed."))
            return 1;
        else if (status.equals("Disbursal failed."))
            return 2;
        else if (status.equals("Repayment failed.")) return 3;
        return 0;
    }

    private String uploadLoan(int rowIndex) {
        Gson gson = new Gson();
        String payload = gson.toJson(loans.get(rowIndex));
        String response = restClient.post("loans", payload);

        return response;
    }

    private String getLoanId(String response) {
        JsonParser parser = new JsonParser();
        JsonObject obj = parser.parse(response).getAsJsonObject();
        return obj.get("loanId").getAsString();
    }

    private Integer uploadLoanApproval(String loanId, int rowIndex) {
        if (approvalDates.get(rowIndex) != null) {
            Gson gson = new Gson();
            String payload = gson.toJson(approvalDates.get(rowIndex));
            restClient.post("loans/" + loanId + "?command=approve", payload);
        }
        return 2;
    }

    private Integer uploadDisbursalData(String loanId, int rowIndex) {
        if (approvalDates.get(rowIndex) != null && disbursalDates.get(rowIndex) != null) {

            DisbursalData disbusalData = disbursalDates.get(rowIndex);
            String linkAccountId = disbusalData.getLinkAccountId();
            Gson gson = new Gson();
            if (linkAccountId != null && linkAccountId != "") {
                String payload = gson.toJson(disbusalData.getLoanDisbursal());
                restClient.post("loans/" + loanId + "?command=disbursetosavings", payload);
            } else {
                String payload = gson.toJson(disbusalData.getLoanDisbursal());
                restClient.post("loans/" + loanId + "?command=disburse", payload);
            }
        }
        return 3;
    }

    private Integer uploadLoanRepayment(String loanId, int rowIndex) {
        Gson gson = new Gson();
        String payload = gson.toJson(loanRepayments.get(rowIndex));
        restClient.post("loans/" + loanId + "/transactions?command=repayment", payload);
        return 4;
    }

    private void setReportHeaders(Sheet sheet) {
        sheet.setColumnWidth(STATUS_COL, 4000);
        Row rowHeader = sheet.getRow(0);
        writeString(STATUS_COL, rowHeader, "Status");
        writeString(LOAN_ID_COL, rowHeader, "Loan ID");
        writeString(FAILURE_REPORT_COL, rowHeader, "Report");
    }

    public List<Loan> getLoans() {
        return loans;
    }

    public List<Approval> getApprovalDates() {
        return approvalDates;
    }

    public List<DisbursalData> getDisbursalDates() {
        return disbursalDates;
    }

    public List<Transaction> getLoanRepayments() {
        return loanRepayments;
    }
}
