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
import org.openmf.mifos.dataimport.dto.Transaction;
import org.openmf.mifos.dataimport.dto.loan.EmiDetail;
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
    private static final int FIXED_EMI_AMOUNT_COL = 35;
    private static final int MAX_OUTSTANDING_LOAN_BALANCE_COL = 36;
    private static final int EMI_EXPECTED_DISBURSAL_DATE_COL = 37;
    private static final int EMI_PRINCIPAL_COL = 38;

    private final List<Loan> loans = new ArrayList<Loan>();
    private final List<Approval> approvalDates = new ArrayList<Approval>();
    private final List<LoanDisbursal> disbursalDates = new ArrayList<LoanDisbursal>();
    private final List<Transaction> loanRepayments = new ArrayList<Transaction>();

    private final RestClient restClient;

    private final Workbook workbook;

    public LoanDataImportHandler(final Workbook workbook, final RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
    }

    @Override
    public Result parse() {
        final Result result = new Result();
        final Sheet loanSheet = this.workbook.getSheet("Loans");
        final Integer noOfEntries = getNumberOfRows(loanSheet, 0);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = loanSheet.getRow(rowIndex);
                if (isNotImported(row, STATUS_COL)) {
                    this.loans.add(parseAsLoan(row));
                    this.approvalDates.add(parseAsLoanApproval(row));
                    this.disbursalDates.add(parseAsLoanDisbursal(row));
                    this.loanRepayments.add(parseAsLoanRepayment(row));
                }
            } catch (final RuntimeException re) {
                logger.error("row = " + rowIndex, re);
                result.addError("Row = " + rowIndex + " , " + re.getMessage());
            }
        }

        return result;
    }

    private Loan parseAsLoan(final Row row) {
        final String externalId = readAsString(EXTERNAL_ID_COL, row);
        final String status = readAsString(STATUS_COL, row);
        final String productName = readAsString(PRODUCT_COL, row);
        final String productId = getIdByName(this.workbook.getSheet("Products"), productName).toString();
        final String loanOfficerName = readAsString(LOAN_OFFICER_NAME_COL, row);
        final String loanOfficerId = getIdByName(this.workbook.getSheet("Staff"), loanOfficerName).toString();
        final String submittedOnDate = readAsDate(SUBMITTED_ON_DATE_COL, row);
        final String fundName = readAsString(FUND_NAME_COL, row);
        String fundId;
        if (fundName.equals("")) {
            fundId = "";
        } else {
            fundId = getIdByName(this.workbook.getSheet("Extras"), fundName).toString();
        }
        final String principal = readAsDouble(PRINCIPAL_COL, row).toString();
        final String numberOfRepayments = readAsString(NO_OF_REPAYMENTS_COL, row);
        final String repaidEvery = readAsString(REPAID_EVERY_COL, row);
        final String repaidEveryFrequency = readAsString(REPAID_EVERY_FREQUENCY_COL, row);
        String repaidEveryFrequencyId = "";
        if (repaidEveryFrequency.equalsIgnoreCase("Days")) {
            repaidEveryFrequencyId = "0";
        } else if (repaidEveryFrequency.equalsIgnoreCase("Weeks")) {
            repaidEveryFrequencyId = "1";
        } else if (repaidEveryFrequency.equalsIgnoreCase("Months")) {
            repaidEveryFrequencyId = "2";
        }
        final String loanTerm = readAsString(LOAN_TERM_COL, row);
        final String loanTermFrequency = readAsString(LOAN_TERM_FREQUENCY_COL, row);
        String loanTermFrequencyId = "";
        if (loanTermFrequency.equalsIgnoreCase("Days")) {
            loanTermFrequencyId = "0";
        } else if (loanTermFrequency.equalsIgnoreCase("Weeks")) {
            loanTermFrequencyId = "1";
        } else if (loanTermFrequency.equalsIgnoreCase("Months")) {
            loanTermFrequencyId = "2";
        }
        final String nominalInterestRate = readAsString(NOMINAL_INTEREST_RATE_COL, row);
        final String amortization = readAsString(AMORTIZATION_COL, row);
        String amortizationId = "";
        if (amortization.equalsIgnoreCase("Equal principal payments")) {
            amortizationId = "0";
        } else if (amortization.equalsIgnoreCase("Equal installments")) {
            amortizationId = "1";
        }
        final String interestMethod = readAsString(INTEREST_METHOD_COL, row);
        String interestMethodId = "";
        if (interestMethod.equalsIgnoreCase("Flat")) {
            interestMethodId = "1";
        } else if (interestMethod.equalsIgnoreCase("Declining Balance")) {
            interestMethodId = "0";
        }
        final String interestCalculationPeriod = readAsString(INTEREST_CALCULATION_PERIOD_COL, row);
        String interestCalculationPeriodId = "";
        if (interestCalculationPeriod.equalsIgnoreCase("Daily")) {
            interestCalculationPeriodId = "0";
        } else if (interestCalculationPeriod.equalsIgnoreCase("Same as repayment period")) {
            interestCalculationPeriodId = "1";
        }
        final String arrearsTolerance = readAsString(ARREARS_TOLERANCE_COL, row);
        final String repaymentStrategy = readAsString(REPAYMENT_STRATEGY_COL, row);
        String repaymentStrategyId = "";
        if (repaymentStrategy.equalsIgnoreCase("Mifos style")) {
            repaymentStrategyId = "1";
        } else if (repaymentStrategy.equalsIgnoreCase("Heavensfamily")) {
            repaymentStrategyId = "2";
        } else if (repaymentStrategy.equalsIgnoreCase("Creocore")) {
            repaymentStrategyId = "3";
        } else if (repaymentStrategy.equalsIgnoreCase("RBI (India)")) {
            repaymentStrategyId = "4";
        } else if (repaymentStrategy.equalsIgnoreCase("Principal Interest Penalties Fees Order")) {
            repaymentStrategyId = "5";
        } else if (repaymentStrategy.equalsIgnoreCase("Interest Principal Penalties Fees Order")) {
            repaymentStrategyId = "6";
        }
        final String graceOnPrincipalPayment = readAsString(GRACE_ON_PRINCIPAL_PAYMENT_COL, row);
        final String graceOnInterestPayment = readAsString(GRACE_ON_INTEREST_PAYMENT_COL, row);
        final String graceOnInterestCharged = readAsString(GRACE_ON_INTEREST_CHARGED_COL, row);
        final String interestChargedFromDate = readAsDate(INTEREST_CHARGED_FROM_COL, row);
        final String firstRepaymentOnDate = readAsDate(FIRST_REPAYMENT_COL, row);
        final String loanType = readAsString(LOAN_TYPE_COL, row).toLowerCase(Locale.ENGLISH);
        final String clientOrGroupName = readAsString(CLIENT_NAME_COL, row);

        // Adding support for tranches
        final String fixedEmiAmount = readAsDouble(FIXED_EMI_AMOUNT_COL, row).toString();
        final String maxOutstandingLoanBalance = readAsDouble(MAX_OUTSTANDING_LOAN_BALANCE_COL, row).toString();
        final List<EmiDetail> disbursementData = new ArrayList<EmiDetail>(1);

        final String emiExpectedDisbursementDate = readAsDate(EMI_EXPECTED_DISBURSAL_DATE_COL, row);
        final String emiPrincipal = readAsDouble(EMI_PRINCIPAL_COL, row).toString();
        final EmiDetail emiDetail = new EmiDetail(emiExpectedDisbursementDate, emiPrincipal);
        disbursementData.add(emiDetail);

        if (loanType.equals("individual")) {
            final String clientId = getIdByName(this.workbook.getSheet("Clients"), clientOrGroupName).toString();
            return new Loan(loanType, clientId, productId, loanOfficerId, submittedOnDate, fundId, principal, numberOfRepayments,
                    repaidEvery, repaidEveryFrequencyId, loanTerm, loanTermFrequencyId, nominalInterestRate, submittedOnDate,
                    amortizationId, interestMethodId, interestCalculationPeriodId, arrearsTolerance, repaymentStrategyId,
                    graceOnPrincipalPayment, graceOnInterestPayment, graceOnInterestCharged, interestChargedFromDate, firstRepaymentOnDate,
                    row.getRowNum(), status, externalId, fixedEmiAmount, maxOutstandingLoanBalance, disbursementData);
        }
        final String groupId = getIdByName(this.workbook.getSheet("Groups"), clientOrGroupName).toString();
        return new GroupLoan(loanType, groupId, productId, loanOfficerId, submittedOnDate, fundId, principal, numberOfRepayments,
                repaidEvery, repaidEveryFrequencyId, loanTerm, loanTermFrequencyId, nominalInterestRate, submittedOnDate, amortizationId,
                interestMethodId, interestCalculationPeriodId, arrearsTolerance, repaymentStrategyId, graceOnPrincipalPayment,
                graceOnInterestPayment, graceOnInterestCharged, interestChargedFromDate, firstRepaymentOnDate, row.getRowNum(), status,
                externalId, fixedEmiAmount, maxOutstandingLoanBalance, disbursementData);
    }

    private Approval parseAsLoanApproval(final Row row) {
        final String approvedDate = readAsDate(APPROVED_DATE_COL, row);
        if (!approvedDate.equals("")) { return new Approval(approvedDate, row.getRowNum()); }
        return null;
    }

    private LoanDisbursal parseAsLoanDisbursal(final Row row) {
        final String disbursedDate = readAsDate(DISBURSED_DATE_COL, row);
        final String paymentType = readAsString(DISBURSED_PAYMENT_TYPE_COL, row);
        final String paymentTypeId = getIdByName(this.workbook.getSheet("Extras"), paymentType).toString();
        if (!disbursedDate.equals("")) { return new LoanDisbursal(disbursedDate, paymentTypeId, row.getRowNum()); }
        return null;
    }

    private Transaction parseAsLoanRepayment(final Row row) {
        final String repaymentAmount = readAsDouble(TOTAL_AMOUNT_REPAID_COL, row).toString();
        final String lastRepaymentDate = readAsDate(LAST_REPAYMENT_DATE_COL, row);
        final String repaymentType = readAsString(REPAYMENT_TYPE_COL, row);
        final String repaymentTypeId = getIdByName(this.workbook.getSheet("Extras"), repaymentType).toString();
        if (!repaymentAmount.equals("0.0")) { return new Transaction(repaymentAmount, lastRepaymentDate, repaymentTypeId, row.getRowNum()); }
        return null;
    }

    @Override
    public Result upload() {
        final Result result = new Result();
        final Sheet loanSheet = this.workbook.getSheet("Loans");
        this.restClient.createAuthToken();
        int progressLevel = 0;
        String loanId;
        for (int i = 0; i < this.loans.size(); i++) {
            final Row row = loanSheet.getRow(this.loans.get(i).getRowIndex());
            final Cell errorReportCell = row.createCell(FAILURE_REPORT_COL);
            final Cell statusCell = row.createCell(STATUS_COL);
            loanId = "";
            try {
                String response = "";
                final String status = this.loans.get(i).getStatus();
                progressLevel = getProgressLevel(status);

                if (progressLevel == 0) {
                    response = uploadLoan(i);
                    loanId = getLoanId(response);
                    progressLevel = 1;
                } else {
                    loanId = readAsInt(LOAN_ID_COL, loanSheet.getRow(this.loans.get(i).getRowIndex()));
                }

                if (progressLevel <= 1) {
                    progressLevel = uploadLoanApproval(loanId, i);
                }

                if (progressLevel <= 2) {
                    progressLevel = uploadLoanDisbursal(loanId, i);
                }

                if (this.loanRepayments.get(i) != null) {
                    progressLevel = uploadLoanRepayment(loanId, i);
                }

                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(this.workbook, IndexedColors.LIGHT_GREEN));
            } catch (final RuntimeException re) {
                final String message = parseStatus(re.getMessage());
                String status = "";

                if (progressLevel == 0) {
                    status = "Creation";
                } else if (progressLevel == 1) {
                    status = "Approval";
                } else if (progressLevel == 2) {
                    status = "Disbursal";
                } else if (progressLevel == 3) {
                    status = "Repayment";
                }
                statusCell.setCellValue(status + " failed.");
                statusCell.setCellStyle(getCellStyle(this.workbook, IndexedColors.RED));

                if (progressLevel > 0) {
                    row.createCell(LOAN_ID_COL).setCellValue(Integer.parseInt(loanId));
                }

                errorReportCell.setCellValue(message);
                result.addError("Row = " + this.loans.get(i).getRowIndex() + " ," + message);
            }
        }

        setReportHeaders(loanSheet);
        return result;
    }

    private int getProgressLevel(final String status) {
        if (status.equals("") || status.equals("Creation failed.")) {
            return 0;
        } else if (status.equals("Approval failed.")) {
            return 1;
        } else if (status.equals("Disbursal failed.")) {
            return 2;
        } else if (status.equals("Repayment failed.")) { return 3; }
        return 0;
    }

    private String uploadLoan(final int rowIndex) {
        final Gson gson = new Gson();
        final String payload = gson.toJson(this.loans.get(rowIndex));
        logger.info(payload);
        final String response = this.restClient.post("loans", payload);

        return response;
    }

    private String getLoanId(final String response) {
        final JsonParser parser = new JsonParser();
        final JsonObject obj = parser.parse(response).getAsJsonObject();
        return obj.get("loanId").getAsString();
    }

    private Integer uploadLoanApproval(final String loanId, final int rowIndex) {
        if (this.approvalDates.get(rowIndex) != null) {
            final Gson gson = new Gson();
            final String payload = gson.toJson(this.approvalDates.get(rowIndex));
            logger.info(payload);
            this.restClient.post("loans/" + loanId + "?command=approve", payload);
        }
        return 2;
    }

    private Integer uploadLoanDisbursal(final String loanId, final int rowIndex) {
        if (this.approvalDates.get(rowIndex) != null && this.disbursalDates.get(rowIndex) != null) {
            final Gson gson = new Gson();
            final String payload = gson.toJson(this.disbursalDates.get(rowIndex));
            logger.info(payload);
            this.restClient.post("loans/" + loanId + "?command=disburse", payload);
        }
        return 3;
    }

    private Integer uploadLoanRepayment(final String loanId, final int rowIndex) {
        final Gson gson = new Gson();
        final String payload = gson.toJson(this.loanRepayments.get(rowIndex));
        logger.info(payload);
        this.restClient.post("loans/" + loanId + "/transactions?command=repayment", payload);
        return 4;
    }

    private void setReportHeaders(final Sheet sheet) {
        sheet.setColumnWidth(STATUS_COL, 4000);
        final Row rowHeader = sheet.getRow(0);
        writeString(STATUS_COL, rowHeader, "Status");
        writeString(LOAN_ID_COL, rowHeader, "Loan ID");
        writeString(FAILURE_REPORT_COL, rowHeader, "Report");
    }

    public List<Loan> getLoans() {
        return this.loans;
    }

    public List<Approval> getApprovalDates() {
        return this.approvalDates;
    }

    public List<LoanDisbursal> getDisbursalDates() {
        return this.disbursalDates;
    }

    public List<Transaction> getLoanRepayments() {
        return this.loanRepayments;
    }
}
