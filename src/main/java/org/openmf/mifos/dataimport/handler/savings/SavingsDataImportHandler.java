package org.openmf.mifos.dataimport.handler.savings;

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
import org.openmf.mifos.dataimport.dto.savings.GroupSavingsAccount;
import org.openmf.mifos.dataimport.dto.savings.SavingsAccount;
import org.openmf.mifos.dataimport.dto.savings.SavingsActivation;
import org.openmf.mifos.dataimport.handler.AbstractDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;

import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class SavingsDataImportHandler extends AbstractDataImportHandler {

     

    @SuppressWarnings("CPD-START")
    private static final int SAVINGS_TYPE_COL = 1;
    private static final int CLIENT_NAME_COL = 2;
    private static final int PRODUCT_COL = 3;
    private static final int FIELD_OFFICER_NAME_COL = 4;
    private static final int SUBMITTED_ON_DATE_COL = 5;
    private static final int APPROVED_DATE_COL = 6;
    private static final int ACTIVATION_DATE_COL = 7;
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
    private static final int OVER_DRAFT_LIMIT_COL = 21;
    private static final int EXTERNAL_ID_COL = 22;
    private static final int STATUS_COL = 23;
    private static final int SAVINGS_ID_COL = 24;
    private static final int FAILURE_REPORT_COL = 25;
    private static final int CHARGE_ID_1 = 34;
    private static final int CHARGE_AMOUNT_1 = 35;
    private static final int CHARGE_DUE_DATE_1 = 36;
    private static final int CHARGE_ID_2 = 37;
    private static final int CHARGE_AMOUNT_2 = 38;
    private static final int CHARGE_DUE_DATE_2 = 39;


    @SuppressWarnings("CPD-END")
    private final RestClient restClient;

    private final Workbook workbook;

    private List<SavingsAccount> savings = new ArrayList<SavingsAccount>();
    private List<Approval> approvalDates = new ArrayList<Approval>();
    private List<SavingsActivation> activationDates = new ArrayList<SavingsActivation>();

    public SavingsDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
    }

    @Override
    public Result parse() {
        Result result = new Result();
        Sheet savingsSheet = workbook.getSheet("Savings");
        Integer noOfEntries = getNumberOfRows(savingsSheet, 0);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = savingsSheet.getRow(rowIndex);
                if (isNotImported(row, STATUS_COL)) {
                    savings.add(parseAsSavings(row));
                    approvalDates.add(parseAsSavingsApproval(row));
                    activationDates.add(parseAsSavingsActivation(row));
                }
            } catch (RuntimeException re) {
                result.addError("Row = " + rowIndex + " , " + re.getMessage());
            }
        }
        return result;
    }

    private SavingsAccount parseAsSavings(Row row) {
        String status = readAsString(STATUS_COL, row);
        String productName = readAsString(PRODUCT_COL, row);
        String productId = getIdByName(workbook.getSheet("Products"), productName).toString();
        String fieldOfficerName = readAsString(FIELD_OFFICER_NAME_COL, row);
        String fieldOfficerId = getIdByName(workbook.getSheet("Staff"), fieldOfficerName).toString();
        String submittedOnDate = readAsDate(SUBMITTED_ON_DATE_COL, row);
        String nominalAnnualInterestRate = readAsInt(NOMINAL_ANNUAL_INTEREST_RATE_COL, row);
        String interestCompoundingPeriodType = readAsString(INTEREST_COMPOUNDING_PERIOD_COL, row);
        String interestCompoundingPeriodTypeId = "";
        if (interestCompoundingPeriodType.equalsIgnoreCase("Daily"))
            interestCompoundingPeriodTypeId = "1";
        else if (interestCompoundingPeriodType.equalsIgnoreCase("Monthly"))
            interestCompoundingPeriodTypeId = "4";
        else if (interestCompoundingPeriodType.equalsIgnoreCase("Quarterly"))
            interestCompoundingPeriodTypeId = "5";
        else if (interestCompoundingPeriodType.equalsIgnoreCase("Semi-Annual"))
        	interestCompoundingPeriodTypeId = "6";
        else if (interestCompoundingPeriodType.equalsIgnoreCase("Annually"))
        	interestCompoundingPeriodTypeId = "7";
        String interestPostingPeriodType = readAsString(INTEREST_POSTING_PERIOD_COL, row);
        String interestPostingPeriodTypeId = "";
        if (interestPostingPeriodType.equalsIgnoreCase("Monthly"))
            interestPostingPeriodTypeId = "4";
        else if (interestPostingPeriodType.equalsIgnoreCase("Quarterly"))
            interestPostingPeriodTypeId = "5";
        else if (interestPostingPeriodType.equalsIgnoreCase("Annually"))
            interestPostingPeriodTypeId = "7";
        else if (interestPostingPeriodType.equalsIgnoreCase("BiAnnual"))
        	interestPostingPeriodTypeId = "6";
        String interestCalculationType = readAsString(INTEREST_CALCULATION_COL, row);
        String interestCalculationTypeId = "";
        if (interestCalculationType.equalsIgnoreCase("Daily Balance"))
            interestCalculationTypeId = "1";
        else if (interestCalculationType.equalsIgnoreCase("Average Daily Balance")) interestCalculationTypeId = "2";
        String interestCalculationDaysInYearType = readAsString(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, row);
        String interestCalculationDaysInYearTypeId = "";
        if (interestCalculationDaysInYearType.equalsIgnoreCase("360 Days"))
            interestCalculationDaysInYearTypeId = "360";
        else if (interestCalculationDaysInYearType.equalsIgnoreCase("365 Days")) interestCalculationDaysInYearTypeId = "365";
        String minRequiredOpeningBalance = readAsInt(MIN_OPENING_BALANCE_COL, row);
        String lockinPeriodFrequency = readAsInt(LOCKIN_PERIOD_COL, row).toString();
        String lockinPeriodFrequencyType = readAsString(LOCKIN_PERIOD_FREQUENCY_COL, row);
        String lockinPeriodFrequencyTypeId = "";
        if (lockinPeriodFrequencyType.equalsIgnoreCase("Days"))
            lockinPeriodFrequencyTypeId = "0";
        else if (lockinPeriodFrequencyType.equalsIgnoreCase("Weeks"))
            lockinPeriodFrequencyTypeId = "1";
        else if (lockinPeriodFrequencyType.equalsIgnoreCase("Months"))
            lockinPeriodFrequencyTypeId = "2";
        else if (lockinPeriodFrequencyType.equalsIgnoreCase("Years")) lockinPeriodFrequencyTypeId = "3";
        String applyWithdrawalFeeForTransfers = readAsBoolean(APPLY_WITHDRAWAL_FEE_FOR_TRANSFERS, row).toString();
        String savingsType = readAsString(SAVINGS_TYPE_COL, row).toLowerCase(Locale.ENGLISH);
        String clientOrGroupName = readAsString(CLIENT_NAME_COL, row);

        String externalId = readAsString(EXTERNAL_ID_COL, row);
        List<Charge> charges = new ArrayList<Charge>();

        
        String allowOverdraft = readAsBoolean(ALLOW_OVER_DRAFT_COL, row).toString();
        String overdraftLimit = readAsInt(OVER_DRAFT_LIMIT_COL, row);
       
        String charge1 = readAsString(CHARGE_ID_1, row);
        String charge2 = readAsString(CHARGE_ID_2, row);
       
        if (!charge1.equalsIgnoreCase("")) {
            charges.add(new Charge(readAsString(CHARGE_ID_1, row), readAsDouble(CHARGE_AMOUNT_1, row), readAsDate(CHARGE_DUE_DATE_1, row)));
        }

        if (!charge2.equalsIgnoreCase("")) {

            charges.add(new Charge(readAsString(CHARGE_ID_2, row), readAsDouble(CHARGE_AMOUNT_2, row), readAsDate(CHARGE_DUE_DATE_2, row)));
        }


        if (savingsType.equals("individual")) {
            String clientId = getIdByName(workbook.getSheet("Clients"), clientOrGroupName).toString();
            return new SavingsAccount(clientId, productId, fieldOfficerId, submittedOnDate, nominalAnnualInterestRate,
                    interestCompoundingPeriodTypeId, interestPostingPeriodTypeId, interestCalculationTypeId,
                    interestCalculationDaysInYearTypeId, minRequiredOpeningBalance, lockinPeriodFrequency, lockinPeriodFrequencyTypeId,
                    applyWithdrawalFeeForTransfers, row.getRowNum(), status, externalId, charges, allowOverdraft, overdraftLimit);
        }
        String groupId = getIdByName(workbook.getSheet("Groups"), clientOrGroupName).toString();
        return new GroupSavingsAccount(groupId, productId, fieldOfficerId, submittedOnDate, nominalAnnualInterestRate,
                interestCompoundingPeriodTypeId, interestPostingPeriodTypeId, interestCalculationTypeId,
                interestCalculationDaysInYearTypeId, minRequiredOpeningBalance, lockinPeriodFrequency, lockinPeriodFrequencyTypeId,
                applyWithdrawalFeeForTransfers, row.getRowNum(), status, externalId, charges, allowOverdraft, overdraftLimit);
    }

    private Approval parseAsSavingsApproval(Row row) {
        String approvalDate = readAsDate(APPROVED_DATE_COL, row);
        if (!approvalDate.equals(""))
            return new Approval(approvalDate, row.getRowNum());
        else
            return null;
    }

    private SavingsActivation parseAsSavingsActivation(Row row) {
        String activationDate = readAsDate(ACTIVATION_DATE_COL, row);
        if (!activationDate.equals(""))
            return new SavingsActivation(activationDate, row.getRowNum());
        else
            return null;
    }

    @Override
    public Result upload() {
        Result result = new Result();
        Sheet savingsSheet = workbook.getSheet("Savings");
        restClient.createAuthToken();
        int progressLevel = 0;
        String savingsId;
        for (int i = 0; i < savings.size(); i++) {
            Row row = savingsSheet.getRow(savings.get(i).getRowIndex());
            Cell statusCell = row.createCell(STATUS_COL);
            Cell errorReportCell = row.createCell(FAILURE_REPORT_COL);
            savingsId = "";
            try {
                String response = "";
                String status = savings.get(i).getStatus();
                progressLevel = getProgressLevel(status);

                if (progressLevel == 0) {
                    response = uploadSavings(i);
                    savingsId = getSavingsId(response);
                    progressLevel = 1;
                } else
                    savingsId = readAsInt(SAVINGS_ID_COL, savingsSheet.getRow(savings.get(i).getRowIndex()));

                if (progressLevel <= 1) progressLevel = uploadSavingsApproval(savingsId, i);

                if (progressLevel <= 2) progressLevel = uploadSavingsActivation(savingsId, i);

                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (RuntimeException re) {
                String message = parseStatus(re.getMessage());
                String status = "";

                if (progressLevel == 0)
                    status = "Creation";
                else if (progressLevel == 1)
                    status = "Approval";
                else if (progressLevel == 2) status = "Activation";
                statusCell.setCellValue(status + " failed.");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));

                if (progressLevel > 0) row.createCell(SAVINGS_ID_COL).setCellValue(Integer.parseInt(savingsId));

                errorReportCell.setCellValue(message);
                result.addError("Row = " + savings.get(i).getRowIndex() + " ," + message);
            }
        }
        setReportHeaders(savingsSheet);
        return result;
    }

    private int getProgressLevel(String status) {
        if (status.equals("") || status.equals("Creation failed."))
            return 0;
        else if (status.equals("Approval failed."))
            return 1;
        else if (status.equals("Activation failed.")) return 2;
        return 0;
    }

    private String uploadSavings(int rowIndex) {
        Gson gson = new Gson();
        String payload = gson.toJson(savings.get(rowIndex));
        String response = restClient.post("savingsaccounts", payload);
        return response;
    }

    private String getSavingsId(String response) {
        JsonParser parser = new JsonParser();
        JsonObject obj = parser.parse(response).getAsJsonObject();
        return obj.get("savingsId").getAsString();
    }

    private Integer uploadSavingsApproval(String savingsId, int rowIndex) {
        Gson gson = new Gson();
        if (approvalDates.get(rowIndex) != null) {
            String payload = gson.toJson(approvalDates.get(rowIndex));
            restClient.post("savingsaccounts/" + savingsId + "?command=approve", payload);
        }
        return 2;
    }

    private Integer uploadSavingsActivation(String savingsId, int rowIndex) {
        Gson gson = new Gson();
        if (activationDates.get(rowIndex) != null) {
            String payload = gson.toJson(activationDates.get(rowIndex));
            restClient.post("savingsaccounts/" + savingsId + "?command=activate", payload);
        }
        return 3;
    }

    private void setReportHeaders(Sheet savingsSheet) {
        savingsSheet.setColumnWidth(STATUS_COL, 4000);
        Row rowHeader = savingsSheet.getRow(0);
        writeString(STATUS_COL, rowHeader, "Status");
        writeString(SAVINGS_ID_COL, rowHeader, "Savings ID");
        writeString(FAILURE_REPORT_COL, rowHeader, "Report");
    }

    public List<SavingsAccount> getSavings() {
        return savings;
    }

    public List<Approval> getApprovalDates() {
        return approvalDates;
    }

    public List<SavingsActivation> getActivationDates() {
        return activationDates;
    }

}
