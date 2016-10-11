package org.openmf.mifos.dataimport.dto.savings;

import java.util.List;
import java.util.Locale;

import org.openmf.mifos.dataimport.dto.Charge;

public class SavingsAccount {

    private final transient Integer rowIndex;

    private final transient String status;

    private final String clientId;

    private final String fieldOfficerId;

    private final String productId;

    private final String submittedOnDate;

    private final String nominalAnnualInterestRate;

    private final String interestCompoundingPeriodType;

    private final String interestPostingPeriodType;

    private final String interestCalculationType;

    private final String interestCalculationDaysInYearType;

    private final String minRequiredOpeningBalance;

    private final String lockinPeriodFrequency;

    private final String lockinPeriodFrequencyType;

    private final String withdrawalFeeForTransfers;

    private final String dateFormat;

    private final String externalId;

    private final Locale locale;

    private final List<Charge> charges;
    
    private final String allowOverdraft;
    
    private final String overdraftLimit;

    public SavingsAccount(String clientId, String productId, String fieldOfficerId, String submittedOnDate,
            String nominalAnnualInterestRate, String interestCompoundingPeriodType, String interestPostingPeriodType,
            String interestCalculationType, String interestCalculationDaysInYearType, String minRequiredOpeningBalance,
            String lockinPeriodFrequency, String lockinPeriodFrequencyType, String withdrawalFeeForTransfers, Integer rowIndex,
            String status, String externalId, List<Charge> charges, String allowOverdraft ,String overdraftLimit ) {
        this.clientId = clientId;
        this.productId = productId;
        this.fieldOfficerId = fieldOfficerId;
        this.submittedOnDate = submittedOnDate;
        this.nominalAnnualInterestRate = nominalAnnualInterestRate;
        this.interestCompoundingPeriodType = interestCompoundingPeriodType;
        this.interestPostingPeriodType = interestPostingPeriodType;
        this.interestCalculationType = interestCalculationType;
        this.interestCalculationDaysInYearType = interestCalculationDaysInYearType;
        this.minRequiredOpeningBalance = minRequiredOpeningBalance;
        this.lockinPeriodFrequency = lockinPeriodFrequency;
        this.lockinPeriodFrequencyType = lockinPeriodFrequencyType;
        this.withdrawalFeeForTransfers = withdrawalFeeForTransfers;
        this.rowIndex = rowIndex;
        this.status = status;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
        this.externalId = externalId;
        this.charges = charges;
        this.allowOverdraft= allowOverdraft;
        this.overdraftLimit= overdraftLimit;
    }

    public String getClientId() {
        return clientId;
    }

    public String getFieldOfficerId() {
        return fieldOfficerId;
    }

    public String getProductId() {
        return productId;
    }

    public String getNominalAnnualInterestRate() {
        return nominalAnnualInterestRate;
    }

    public String getInterestCompoundingPeriodType() {
        return interestCompoundingPeriodType;
    }

    public String getInterestPostingPeriodType() {
        return interestPostingPeriodType;
    }

    public String getInterestCalculationType() {
        return interestCalculationType;
    }

    public String getInterestCalculationDaysInYearType() {
        return interestCalculationDaysInYearType;
    }

    public String getMinRequiredOpeningBalance() {
        return minRequiredOpeningBalance;
    }

    public String getLockinPeriodFrequency() {
        return lockinPeriodFrequency;
    }

    public String getLockinPeriodFrequencyType() {
        return lockinPeriodFrequencyType;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public Locale getLocale() {
        return locale;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public String getStatus() {
        return status;
    }

    public String getSubmittedOnDate() {
        return submittedOnDate;
    }

    public String getWithdrawalFeeForTransfers() {
        return withdrawalFeeForTransfers;
    }

    public String getExternalId() {
        return externalId;

    }

    public List<Charge> getCharges() {
        return charges;
    }

    public String getAllowOverdraft() {
        return allowOverdraft;
    }
    
    public String getOverdraftLimit() {
        return overdraftLimit;
    }
}
