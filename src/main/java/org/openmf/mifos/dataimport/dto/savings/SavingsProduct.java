package org.openmf.mifos.dataimport.dto.savings;

import org.openmf.mifos.dataimport.dto.Currency;
import org.openmf.mifos.dataimport.dto.Type;

public class SavingsProduct {

    private final Integer id;

    private final String name;

    private final Currency currency;

    private final Double nominalAnnualInterestRate;

    private final Type interestCompoundingPeriodType;

    private final Type interestPostingPeriodType;

    private final Type interestCalculationType;

    private final Type interestCalculationDaysInYearType;

    private final Double minRequiredOpeningBalance;

    private final Integer lockinPeriodFrequency;

    private final Type lockinPeriodFrequencyType;

    private final String withdrawalFeeForTransfers;
    
    private final String allowOverdraft;
    
    private final Integer overdraftLimit;

    public SavingsProduct(Integer id, String name, Currency currency, Double nominalAnnualInterestRate, Type interestCompoundingPeriodType,
            Type interestPostingPeriodType, Type interestCalculationType, Type interestCalculationDaysInYearType,
            Double minRequiredOpeningBalance, Integer lockinPeriodFrequency, Type lockinPeriodFrequencyType, String withdrawalFeeForTransfers,
            String allowOverdraft, Integer overdraftLimit) {
        this.id = id;
        this.name = name;
        this.currency = currency;
        this.nominalAnnualInterestRate = nominalAnnualInterestRate;
        this.interestCompoundingPeriodType = interestCompoundingPeriodType;
        this.interestPostingPeriodType = interestPostingPeriodType;
        this.interestCalculationType = interestCalculationType;
        this.interestCalculationDaysInYearType = interestCalculationDaysInYearType;
        this.minRequiredOpeningBalance = minRequiredOpeningBalance;
        this.lockinPeriodFrequency = lockinPeriodFrequency;
        this.lockinPeriodFrequencyType = lockinPeriodFrequencyType;
        this.withdrawalFeeForTransfers = withdrawalFeeForTransfers;
        this.overdraftLimit = overdraftLimit;
        this.allowOverdraft = allowOverdraft;
    }

    public Integer getId() {
        return this.id;
    }

    public String getName() {
        return this.name;
    }

    public Currency getCurrency() {
        return this.currency;
    }

    public Double getNominalAnnualInterestRate() {
        return this.nominalAnnualInterestRate;
    }

    public Type getInterestCompoundingPeriodType() {
        return this.interestCompoundingPeriodType;
    }

    public Type getInterestPostingPeriodType() {
        return this.interestPostingPeriodType;
    }

    public Type getInterestCalculationType() {
        return this.interestCalculationType;
    }

    public Type getInterestCalculationDaysInYearType() {
        return this.interestCalculationDaysInYearType;
    }

    public Double getMinRequiredOpeningBalance() {
        return this.minRequiredOpeningBalance;
    }

    public Integer getLockinPeriodFrequency() {
        return this.lockinPeriodFrequency;
    }

    public Type getLockinPeriodFrequencyType() {
        return this.lockinPeriodFrequencyType;
    }

	public String getWithdrawalFeeForTransfers() {
		return withdrawalFeeForTransfers;
	}

	public String getAllowOverdraft() {
		return allowOverdraft;
	}

	public Integer getOverdraftLimit() {
		return overdraftLimit;
	}
}
