package org.openmf.mifos.dataimport.dto.savings;

import org.openmf.mifos.dataimport.dto.Currency;
import org.openmf.mifos.dataimport.dto.Type;

public class RecurringDepositProduct {
	
	private final Integer id;

    private final String name;
    
    private final String shortName;

    private final Currency currency;

    private final Double nominalAnnualInterestRate;
    
    private final Double minDepositAmount;
    
    private final Double maxDepositAmount;
    
    private final Double depositAmount;

    private final Type interestCompoundingPeriodType;

    private final Type interestPostingPeriodType;

    private final Type interestCalculationType;

    private final Type interestCalculationDaysInYearType;

    private final Integer lockinPeriodFrequency;

    private final Type lockinPeriodFrequencyType;
    
    private final String preClosurePenalApplicable;
    
    private final Double preClosurePenalInterest;
    
    private final Type preClosureInterestOnType;
    
    private final Integer minDepositTerm;
    
    private final Integer maxDepositTerm;
    
    private final Type minDepositTermType;
    
    private final Type maxDepositTermType;
    
    private final Integer inMultiplesOfDepositTerm;
    
    private final Type inMultiplesOfDepositTermType;
    
    private final String isMandatoryDeposit;
    
    private final String allowWithdrawal;
    
    private final String adjustAdvanceTowardsFuturePayments;
    
    public RecurringDepositProduct(Integer id, String name, String shortName, Currency currency, Double nominalAnnualInterestRate,
    		Type interestCompoundingPeriodType, Type interestPostingPeriodType, Type interestCalculationType, Type interestCalculationDaysInYearType,
            Integer lockinPeriodFrequency, Type lockinPeriodFrequencyType, Double minDepositAmount, Double maxDepositAmount,
            Double depositAmount, String preClosurePenalApplicable, Double preClosurePenalInterest, Type preClosureInterestOnType,
            Integer minDepositTerm, Integer maxDepositTerm, Type minDepositTermType, Type maxDepositTermType, Integer inMultiplesOfDepositTerm,
            Type inMultiplesOfDepositTermType, String isMandatoryDeposit, String allowWithdrawal, 
            String adjustAdvanceTowardsFuturePayments) {
    	
        this.id = id;
        this.name = name;
        this.shortName = shortName;
        this.minDepositAmount = minDepositAmount;
        this.maxDepositAmount = maxDepositAmount;
        this.depositAmount = depositAmount;
        this.preClosurePenalApplicable = preClosurePenalApplicable;
        this.preClosurePenalInterest = preClosurePenalInterest;
        this.preClosureInterestOnType = preClosureInterestOnType;
        this.minDepositTerm = minDepositTerm;
        this.maxDepositTerm = maxDepositTerm;
        this.minDepositTermType = minDepositTermType;
        this.maxDepositTermType = maxDepositTermType;
        this.inMultiplesOfDepositTerm = inMultiplesOfDepositTerm;
        this.inMultiplesOfDepositTermType = inMultiplesOfDepositTermType;
        this.currency = currency;
        this.nominalAnnualInterestRate = nominalAnnualInterestRate;
        this.interestCompoundingPeriodType = interestCompoundingPeriodType;
        this.interestPostingPeriodType = interestPostingPeriodType;
        this.interestCalculationType = interestCalculationType;
        this.interestCalculationDaysInYearType = interestCalculationDaysInYearType;
        this.lockinPeriodFrequency = lockinPeriodFrequency;
        this.lockinPeriodFrequencyType = lockinPeriodFrequencyType;
        this.allowWithdrawal = allowWithdrawal;
        this.isMandatoryDeposit = isMandatoryDeposit;
        this.adjustAdvanceTowardsFuturePayments = adjustAdvanceTowardsFuturePayments;
    }

	public Integer getId() {
		return id;
	}

	public String getName() {
		return name;
	}

	public Currency getCurrency() {
		return currency;
	}

	public Double getNominalAnnualInterestRate() {
		return nominalAnnualInterestRate;
	}

	public Type getInterestCompoundingPeriodType() {
		return interestCompoundingPeriodType;
	}

	public Type getInterestPostingPeriodType() {
		return interestPostingPeriodType;
	}

	public Type getInterestCalculationType() {
		return interestCalculationType;
	}

	public Integer getLockinPeriodFrequency() {
		return lockinPeriodFrequency;
	}

	public Type getInterestCalculationDaysInYearType() {
		return interestCalculationDaysInYearType;
	}

	public Type getLockinPeriodFrequencyType() {
		return lockinPeriodFrequencyType;
	}

	public String getPreClosurePenalApplicable() {
		return preClosurePenalApplicable;
	}

	public Double getPreClosurePenalInterest() {
		return preClosurePenalInterest;
	}

	public Type getPreClosureInterestOnType() {
		return preClosureInterestOnType;
	}

	public Integer getMinDepositTerm() {
		return minDepositTerm;
	}

	public Integer getMaxDepositTerm() {
		return maxDepositTerm;
	}

	public Type getMaxDepositTermType() {
		return maxDepositTermType;
	}

	public Type getMinDepositTermType() {
		return minDepositTermType;
	}

	public Integer getInMultiplesOfDepositTerm() {
		return inMultiplesOfDepositTerm;
	}

	public Type getInMultiplesOfDepositTermType() {
		return inMultiplesOfDepositTermType;
	}

	public Double getMinDepositAmount() {
		return minDepositAmount;
	}

	public Double getMaxDepositAmount() {
		return maxDepositAmount;
	}

	public Double getDepositAmount() {
		return depositAmount;
	}

	public String getShortName() {
		return shortName;
	}

	public String getIsMandatoryDeposit() {
		return isMandatoryDeposit;
	}

	public String getAllowWithdrawal() {
		return allowWithdrawal;
	}

	public String getAdjustAdvanceTowardsFuturePayments() {
		return adjustAdvanceTowardsFuturePayments;
	}
}
