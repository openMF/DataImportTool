package org.openmf.mifos.dataimport.dto.savings;

import java.util.List;
import java.util.Locale;

import org.openmf.mifos.dataimport.dto.Charge;

public class RecurringDepositAccount {

	private final transient Integer rowIndex;

    private final transient String status;

    private final String clientId;

    private final String fieldOfficerId;

    private final String productId;

    private final String submittedOnDate;

    private final String interestCompoundingPeriodType;

    private final String interestPostingPeriodType;

    private final String interestCalculationType;

    private final String interestCalculationDaysInYearType;

    private final String lockinPeriodFrequency;

    private final String lockinPeriodFrequencyType;
    
    private final String mandatoryRecommendedDepositAmount;

    private final String depositPeriod;

    private final String depositPeriodFrequencyId;
    
    private final String expectedFirstDepositOnDate;
    
    private final String recurringFrequency;
    
    private final String recurringFrequencyType;
    
    private final String isCalendarInherited;
    
    private final String isMandatoryDeposit;
    
    private final String allowWithdrawal;
    
    private final String adjustAdvanceTowardsFuturePayments;
    
    private final String externalId;
    
    private final List<Charge> charges;

    private final String dateFormat;

    private final Locale locale;

    public RecurringDepositAccount(String clientId, String productId, String fieldOfficerId, String submittedOnDate,
            String interestCompoundingPeriodType, String interestPostingPeriodType,
            String interestCalculationType, String interestCalculationDaysInYearType,
            String lockinPeriodFrequency, String lockinPeriodFrequencyType, String mandatoryRecommendedDepositAmount,
            String depositPeriod, String depositPeriodFrequencyId,
            String expectedFirstDepositOnDate, String recurringFrequency, String recurringFrequencyType,
            String isCalendarInherited, String isMandatoryDeposit, String allowWithdrawal,
            String adjustAdvanceTowardsFuturePayments, String externalId,List<Charge> charges,
            Integer rowIndex, String status) {
        this.clientId = clientId;
        this.productId = productId;
        this.fieldOfficerId = fieldOfficerId;
        this.submittedOnDate = submittedOnDate;
        this.interestCompoundingPeriodType = interestCompoundingPeriodType;
        this.interestPostingPeriodType = interestPostingPeriodType;
        this.interestCalculationType = interestCalculationType;
        this.interestCalculationDaysInYearType = interestCalculationDaysInYearType;
        this.lockinPeriodFrequency = lockinPeriodFrequency;
        this.lockinPeriodFrequencyType = lockinPeriodFrequencyType;
        this.mandatoryRecommendedDepositAmount = mandatoryRecommendedDepositAmount;
        this.depositPeriod = depositPeriod;
        this.depositPeriodFrequencyId = depositPeriodFrequencyId;
        this.expectedFirstDepositOnDate = expectedFirstDepositOnDate;
        this.isCalendarInherited = isCalendarInherited;
        this.isMandatoryDeposit = isMandatoryDeposit;
        this.allowWithdrawal = allowWithdrawal;
        this.adjustAdvanceTowardsFuturePayments = adjustAdvanceTowardsFuturePayments;
        this.recurringFrequency = recurringFrequency;
        this.recurringFrequencyType = recurringFrequencyType;
        this.externalId = externalId;
        this.charges = charges;
        this.rowIndex = rowIndex;
        this.status = status;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
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

	public String getMandatoryRecommendedDepositAmount() {
		return mandatoryRecommendedDepositAmount;
	}

	public String getDepositPeriod() {
		return depositPeriod;
	}

	public String getDepositPeriodFrequencyId() {
		return depositPeriodFrequencyId;
	}

	public String getExpectedFirstDepositOnDate() {
		return expectedFirstDepositOnDate;
	}

	public String getRecurringFrequency() {
		return recurringFrequency;
	}

	public String getIsCalendarInherited() {
		return isCalendarInherited;
	}

	public String getRecurringFrequencyType() {
		return recurringFrequencyType;
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

	public String getExternalId() {
		return externalId;
	}

	public List<Charge> getCharges() {
		return charges;
	}
}
