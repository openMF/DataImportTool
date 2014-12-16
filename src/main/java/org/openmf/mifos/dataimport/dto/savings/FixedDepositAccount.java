package org.openmf.mifos.dataimport.dto.savings;

import java.util.List;
import java.util.Locale;

import org.openmf.mifos.dataimport.dto.Charge;

public class FixedDepositAccount {
	
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
    
    private final String depositAmount;

    private final String depositPeriod;

    private final String depositPeriodFrequencyId;
    
    private final String externalId;

    private final String dateFormat;

    private final Locale locale;
    
    private final List<Charge> charges;

    public FixedDepositAccount(String clientId, String productId, String fieldOfficerId, String submittedOnDate,
            String interestCompoundingPeriodType, String interestPostingPeriodType,
            String interestCalculationType, String interestCalculationDaysInYearType,
            String lockinPeriodFrequency, String lockinPeriodFrequencyType, String depositAmount, String depositPeriod,
            String depositPeriodFrequencyId, String externalId,List<Charge> charges, Integer rowIndex, String status) {
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
        this.depositAmount = depositAmount;
        this.depositPeriod = depositPeriod;
        this.depositPeriodFrequencyId = depositPeriodFrequencyId;
        this.externalId = externalId;
        this.rowIndex = rowIndex;
        this.status = status;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
        this.charges = charges;
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

	public String getDepositAmount() {
		return depositAmount;
	}

	public String getDepositPeriod() {
		return depositPeriod;
	}

	public String getDepositPeriodFrequencyId() {
		return depositPeriodFrequencyId;
	}

	public String getExternalId() {
		return externalId;
	}

	public List<Charge> getCharges() {
		return charges;
	}
}
