package org.openmf.mifos.dataimport.dto.savings;

public class GroupSavingsAccount extends SavingsAccount{
	
	private final String groupId;

	public GroupSavingsAccount(String groupId, String productId, String fieldOfficerId, String submittedOnDate, String nominalAnnualInterestRate,
			String interestCompoundingPeriodType,	String interestPostingPeriodType, String interestCalculationType, String interestCalculationDaysInYearType,
			String minRequiredOpeningBalance, String lockinPeriodFrequency,	String lockinPeriodFrequencyType, String withdrawalFeeAmount,
			String withdrawalFeeType, String annualFeeAmount, String annualFeeOnMonthDay, String withdrawalFeeForTransfers,	Integer rowIndex, String status) {
		super(null, productId, fieldOfficerId, submittedOnDate,	nominalAnnualInterestRate, interestCompoundingPeriodType, interestPostingPeriodType,
				interestCalculationType, interestCalculationDaysInYearType, minRequiredOpeningBalance, lockinPeriodFrequency, lockinPeriodFrequencyType, withdrawalFeeAmount,
				withdrawalFeeType, annualFeeAmount, annualFeeOnMonthDay, withdrawalFeeForTransfers, rowIndex, status);
		this.groupId = groupId;
	}

	public String getGroupId() {
		return groupId;
	}

}
