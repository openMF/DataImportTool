package org.openmf.mifos.dataimport.dto.loan;

public class GroupLoan extends Loan{
	
	private final String groupId;
	private final String externalId;
	public GroupLoan(String loanType, String groupId, String productId,	String loanOfficerId, String submittedOnDate, String fundId,
			String principal, String numberOfRepayments, String repaymentEvery,	String repaymentFrequencyType, String loanTermFrequency,
			String loanTermFrequencyType, String interestRatePerPeriod,	String expectedDisbursementDate, String amortizationType,
			String interestType, String interestCalculationPeriodType, String inArrearsTolerance, String transactionProcessingStrategyId,
			String graceOnPrincipalPayment, String graceOnInterestPayment, String graceOnInterestCharged, String interestChargedFromDate,
			String repaymentsStartingFromDate, Integer rowIndex, String status,String externalId,String linkAccountId) {
		
		super(loanType, null, productId, loanOfficerId, submittedOnDate, fundId, principal, numberOfRepayments, repaymentEvery, repaymentFrequencyType,
				loanTermFrequency, loanTermFrequencyType, interestRatePerPeriod, expectedDisbursementDate, amortizationType, interestType,
				interestCalculationPeriodType, inArrearsTolerance, transactionProcessingStrategyId, graceOnPrincipalPayment,
				graceOnInterestPayment, graceOnInterestCharged,	interestChargedFromDate, repaymentsStartingFromDate, rowIndex, status, externalId, null,null,linkAccountId );
		
		this.groupId = groupId;
		this.externalId=externalId;
	   }

	public String getGroupId() {
		return groupId;
	}

}
