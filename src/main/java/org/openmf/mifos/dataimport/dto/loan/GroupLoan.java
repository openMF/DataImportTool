package org.openmf.mifos.dataimport.dto.loan;

import java.util.List;

public class GroupLoan extends Loan {

    private final String groupId;

    public GroupLoan(final String loanType, final String groupId, final String productId, final String loanOfficerId,
            final String submittedOnDate, final String fundId, final String principal, final String numberOfRepayments,
            final String repaymentEvery, final String repaymentFrequencyType, final String loanTermFrequency,
            final String loanTermFrequencyType, final String interestRatePerPeriod, final String expectedDisbursementDate,
            final String amortizationType, final String interestType, final String interestCalculationPeriodType,
            final String inArrearsTolerance, final String transactionProcessingStrategyId, final String graceOnPrincipalPayment,
            final String graceOnInterestPayment, final String graceOnInterestCharged, final String interestChargedFromDate,
            final String repaymentsStartingFromDate, final Integer rowIndex, final String status, final String externalId,
            final String fixedEmiAmount, final String maxOutstandingLoanBalance, final List<EmiDetail> disbursementData) {

        super(loanType, null, productId, loanOfficerId, submittedOnDate, fundId, principal, numberOfRepayments, repaymentEvery,
                repaymentFrequencyType, loanTermFrequency, loanTermFrequencyType, interestRatePerPeriod, expectedDisbursementDate,
                amortizationType, interestType, interestCalculationPeriodType, inArrearsTolerance, transactionProcessingStrategyId,
                graceOnPrincipalPayment, graceOnInterestPayment, graceOnInterestCharged, interestChargedFromDate,
                repaymentsStartingFromDate, rowIndex, status, externalId, fixedEmiAmount, maxOutstandingLoanBalance, disbursementData);

        this.groupId = groupId;
    }

    public String getGroupId() {
        return this.groupId;
    }

}
