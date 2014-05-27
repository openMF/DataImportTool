package org.openmf.mifos.dataimport.dto.loan;

public class EmiDetail {

    private final String expectedDisbursementDate;
    private final String principal;

    public EmiDetail(final String expectedDisbursementDate, final String principal) {
        this.expectedDisbursementDate = expectedDisbursementDate;
        this.principal = principal;
    }

    public String getExpectedDisbursementDate() {
        return expectedDisbursementDate;
    }

    public String getPrincipal() {
        return principal;
    }

}
