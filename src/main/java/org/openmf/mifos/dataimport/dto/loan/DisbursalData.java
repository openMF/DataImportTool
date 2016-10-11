package org.openmf.mifos.dataimport.dto.loan;

public class DisbursalData {

	private final String linkAccountId;

	private final LoanDisbursal loanDisbursal;
	
	public DisbursalData(String linkAccountId, LoanDisbursal loanDisbursal) {
		this.linkAccountId = linkAccountId;
		this.loanDisbursal = loanDisbursal;
	}	
	
	public String getLinkAccountId() {
		return this.linkAccountId;
	}

	public LoanDisbursal getLoanDisbursal() {
		return loanDisbursal;
	}
	

}
