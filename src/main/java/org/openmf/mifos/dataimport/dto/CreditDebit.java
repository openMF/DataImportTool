package org.openmf.mifos.dataimport.dto;

public class CreditDebit {

	private final String glAccountId;
	private final String amount;
	
	
	public CreditDebit(String glAccountId, String amount) {
		this.glAccountId = glAccountId;
		this.amount = amount;
		
	}


	public String getGlAccountId() {
		return glAccountId;
	}


	public String getAmount() {
		return amount;
	}
}



