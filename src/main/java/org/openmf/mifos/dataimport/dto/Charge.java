package org.openmf.mifos.dataimport.dto;

public class Charge {
	
	private final String chargeId;
	private final String amount;
	private final String dueDate;
	
	public Charge(String chargeId, String amount, String dueDate) {
		this.chargeId = chargeId;
		this.amount = amount;
		this.dueDate = dueDate;
	}

	public String getChargeId() {
		return chargeId;
	}

	public String getAmount() {
		return amount;
	}

	public String getDueDate() {
		return dueDate;
	}	

}
