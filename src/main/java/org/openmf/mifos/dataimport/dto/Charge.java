package org.openmf.mifos.dataimport.dto;

public class Charge {
	
	private final String chargeId;
	private final Double amount;
	private final String dueDate;
	
	public Charge(String chargeId, Double amount, String dueDate) {
		this.chargeId = chargeId;
		this.amount = amount;
		this.dueDate = dueDate;
	}

	public String getChargeId() {
		return chargeId;
	}

	public Double getAmount() {
		return amount;
	}

	public String getDueDate() {
		return dueDate;
	}	

}
