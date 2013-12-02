package org.openmf.mifos.dataimport.dto.loan;

import java.util.ArrayList;

public class LoanTimeline {
	
	private final ArrayList<Integer> actualDisbursementDate;
	
	public LoanTimeline(ArrayList<Integer> actualDisbursementDate) {
		this.actualDisbursementDate = actualDisbursementDate;
	}
	
	public ArrayList<Integer> getActualDisbursementDate() {
    	return this.actualDisbursementDate;
    }
	
	
}
