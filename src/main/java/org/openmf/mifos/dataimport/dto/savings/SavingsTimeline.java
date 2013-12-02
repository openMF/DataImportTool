package org.openmf.mifos.dataimport.dto.savings;

import java.util.ArrayList;

public class SavingsTimeline {

private final ArrayList<Integer> activatedOnDate;
	
	public SavingsTimeline(ArrayList<Integer> activatedOnDate) {
		this.activatedOnDate = activatedOnDate;
	}
	
	public ArrayList<Integer> getActivatedOnDate() {
    	return this.activatedOnDate;
    }
}
