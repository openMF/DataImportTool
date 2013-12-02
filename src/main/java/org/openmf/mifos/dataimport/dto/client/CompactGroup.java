package org.openmf.mifos.dataimport.dto.client;

import java.util.ArrayList;

public class CompactGroup {
    
    private final Integer id;
	
    private final String name;
    
    private final String officeName;
    
    private final ArrayList<Integer> activationDate;
	
    private final Boolean active;
    
    public CompactGroup(Integer id, String name,  String officeName, ArrayList<Integer> activationDate, Boolean active) {
		this.id = id;
        this.name = name;
        this.activationDate = activationDate;
        this.officeName = officeName;
        this.active = active;
    }

	public Integer getId() {
		return id;
	}

	public String getName() {
		return name;
	}

	public String getOfficeName() {
		return officeName;
	}

	public ArrayList<Integer> getActivationDate() {
		return activationDate;
	}

	public Boolean isActive() {
		return active;
	} 
}
