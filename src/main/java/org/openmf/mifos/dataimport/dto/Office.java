package org.openmf.mifos.dataimport.dto;

import java.util.ArrayList;

public class Office {
    
    private final Integer id;
    
    private final String name;
    
    private final String externalId;
    
    private final ArrayList<Integer> openingDate;
    
    private final String parentName;
	
    private final String hierarchy;

    public Office(Integer id, String name, String externalId, ArrayList<Integer> openingDate, String parentName, String hierarchy ) {
        this.id = id;
        this.name = name;
        this.parentName = parentName;
        this.externalId = externalId;
        this.openingDate = openingDate;
        this.hierarchy = hierarchy;
    }
    
    @Override
	public String toString() {
	   return "OfficeObject [id=" + id + ", name=" + name + ", externalId=" + externalId + ", openingDate=" + openingDate + ", parentName=" + parentName + "]";
	}
    
    public Integer getId() {
    	return this.id;
    }

    public String getName() {
        return this.name;
    }

    public String getParentName() {
        return this.parentName;
    }

    public String getExternalId() {
        return this.externalId;
    }

    public ArrayList<Integer> getOpeningDate() {
        return this.openingDate;
    }
    
    public String getHierarchy() {
        return this.hierarchy;
    }

}
