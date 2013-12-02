package org.openmf.mifos.dataimport.dto.loan;


public class Fund {
	
    private final Integer id;
    
    private final String name;
    
	public Fund(Integer id, String name) {
		this.id = id;
		this.name = name;
	}
	
	public Integer getId() {
    	return this.id;
    }

    public String getName() {
        return this.name;
    }
}
