package org.openmf.mifos.dataimport.dto;


public class PaymentType {

    private final Integer id;
    
    private final String name;
    
	public PaymentType(Integer id, String name) {
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
