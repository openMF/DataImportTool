package org.openmf.mifos.dataimport.dto;

import java.util.Locale;

public class AddGuarantor {
	
private final transient Integer rowIndex;
private final transient Integer accountId;
	private final String dateFormat;
    private final Locale locale;
    private final String guarantorTypeId;
    private final String clientRelationshipTypeId;
    private final String entityId;
    private final String firstname;
    private final String lastname;
    private final String addressLine1;
    private final String addressLine2;
    private final String city;
    private final String dob;
    private final String zip;
    private final String savingsId;
    private final String amount;
    
   // private final String clientId;
   // private final String clientName;
   
    
    public AddGuarantor (String guarantorTypeId,String clientRelationshipTypeId,String entityId,String firstname,
    		String lastname,String addressLine1,String addressLine2,String city,String dob,String zip,String savingsId,String amount,Integer rowIndex,Integer accountId){
    	
    	this.guarantorTypeId=guarantorTypeId;
    	this.clientRelationshipTypeId=clientRelationshipTypeId;
    	this.entityId=entityId;
    	this.firstname=firstname;
    	this.lastname=lastname;
    	this.addressLine1=addressLine1;
    	this.addressLine2=addressLine2;
    	this.city=city;
    	this.dob=dob;
    	this.zip=zip;
    	 this.locale = Locale.ENGLISH;
    	 this.dateFormat = "dd MMMM yyyy";
    	this.rowIndex=rowIndex;
    	this.accountId=accountId;
    	this.savingsId=savingsId;
    	this.amount=amount;
    	//this.clientId=clientId;
    	//this.clientName=clientName;
    			
    	
    }

	public Integer getRowIndex() {
		return rowIndex;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public Locale getLocale() {
		return locale;
	}

	public String getGuarantorTypeId() {
		return guarantorTypeId;
	}

	public String getClientRelationshipTypeId() {
		return clientRelationshipTypeId;
	}

	public String getEntityId() {
		return entityId;
	}

	public String getFirstname() {
		return firstname;
	}

	public String getLastname() {
		return lastname;
	}

	public String getAddressLine1() {
		return addressLine1;
	}

	public String getAddressLine2() {
		return addressLine2;
	}

	public String getCity() {
		return city;
	}

	public String getDob() {
		return dob;
	}

	public String getZip() {
		return zip;
	}

	public Integer getAccountId() {
		return accountId;
	}

	public String getSavingsId() {
		return savingsId;
	}

	public String getAmount() {
		return amount;
	}

	/*public String getClientId() {
		return clientId;
	}

	public String getClientName() {
		return clientName;
	}*/
    
    
    
	
	
	

}
