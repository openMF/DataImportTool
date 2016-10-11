package org.openmf.mifos.dataimport.dto.loan;

import java.util.Locale;

public class LoanDisbursal {

	 private final transient Integer rowIndex;

	 private final String actualDisbursementDate;
	
	 private final String paymentTypeId;
	
	 private final String dateFormat;
	
	 private final Locale locale;
	 
	 private final String note;
	 
	 private final String accountNumber;
	 
	 private final String routingCode;
	 
	 private final String receiptNumber;
	 
	 private final String bankNumber;
	 
	 private final String checkNumber;
	 
	 public LoanDisbursal(String actualDisbursementDate, String paymentTypeId, Integer rowIndex) {
	        this.actualDisbursementDate = actualDisbursementDate;
	        this.paymentTypeId = paymentTypeId;
	        this.rowIndex = rowIndex;
	        this.dateFormat = "dd MMMM yyyy";
	        this.locale = Locale.ENGLISH;
	        this.note = "";
	        this.accountNumber = null;
	        this.routingCode = null;
	        this.receiptNumber = null;
	        this.bankNumber = null;
	        this.checkNumber = null;
	    }
	 
	 public String getActualDisbursementDate() {
		    return actualDisbursementDate;
	  }
	 
	 public String getPaymentTypeId() {
		 return paymentTypeId;
	 }
	 
	  public Locale getLocale() {
	    	return locale;
	    }
	    
	    public String getDateFormat() {
	    	return dateFormat;
	    }

	    public Integer getRowIndex() {
	        return rowIndex;
	    }
	    
	    public String getNote() {
	    	return note;
	    }
	    
	    public String getAccountNumber() {
	    	return accountNumber;
	    }
	    
	    public String getRoutingCode() {
	    	return routingCode;
	    }
	    
	    public String getReceiptNumber() {
	    	return receiptNumber;
	    }
	    
	    public String getBankNumber() {
	    	return bankNumber;
	    }
	    
	    public String getCheckNumber() {
	    	return checkNumber;
	    }
}
