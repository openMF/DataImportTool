package org.openmf.mifos.dataimport.dto;

import java.util.Locale;

public class ClosingOfSavingsAccounts {
	
	 private final transient Integer rowIndex;
	 
	 private final transient Integer accountId;

	 private final String closedOnDate;
	 
	 private final String onAccountClosureId;
	 
	 private final String toSavingsAccountId;
	
	 private final String dateFormat;
	 
	 private final String accountType;
	
	 private final Locale locale;
	 
	 private final String note;
	 
	 public ClosingOfSavingsAccounts(Integer accountId, String closedOnDate, String onAccountClosureId,String toSavingsAccountId, String accountType,Integer rowIndex ) {
	        this.accountId = accountId;
		 	this.closedOnDate = closedOnDate;
	        this.onAccountClosureId = onAccountClosureId;
	        this.toSavingsAccountId = toSavingsAccountId;
	        this.accountType=accountType;
	        this.rowIndex = rowIndex;
	        this.dateFormat = "dd MMMM yyyy";
	        this.locale = Locale.ENGLISH;
	        this.note = "";
	    }
	 
	  public String getClosedOnDate() {
		    return closedOnDate;
	  }
	 
	  public Locale getLocale() {
	    	return locale;
	    }
	    
	    public String getDateFormat() {
	    	return dateFormat;
	    }

		public String getOnAccountClosureId() {
			return onAccountClosureId;
		}

		public String getToSavingsAccountId() {
			return toSavingsAccountId;
		}

		public Integer getRowIndex() {
			return rowIndex;
		}

		public String getNote() {
			return note;
		}

		public Integer getAccountId() {
			return accountId;
		}

		public String getAccountType() {
			return accountType;
		}
	

}
