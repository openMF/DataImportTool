package org.openmf.mifos.dataimport.dto.client;

import java.util.Locale;

public class Center {
	
	 	private final transient Integer rowIndex;	    
	    private final transient String status;	    
	    private final String dateFormat;	    
	    private final Locale locale;	    
	    private final String name;	    	      
	    private final String officeId;	    
	    private final String staffId;	    
	    private final String externalId;	    
	    private final String active;	    
	    private final String activationDate;
	    
	    public Center(String name, String activationDate, String active, String externalId, String officeId, String staffId, Integer rowIndex, String status) {
	        this.name = name;
	        this.activationDate = activationDate;
	        this.active = active;
	        this.externalId = externalId;
	        this.officeId = officeId;
	        this.staffId = staffId;
	        this.rowIndex = rowIndex;
	        this.status = status;
	        this.dateFormat = "dd MMMM yyyy";
	        this.locale = Locale.ENGLISH;
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

		public String getName() {
			return name;
		}

		public String getOfficeId() {
			return officeId;
		}

		public String getStaffId() {
			return staffId;
		}

		public String getExternalId() {
			return externalId;
		}

		public String isActive() {
			return active;
		}

		public String getActivationDate() {
			return activationDate;
		}

		public String getStatus() {
			return status;
		}

}
