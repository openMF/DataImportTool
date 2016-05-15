package org.openmf.mifos.dataimport.dto.client;

import java.util.ArrayList;
import java.util.Locale;

import org.openmf.mifos.dataimport.utils.StringUtils;

public class Group {
	
	    private final transient Integer rowIndex;
	    private final transient String status;
	    private final String dateFormat;	    
	    private final Locale locale;	    
	    private final String name;	    
	    private final ArrayList<String> clientMembers;	    
	    private final String officeId;	    
	    private final String staffId;
	    private final String centerId;
	    private final String externalId;	    
	    private final String active;	    
	    private final String activationDate;
	    
	    public Group(String name, ArrayList<String> clientMembers, String activationDate, String active, String externalId, String officeId, String staffId, String centerId,Integer rowIndex, String status) {
	        this.name = name;
	        this.clientMembers = clientMembers;
	        this.activationDate = activationDate;
	        this.active = active;
	        this.externalId = externalId;
	        this.officeId = officeId;
	        this.staffId = staffId;
	        if(StringUtils.isBlank(centerId) || centerId.equalsIgnoreCase("0"))
	        	this.centerId=null;
	        else
	        	this.centerId= centerId;
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

		public ArrayList<String> getClientMembers() {
			return clientMembers;
		}

		public String getOfficeId() {
			return officeId;
		}

		public String getStaffId() {
			return staffId;
		}
		public String getCenterId() {
			return centerId;
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
