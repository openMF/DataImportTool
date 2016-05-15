package org.openmf.mifos.dataimport.dto.client;

import java.util.Locale;

//Meeting
public class Meeting {
	
    private final transient Integer rowIndex;    
    private transient String groupId; 
    private transient String centerId;
    private final String dateFormat;    
    private final Locale locale;    
    private final String description;    
    private final String typeId;    
    private String title;    
    private final String startDate;    
    private final String repeating;    
    private final String frequency;    
    private final String interval;
    
    public Meeting(String startDate, String repeating, String frequency, String interval, Integer rowIndex ) {
        this.startDate = startDate;
        this.repeating = repeating;
        this.frequency = frequency;
        this.interval = interval;
        this.rowIndex = rowIndex;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
        this.description = "";
        this.typeId = "1";
    }
    
    public void setGroupId(String groupId) {
    	this.groupId = groupId;
    }
    
    public void setTitle(String title) {
    	this.title = title;
    }

	public Integer getRowIndex() {
		return rowIndex;
	}

	public Locale getLocale() {
		return locale;
	}

	public String getDescription() {
		return description;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public String getTypeId() {
		return typeId;
	}

	public String getTitle() {
		return title;
	}

	public String getStartDate() {
		return startDate;
	}

	public String isRepeating() {
		return repeating;
	}

	public String getFrequency() {
		return frequency;
	}

	public String getInterval() {
		return interval;
	}

	public String getGroupId() {
		return groupId;
	}

	public String getCenterId() {
		return centerId;
	}

	public void setCenterId(String centerId) {
		this.centerId = centerId;
	}

}
