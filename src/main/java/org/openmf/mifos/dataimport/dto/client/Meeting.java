package org.openmf.mifos.dataimport.dto.client;

import java.util.Locale;

//Meeting
public class Meeting {
	
    private final transient Integer rowIndex;
    
    private transient String groupId;
    
    private final String dateFormat;
    
    private final Locale locale;
    
    private final String description;
    
    private final Integer typeId;
    
    private String title;
    
    private final String startDate;
    
    private final String repeating;
    
    private final String repeats;
    
    private final String repeatsEvery;
    
    public Meeting(String startDate, String repeating, String repeats, String repeatsEvery, Integer rowIndex ) {
        this.startDate = startDate;
        this.repeating = repeating;
        this.repeats = repeats;
        this.repeatsEvery = repeatsEvery;
        this.rowIndex = rowIndex;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
        this.description = "";
        this.typeId = 1;
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

	public Integer getTypeId() {
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

	public String getRepeats() {
		return repeats;
	}

	public String getRepeatsEvery() {
		return repeatsEvery;
	}

	public String getGroupId() {
		return groupId;
	}

}
