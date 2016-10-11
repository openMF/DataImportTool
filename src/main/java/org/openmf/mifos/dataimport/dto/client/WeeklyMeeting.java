package org.openmf.mifos.dataimport.dto.client;


public class WeeklyMeeting extends Meeting{
	
    private final String repeatsOnDay;
    
    public WeeklyMeeting(String startDate, String repeating, String frequency, String interval, String repeatsOnDay, Integer rowIndex  ) {
        super(startDate, repeating, frequency, interval, rowIndex );
    	this.repeatsOnDay = repeatsOnDay;
    }

	public String getRepeatsOnDay() {
		return repeatsOnDay;
	}

}
