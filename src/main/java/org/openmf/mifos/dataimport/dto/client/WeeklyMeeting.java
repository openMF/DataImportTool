package org.openmf.mifos.dataimport.dto.client;


public class WeeklyMeeting extends Meeting{
	
    private final String repeatsOnDay;
    
    public WeeklyMeeting(String startDate, String repeating, String repeats, String repeatsEvery, String repeatsOnDay, Integer rowIndex  ) {
        super(startDate, repeating, repeats, repeatsEvery, rowIndex );
    	this.repeatsOnDay = repeatsOnDay;
    }

	public String getRepeatsOnDay() {
		return repeatsOnDay;
	}

}
