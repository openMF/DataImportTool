package org.openmf.mifos.dataimport.dto.loan;

import java.util.Locale;

public class LoanDisbursToSavings {

    private final transient Integer rowIndex;

    private final String actualDisbursementDate;

    private final String dateFormat;

    private final Locale locale;

    private final String note;

    public LoanDisbursToSavings(String actualDisbursementDate, Integer rowIndex) {
        this.actualDisbursementDate = actualDisbursementDate;
        this.rowIndex = rowIndex;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
        this.note = "";
    }

    public String getActualDisbursementDate() {
        return actualDisbursementDate;
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

}
