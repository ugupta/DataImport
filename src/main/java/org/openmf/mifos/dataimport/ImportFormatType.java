package org.openmf.mifos.dataimport;


public enum ImportFormatType {
    
    XLSX_OPEN ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    XLS ("application/vnd.ms-excel"),
    ODS ("application/vnd.oasis.opendocument.spreadsheet");
    
    
    private final String format;

    private ImportFormatType(String format) {
        this.format= format;
    }

    public String getFormat() {
        return format;
    }
}
