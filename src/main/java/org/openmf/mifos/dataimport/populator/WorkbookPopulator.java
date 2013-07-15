package org.openmf.mifos.dataimport.populator;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookPopulator {

    void downloadAndParse();
    
    void populate(Workbook workbook);

}