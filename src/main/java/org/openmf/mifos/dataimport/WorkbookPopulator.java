package org.openmf.mifos.dataimport;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookPopulator {

    Result download();
    
    Result populate(Workbook workbook);

}