package org.openmf.mifos.dataimport;


public interface DataImportHandler {

    Result parse();
    
    Result upload();

}
