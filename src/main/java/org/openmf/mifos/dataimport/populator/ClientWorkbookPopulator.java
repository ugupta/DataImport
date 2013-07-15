package org.openmf.mifos.dataimport.populator;

import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.openmf.mifos.dataimport.http.RestClient;

public class ClientWorkbookPopulator extends AbstractWorkbookPopulator {
	
	public static final int FIRST_NAME_COL = 0;
    public static final int LAST_NAME_COL = 1;
    public static final int MIDDLE_NAME_COL = 2;
    public static final int OFFICE_NAME_COL = 3;
    public static final int STAFF_NAME_COL = 4;
    public static final int EXTERNAL_ID_COL = 5;
    public static final int ACTIVATION_DATE_COL = 6;
    public static final int ACTIVE_COL = 7;

	private final RestClient client;
	
	private OfficeSheetPopulator osp;
	
	private PersonnelSheetPopulator psp;

    public ClientWorkbookPopulator(RestClient client) {
        this.client = client;
    }
    
    @Override
    public void downloadAndParse() {
        osp = new OfficeSheetPopulator(client);
        osp.downloadAndParse();
        psp = new PersonnelSheetPopulator(client);
        psp.downloadAndParse();
    }

    @Override
    public void populate(Workbook workbook) {
        osp.populate(workbook);
        psp.populate(workbook);
        Sheet clientSheet = workbook.createSheet("Clients");
        setLayout(clientSheet);
        setRules(clientSheet);
    }
    
    public void setLayout(Sheet worksheet) {
    	worksheet.setColumnWidth(0, 7000);
        worksheet.setColumnWidth(1, 7000);
        worksheet.setColumnWidth(2, 7000);
        worksheet.setColumnWidth(3, 7000);
        worksheet.setColumnWidth(4, 7000);
        worksheet.setColumnWidth(5, 5000);
        worksheet.setColumnWidth(6, 4000);
        worksheet.setColumnWidth(7, 2000);
        Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        writeString(FIRST_NAME_COL, rowHeader, "First Name");
        writeString(LAST_NAME_COL, rowHeader, "Last Name");
        writeString(MIDDLE_NAME_COL, rowHeader, "Middle Name");
        writeString(OFFICE_NAME_COL, rowHeader, "Office Name");
        writeString(STAFF_NAME_COL, rowHeader, "Staff Name");
        writeString(EXTERNAL_ID_COL, rowHeader, "External ID");
        writeString(ACTIVATION_DATE_COL, rowHeader, "Activation Date");
        writeString(ACTIVE_COL, rowHeader, "Active");
    }
    
    public void setRules(Sheet worksheet) {
    	CellRangeAddressList officeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), 3, 3);
    	CellRangeAddressList staffNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), 4, 4);
    	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
    	Name officeGroup = worksheet.getWorkbook().createName();
    	Name staffGroup = worksheet.getWorkbook().createName();
    	officeGroup.setNameName("Office");
    	officeGroup.setRefersToFormula("Offices!$B$2:$B$"+osp.getOfficeSize()+1);
    	staffGroup.setNameName("Personnel");
    	staffGroup.setRefersToFormula("Staff!$B$2:$B$"+psp.getPersonnelSize()+1);
    	DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
    	DataValidationConstraint staffNameConstraint = validationHelper.createFormulaListConstraint("Personnel");
    	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
    	officeValidation.setSuppressDropDownArrow(false);  
    	DataValidation staffValidation = validationHelper.createValidation(staffNameConstraint, staffNameRange);
    	staffValidation.setSuppressDropDownArrow(false);
        worksheet.addValidationData(officeValidation);
        worksheet.addValidationData(staffValidation);
        //TODO: Set rules for all columns.
    }
    
}
