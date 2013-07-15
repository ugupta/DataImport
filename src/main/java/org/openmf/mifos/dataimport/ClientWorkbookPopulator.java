package org.openmf.mifos.dataimport;

import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.http.RestClient;

public class ClientWorkbookPopulator extends AbstractWorkbookPopulator {

	private final RestClient client;
	
	private OfficeSheetPopulator osp;

    public ClientWorkbookPopulator(RestClient client) {
        this.client = client;
    }
    
    @Override
    public Result download() {
        Result result = new Result();
        osp = new OfficeSheetPopulator(client);
        result = osp.download();
        return result;
    }

    @Override
    public Result populate(Workbook workbook) {
        Result result = new Result();
        result = osp.populate(workbook);
        return result;
    }
    
}
