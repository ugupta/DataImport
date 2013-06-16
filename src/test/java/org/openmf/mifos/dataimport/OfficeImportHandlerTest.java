package org.openmf.mifos.dataimport;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.Mock;
import org.mockito.runners.MockitoJUnitRunner;
import org.openmf.mifos.dataimport.dto.Office;
import org.openmf.mifos.dataimport.http.RestClient;

@RunWith(MockitoJUnitRunner.class)
public class OfficeImportHandlerTest {
    
    @Mock
    RestClient client;
    
    @Test
    public void shouldParseBasic2007MSXls() throws IOException {
        
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("office/office-ms-97-2003.xls");
        Workbook book = new HSSFWorkbook(is);
        OfficeDataImportHandler handler = new OfficeDataImportHandler(book.getSheetAt(0), client);
        Result result = handler.parse();
        Assert.assertTrue(result.isSuccess());
        Assert.assertEquals(1, handler.getOffices().size());
        Office office = handler.getOffices().get(0);
        Assert.assertEquals("Branch Office 1", office.getName());
        // TODO add more asserts
        
        
    }

}
