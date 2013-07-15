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
import org.openmf.mifos.dataimport.dto.Client;
import org.openmf.mifos.dataimport.handler.ClientDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;

@RunWith(MockitoJUnitRunner.class)
public class ClientImportHandlerTest {
    
    @Mock
    RestClient restClient;
    
    @Test
    public void shouldParseBasic2007MSXls() throws IOException {
        
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("office/office-ms-97-2003.xls");
        Workbook book = new HSSFWorkbook(is);
        ClientDataImportHandler handler = new ClientDataImportHandler(book, restClient);
        Result result = handler.parse();
        Assert.assertTrue(result.isSuccess());
        Assert.assertEquals(1, handler.getClients().size());
        Client client = handler.getClients().get(0);
        Assert.assertEquals("Kirsten", client.getFirstName());
        // TODO add more asserts
        
        
    }

}
