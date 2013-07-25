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
import org.openmf.mifos.dataimport.dto.CorporateClient;
import org.openmf.mifos.dataimport.handler.ClientDataImportHandler;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;

@RunWith(MockitoJUnitRunner.class)
public class ClientImportHandlerTest {
    
    @Mock
    RestClient restClient;
    
    @Test
    public void shouldParseIndividualClientMSXls() throws IOException {
        
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("client/client.xls");
        Workbook book = new HSSFWorkbook(is);
        ClientDataImportHandler handler = new ClientDataImportHandler(book, restClient);
        Result result = handler.parse();
        Assert.assertTrue(result.isSuccess());
        Assert.assertEquals(1, handler.getClients().size());
        Client client = handler.getClients().get(0);
        Assert.assertEquals("David", client.getFirstName());
        Assert.assertEquals("Spade", client.getLastName());
        Assert.assertEquals(handler.getIdByName(book.getSheet("Offices"), "Branch_5").toString(), client.getOfficeId());
        Assert.assertEquals(handler.getIdByName(book.getSheet("Staff"), "Raul Albiol").toString(), client.getStaffId());
        Assert.assertEquals("19 May 2013", client.getActivationDate());
        Assert.assertEquals("true", client.isActive());
    }
    
    @Test
    public void shouldParseCorporateClientMSXls() throws IOException {
    	InputStream is = this.getClass().getClassLoader().getResourceAsStream("client/client-corporate.xls");
    	Workbook book = new HSSFWorkbook(is);
        ClientDataImportHandler handler = new ClientDataImportHandler(book, restClient);
        Result result = handler.parse();
        Assert.assertTrue(result.isSuccess());
        Assert.assertEquals(1, handler.getCorporateClients().size());
        CorporateClient client = handler.getCorporateClients().get(0);
        Assert.assertEquals("Remo Fernandez", client.getFullName());
        Assert.assertEquals(handler.getIdByName(book.getSheet("Offices"), "Branch_7").toString(), client.getOfficeId());
        Assert.assertEquals(handler.getIdByName(book.getSheet("Staff"), "Tomas Rosicky").toString(), client.getStaffId());
        Assert.assertEquals("19 May 2013", client.getActivationDate());
        Assert.assertEquals("true", client.isActive());
    }

}
