package org.openmf.mifos.dataimport;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.runners.MockitoJUnitRunner;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.openmf.mifos.dataimport.populator.ClientWorkbookPopulator;

@RunWith(MockitoJUnitRunner.class)
public class ClientWorkbookPopulatorTest {

    // SUT - System Under Test
    ClientWorkbookPopulator populator;

    @Mock
	RestClient restClient;

	    @Test
	    public void shouldDownloadAndParseOfficesAndStaff() {

	        // create test data
	        String clientType = "individual";
	        Mockito.when(restClient.get("client")).thenReturn("{JSON STRING}");

	        // create SUT instance
	    	populator = new ClientWorkbookPopulator(clientType, restClient);

	    	// run
	    	Result result = populator.downloadAndParse();

	    	// verify/assert results
	    	Assert.assertTrue(result.isSuccess());
	    	Mockito.verify(restClient, Mockito.atLeastOnce()).get("client");
	    }

	    @Test
	    public void shouldPopulateClientWorkbook() {

            // create test data
	        String clientType = "individual";
	        Mockito.when(restClient.get("client")).thenReturn("{JSON STRING}");

            // create SUT instance
	    	ClientWorkbookPopulator populator = new ClientWorkbookPopulator(clientType, restClient);

	    	// run
	    	Result result = populator.downloadAndParse();
	    	Workbook book = new HSSFWorkbook();
	    	result = populator.populate(book);

	    	// verify/assert
	    	Assert.assertTrue(result.isSuccess());
            Mockito.verify(restClient, Mockito.atLeastOnce()).get("client");
	    }
}
