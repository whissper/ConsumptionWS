package org.whissper;

import javax.jws.WebService;
import javax.jws.WebMethod;
import javax.jws.WebParam;
import javax.jws.soap.SOAPBinding;
import javax.xml.ws.Holder;

/**
 * ConsumptionWS
 * @author SAV2
 */
@WebService(serviceName = "ConsumptionWS")
@SOAPBinding(parameterStyle = SOAPBinding.ParameterStyle.WRAPPED)
public class ConsumptionWS {

    @WebMethod(operationName = "loadXLSX")
    public void loadXLSX(@WebParam(name = "month", mode = WebParam.Mode.IN) String monthValue,
                         @WebParam(name = "year", mode = WebParam.Mode.IN) String yearValue,
                         @WebParam(name = "reference", mode = WebParam.Mode.OUT) Holder<String> refValue)
    {
        refValue.value = new ExcelLoaderEngine("C:/Server/data/htdocs/orgipu/php/getfile/", monthValue, yearValue).loadData();
    }
    
}
