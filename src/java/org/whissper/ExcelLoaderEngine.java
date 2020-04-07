package org.whissper;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import org.apache.http.NameValuePair;
import org.apache.http.client.CookieStore;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.protocol.HttpClientContext;
import org.apache.http.impl.client.BasicCookieStore;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.impl.cookie.BasicClientCookie;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.protocol.BasicHttpContext;
import org.apache.http.protocol.HttpContext;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * ExcelLoader class
 * @author SAV2
 */
public class ExcelLoaderEngine 
{
    private String resultStr;
    
    private String path;
    private File xlsxFile;
    private String month;
    private String year;
    
    private CookieStore cookieStore;
    private CloseableHttpClient httpclient;
    private HttpContext httpContext;
    private HttpPost httpPost;
    private ArrayList<NameValuePair> nvps;
    private CloseableHttpResponse response;
    
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    
    private Row row;
    private Cell cell;
    private int rowNum = 0;
    
    private Double pages = 0.0;
    
    public ExcelLoaderEngine(String pathVal, String monthVal, String yearVal){
        this.path = pathVal;
        this.month = monthVal;
        this.year = yearVal;
        
        resultStr = "";
    }
    
    private void alertError(String context){
        resultStr = "ERROR|"+context;
    }
    
    private void doLogin(){
        cookieStore = new BasicCookieStore();
        
        BasicClientCookie cookie = new BasicClientCookie("PHPSESSID","");

        cookie.setDomain("kom-es01-app25");
        cookie.setPath("/");
        cookieStore.addCookie(cookie);

        httpclient = HttpClients.custom().setDefaultCookieStore(cookieStore).build();

        httpContext = new BasicHttpContext();
        httpContext.setAttribute(HttpClientContext.COOKIE_STORE, cookieStore);

        httpPost = new HttpPost("http://kom-es01-app25/orgipu/php/MainEntrance.php?action=login");

        nvps = new ArrayList<>();
        nvps.add(new BasicNameValuePair("id", "isuservalid"));
        nvps.add(new BasicNameValuePair("pwd", "guest"));
        nvps.add(new BasicNameValuePair("usr", "guest"));


        try {
            httpPost.setEntity(new UrlEncodedFormEntity(nvps));
        } catch (UnsupportedEncodingException ex) {
            alertError("Exception -- doLogin() -- : " + ex);
            return;
        } 

        try {
            response = httpclient.execute(httpPost, httpContext);
            EntityUtils.consume(response.getEntity());
            response.close();
        } catch (IOException ex){
            alertError("Exception -- doLogin() -- : " + ex);
        }
    }
    
    private ArrayList<String> getRowData(JsonArray jsonArr){
        ArrayList<String> currentRow = new ArrayList<>();
        
        if(jsonArr.size() != 0){
            currentRow.add(jsonArr.getAsJsonArray().get(0).getAsString());
            currentRow.add(jsonArr.getAsJsonArray().get(2).getAsString());
            currentRow.add(jsonArr.getAsJsonArray().get(4).getAsString());
            
            if(jsonArr.getAsJsonArray().get(5).getAsString().contains("NORMATIVE")){
                currentRow.add("Норматив");
            }else if(jsonArr.getAsJsonArray().get(5).getAsString().contains("ERROR")){
                currentRow.add("Ошибка");
            }else if(jsonArr.getAsJsonArray().get(5).getAsString().contains("BOILER")){
                currentRow.add("Бойлер");
            }else if(jsonArr.getAsJsonArray().get(5).getAsString().contains("HEATMETER")){
                currentRow.add("Теплосчетчик");
            }else if(jsonArr.getAsJsonArray().get(5).getAsString().contains("ZERO")){
                currentRow.add("0");
            }else {
                currentRow.add(jsonArr.getAsJsonArray().get(5).getAsString());
            }
            
            if(jsonArr.getAsJsonArray().get(6).getAsString().contains("NORMATIVE")){
                currentRow.add("Норматив");
            }else if(jsonArr.getAsJsonArray().get(6).getAsString().contains("ERROR")){
                currentRow.add("Ошибка");
            }else if(jsonArr.getAsJsonArray().get(6).getAsString().contains("ZERO")){
                currentRow.add("0");
            }else {
                currentRow.add(jsonArr.getAsJsonArray().get(6).getAsString());
            }
        }
        
        return currentRow;
    }
    
    private void fillExcelRow(ArrayList<String> rowData){
        if( !rowData.isEmpty() ){
            row = sheet.createRow(rowNum++);

            cell = row.createCell(0);
            cell.setCellValue(rowData.get(0));
            cell = row.createCell(1);
            cell.setCellValue(rowData.get(1));
            cell = row.createCell(2);
            cell.setCellValue(rowData.get(2));
            cell = row.createCell(3);
            cell.setCellValue(rowData.get(3));
            cell = row.createCell(4);
            cell.setCellValue(rowData.get(4));
        }
    }
    
    private void fillExcelHeader(){
        row = sheet.createRow(rowNum++);
                
        cell = row.createCell(0);
        cell.setCellValue("№ договора");
        cell = row.createCell(1);
        cell.setCellValue("Теплоустановка");
        cell = row.createCell(2);
        cell.setCellValue("№ прибора");
        cell = row.createCell(3);
        cell.setCellValue("Расход (м3)");
        cell = row.createCell(4);
        cell.setCellValue("Расход (Гкал)");
    }
    
    private void initWorkbook(){
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Расчет расходов за "+ month +"."+ year);
    }
    
    private void loadPages(){
        JsonParser jsonParser = new JsonParser();
        JsonObject jobject;
        
        initWorkbook();
        fillExcelHeader();
        
        httpPost = new HttpPost("http://kom-es01-app25/orgipu/php/MainEntrance.php?action=consumptions");
        //first page
        nvps.clear();
            
        nvps.add(new BasicNameValuePair("page", "0"));
        nvps.add(new BasicNameValuePair("id", ""));
        nvps.add(new BasicNameValuePair("hide_normative_vals", "0"));
        nvps.add(new BasicNameValuePair("heated_object_name", ""));
        nvps.add(new BasicNameValuePair("heated_object_id", ""));
        nvps.add(new BasicNameValuePair("device_num", ""));
        nvps.add(new BasicNameValuePair("contractnum", ""));
        nvps.add(new BasicNameValuePair("calc_year", year));
        nvps.add(new BasicNameValuePair("calc_month", month));

        try {
            httpPost.setEntity(new UrlEncodedFormEntity(nvps));
        } catch (UnsupportedEncodingException ex) {
            alertError("Exception -- loadPages() -- : " + ex);
            return;
        }

        try {
            response = httpclient.execute(httpPost, httpContext);

            jobject = jsonParser.parse( EntityUtils.toString(response.getEntity()) ).getAsJsonObject();
            response.close();

            pages = Math.ceil(jobject.get("countrows").getAsDouble()/jobject.get("perpage").getAsDouble());

            for(JsonElement rowItem : jobject.getAsJsonArray("rowitems"))
            {
                fillExcelRow( getRowData(rowItem.getAsJsonArray()) );
            }

        } catch (IOException ex) {
            alertError("Exception -- loadPages() -- : " + ex);
            return;
        }
        //last pages
        for(int i=1; i<pages.intValue(); i++){
            nvps.clear();
            
            nvps.add(new BasicNameValuePair("page", Integer.toString(i)));
            nvps.add(new BasicNameValuePair("id", ""));
            nvps.add(new BasicNameValuePair("hide_normative_vals", "0"));
            nvps.add(new BasicNameValuePair("heated_object_name", ""));
            nvps.add(new BasicNameValuePair("heated_object_id", ""));
            nvps.add(new BasicNameValuePair("device_num", ""));
            nvps.add(new BasicNameValuePair("contractnum", ""));
            nvps.add(new BasicNameValuePair("calc_year", year));
            nvps.add(new BasicNameValuePair("calc_month", month));
            
            try {
                httpPost.setEntity(new UrlEncodedFormEntity(nvps));
            } catch (UnsupportedEncodingException ex) {
                alertError("Exception -- loadPages() -- : " + ex);
                return;
            }
            
            try {
                response = httpclient.execute(httpPost, httpContext);
                
                jobject = jsonParser.parse( EntityUtils.toString(response.getEntity()) ).getAsJsonObject();
                response.close();
                
                for(JsonElement rowItem : jobject.getAsJsonArray("rowitems"))
                {
                    fillExcelRow( getRowData(rowItem.getAsJsonArray()) );
                }
                
            } catch (IOException ex) {
                alertError("Exception -- loadPages() -- : " + ex);
            }
        }
    }
    
    private void decorateTable(){
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);
    }
    
    private void createFile(){
        xlsxFile = new File(this.path + "Расчет_расходов_"+ month +"-"+ year +".xlsx");
        if(xlsxFile.getParentFile()!=null){ xlsxFile.getParentFile().mkdirs(); }// Will create parent directories if not exists
        try {
            xlsxFile.createNewFile();
        } catch (IOException ex) {
            alertError("Exception -- createFile() -- : " + ex);
        }
    }
    
    private void writeFile(){
        try {
            FileOutputStream outputStream = new FileOutputStream(xlsxFile);
            workbook.write(outputStream);
            workbook.close();
            resultStr = "php/getfile/" + xlsxFile.getName();
        } catch (FileNotFoundException ex) {
            alertError("Exception -- writeFile() -- : " + ex);
        } catch (IOException ex) {
            alertError("Exception -- writeFile() -- : " + ex);
        }
    }
    
    public String loadData(){
        doLogin();
        if( resultStr.contains("ERROR") ){
            return resultStr;
        }
        loadPages();
        if( resultStr.contains("ERROR") ){
            return resultStr;
        }
        decorateTable();
        if( resultStr.contains("ERROR") ){
            return resultStr;
        }
        createFile();
        if( resultStr.contains("ERROR") ){
            return resultStr;
        }
        writeFile();
        
        return resultStr;
    }
}
