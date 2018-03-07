package vrs;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

/**
 * Created by W3E64 on 3/7/2018.
 */
public class exceljava {


  public  HSSFWorkbook workbook;
    Map<String, Object[]> testresultdata;
    HSSFSheet sheet;
    int eid=0;


    public int getNumbervalue() {
        eid=eid+1;

        return eid;
    }
    public String convertnum(){

        String eidn=""+getNumbervalue();
        return eidn;
    }

    public void setupAfterSuite() {
        //write excel file and file name is// TestResult.xls

        Set<String> keyset = testresultdata.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = testresultdata.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof Date)
                    cell.setCellValue((Date) obj);
                else if (obj instanceof Boolean)
                    cell.setCellValue((Boolean) obj);
                else if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Double)
                    cell.setCellValue((Double) obj);
            }


        }
        try {

            FileOutputStream out = new FileOutputStream(new File("PageSpeed.xls"));
            workbook.write(out);
            out.close();
            //System.out.println("Excel written successfully..");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        //close the browser


    }
    public void before(){
        workbook = new HSSFWorkbook();
        //create a new work sheet
        sheet = workbook.createSheet("Test Result");
        testresultdata = new LinkedHashMap<String, Object[]>();
        //add test result excel file column header
        //write the header in the first row
        try {
        } catch (Exception e) {
            throw new IllegalStateException("Can't start Web Driver", e);
        }
        testresultdata.put(convertnum(), new Object[]{"","Homepage Mobile","Homepage Desktop","Listing Mobile","listing Desktop","Booking Details Mobile","Booking Details Mobile","Homaway Details Mobile","Homaway Details Desktop"});
    }


}
