import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;



public class ExcelWrite {


    public static void main(String[] args) {
        //Creates new workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Student data");

        //Create the data for the excel sheet
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{"ID", "FIRST NAME", "LAST NAME"});
        data.put("2", new Object[]{1, "Nicholas", "Sowers"});
        data.put("3", new Object[]{2, "Anthony", "Kwiatanowski"});
        data.put("4", new Object[]{3, "Andrew", "Sowers"});
        data.put("5", new Object[]{4, "Michael", "Sowers"});

        //Iterate over data and write it to sheet
        Set keySet = data.keySet();
        int rowNum = 0;
        for (Object key : keySet) {
            Row row = sheet.createRow(rowNum++);
            Object[] objArr = data.get(key);
            int cellNum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellNum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }//end of inner for loop
        }//end of 1st for loop
        try{
            FileOutputStream out = new FileOutputStream(
                    new File("C:\\Users\\sowe9352\\Documents"));
            workbook.write(out);
            out.close();
            System.out.println("testFile.xlsx Successfully created");
        }catch (FileNotFoundException e){
            e.printStackTrace();
        }catch(IOException e){
            e.printStackTrace();
        }
    }//end of main
}//end of ExcelWrite
