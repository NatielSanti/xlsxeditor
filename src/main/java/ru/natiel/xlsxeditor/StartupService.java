package ru.natiel.xlsxeditor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;

@Component
public class StartupService {

    @Autowired
    private JdbcTemplate jdbcTemplate;

    private File myFile;
    private XSSFSheet mySheet;

    private static final String SCRIPT =
            "select imsi, brand_cd from h_schema.device_mst where imsi in (%s) " +
            "UNION all " +
            "select imsi, brand_cd from k_schema.device_mst where imsi in (%s) " +
            "UNION all " +
            "select imsi, brand_cd from g_schema.device_mst where imsi in (%s)";

    public void start() throws IOException {
        System.out.println(" Application START !!!!!!!!!!!!!!");
        XSSFWorkbook myWorkBook = getXssfSheets();

        mySheet = myWorkBook.getSheet("Data");
        String imsiList = getImsiList(mySheet);

        List<Map<String, Object>> maps = jdbcTemplate.queryForList(String.format(SCRIPT, imsiList, imsiList, imsiList));
        Map<String, String> imsiMap = new HashMap<>();
        maps.forEach(x -> imsiMap.put(x.get("imsi").toString(), x.get("brand_cd").toString()));

        fillXLSX(myWorkBook, imsiMap);
        System.out.println("");

    }

    private void fillXLSX(XSSFWorkbook myWorkBook, Map<String, String> imsiMap) throws IOException {
        FileOutputStream os = new FileOutputStream(myFile);
        Iterator<Row> rowIterator = mySheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell imsiCell = row.getCell(4);
            Cell brandCell = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            brandCell.setCellValue(imsiMap.get(imsiCell.getStringCellValue()));
        }
		myWorkBook.write(os);
    }

    private String getImsiList(XSSFSheet mySheet) {
        StringBuilder result = new StringBuilder();
        Iterator<Row> rowIterator = mySheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(4);
            result.append("\'").append(cell.getStringCellValue()).append("\'");
            if(rowIterator.hasNext())
                result.append(",");
        }
        return result.toString();
    }

    private XSSFWorkbook getXssfSheets() throws IOException {
        myFile = new File("C:\\NATI\\projects\\xlsxeditor\\src\\main\\resources\\Summarised_SMS_usage_per_IMSI_(STCU).xlsx");
        FileInputStream fis = new FileInputStream(myFile);
        return new XSSFWorkbook(fis);
    }

    private class Line {
        String imsi;
        String brand;

        public Line (String imsi, String brand){
            this.imsi = imsi;
            this.brand = brand;
        }

        public String getImsi() { return imsi; }

        public String getBrand() { return brand; }
    }
}
