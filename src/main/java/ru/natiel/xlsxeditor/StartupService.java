package ru.natiel.xlsxeditor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;

@Component
public class StartupService {

    private final JdbcTemplate jdbcTemplate;

    private File myFile;
    private XSSFSheet mySheet;
    int imsiRowNum = 0;

    private static final String SCRIPT =
            "select imsi, brand_cd from h_schema.device_mst where imsi in (%s) " +
            "UNION all " +
            "select imsi, brand_cd from k_schema.device_mst where imsi in (%s) " +
            "UNION all " +
            "select imsi, brand_cd from g_schema.device_mst where imsi in (%s)";

    public StartupService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public void start() throws IOException {
        System.out.println("Application START !!!!!!!!!!!!!!");
        XSSFWorkbook myWorkBook = getXssfSheets();

        if(myWorkBook != null){
            mySheet = myWorkBook.getSheet("Data");
            String imsiList = getImsiList(mySheet);

            List<Map<String, Object>> maps = jdbcTemplate.queryForList(String.format(SCRIPT, imsiList, imsiList, imsiList));

            fillXLSX(myWorkBook, maps);
        }

        System.out.println("Application FINISH !!!!!!!!!");
    }

    private void fillXLSX(XSSFWorkbook myWorkBook, List<Map<String, Object>> maps) throws IOException {
        Map<String, String> imsiMap = new HashMap<>();
        maps.forEach(x -> imsiMap.put(x.get("imsi").toString(), x.get("brand_cd").toString()));

        FileOutputStream os = new FileOutputStream(myFile);
        Iterator<Row> rowIterator = mySheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell imsiCell = row.getCell(imsiRowNum);
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
            if(row.getRowNum() == 0){
                for (int i = 0; i <= row.getLastCellNum(); i++){
                    if(row.getCell(i).getCellType() == Cell.CELL_TYPE_STRING &&
                            row.getCell(i).getStringCellValue().equals("IMSI")){
                        imsiRowNum = i;
                        break;
                    }
                }
            }
            Cell cell = row.getCell(imsiRowNum);
            int cellType = cell.getCellType();
            String cellValue = cellType == 1 ?
                    cell.getStringCellValue() : Double.toString(cell.getNumericCellValue());
            result.append("\'").append(cellValue).append("\'");
            if(rowIterator.hasNext())
                result.append(",");
        }
        return result.toString();
    }

    private XSSFWorkbook getXssfSheets() throws IOException {
        File f = new File("./");
        FilenameFilter filter = new FilenameFilter() {
            @Override
            public boolean accept(File f, String name) {
                return name.endsWith(".xlsx");
            }
        };
        File[] files = f.listFiles(filter);
        if(files.length != 1){
            System.out.println("You should delete other .xlsx files");
            return null;
        }
        myFile = new File(files[0].getName());
        FileInputStream fis = new FileInputStream(myFile);
        return new XSSFWorkbook(fis);
    }

}
