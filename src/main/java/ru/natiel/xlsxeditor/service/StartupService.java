package ru.natiel.xlsxeditor.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;
import ru.natiel.xlsxeditor.constants.SourceENUM;

import java.io.*;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

@Component
public class StartupService {
    private static final Logger LOGGER =
            Logger.getLogger(StartupService.class.getName());

    private final JdbcTemplate jdbcTemplatePrd;
    private final JdbcTemplate jdbcTemplateStg;

    private File myFile;
    private XSSFSheet mySheet;
    int imsiRowNum = 0;

    private static final String SCRIPT =
            "select imsi, brand_cd from h_schema.device_mst where imsi in (%s) " +
            "UNION all " +
            "select imsi, brand_cd from k_schema.device_mst where imsi in (%s) " +
            "UNION all " +
            "select imsi, brand_cd from g_schema.device_mst where imsi in (%s)";

    public StartupService(@Qualifier("jdbcTemplatePrd") JdbcTemplate jdbcTemplatePrd,
                          @Qualifier("jdbcTemplateStg") JdbcTemplate jdbcTemplateStg) {
        this.jdbcTemplatePrd = jdbcTemplatePrd;
        this.jdbcTemplateStg = jdbcTemplateStg;
    }

    public void start(String fileName) throws IOException {
        XSSFWorkbook myWorkBook = getXssfSheets(fileName);

        if(myWorkBook != null){
            mySheet = myWorkBook.getSheet("Data");
            mySheet.setDefaultColumnWidth(10);
            fillPage(myWorkBook);
        }
    }

    private XSSFWorkbook getXssfSheets(String fileName) throws IOException {
        if(fileName.isEmpty()){
            LOGGER.log(Level.WARNING, "Need to specify file name");
            return null;
        }
        myFile = new File(fileName);
        FileInputStream fis = new FileInputStream(myFile);
        return new XSSFWorkbook(fis);
    }

    private void fillPage(XSSFWorkbook myWorkBook) throws IOException {
        Map<String, String> imsiMapStg = new HashMap<>();
        Map<String, String> imsiMapPrd = new HashMap<>();
        String imsiList = getImsiList(mySheet);

        List<Map<String, Object>> mapStg = jdbcTemplateStg.queryForList(String.format(SCRIPT, imsiList, imsiList, imsiList));
        List<Map<String, Object>> mapPrd = jdbcTemplatePrd.queryForList(String.format(SCRIPT, imsiList, imsiList, imsiList));

        mapStg.forEach(x -> imsiMapStg.put(x.get("imsi").toString(), x.get("brand_cd").toString()));
        mapPrd.forEach(x -> imsiMapPrd.put(x.get("imsi").toString(), x.get("brand_cd").toString()));

        File output = new File("output_" + myFile.getName());
        FileOutputStream os = new FileOutputStream(output);
        fillCells(imsiMapStg, imsiMapPrd, getInvoiceList(), myWorkBook.createCellStyle());
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

    private Map<String, String> getInvoiceList() throws IOException {
        Map<String, String> result = new HashMap<>();

        File dir = new File("./");
        FilenameFilter filter = (f, name) -> name.toLowerCase().endsWith(".xlsx") &&
                                            name.toLowerCase().startsWith("ru ccs sim");
        File[] files = dir.listFiles(filter);
        if(files!= null && files.length != 1){
            LOGGER.log(Level.WARNING, "You should delete other .xlsx files");
            return null;
        }
        File invoiceFile = files[0];
        FileInputStream fis = new FileInputStream(invoiceFile);
//        FileInputStream fis = new FileInputStream("sim.xlsx");
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
        XSSFSheet invoiceSheet = myWorkBook.getSheet("VIN LIST");

        int imsiIndex = 0;
        int invoiceIndex = 0;

        Iterator<Row> rowIterator = invoiceSheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if(row.getRowNum() == 0){
                for (int i = 0; i <= row.getLastCellNum(); i++){
                    Cell cell = row.getCell(i);
                    if(cell != null && cell.getCellType() == 1 && cell.getStringCellValue().equals("IMSI")){
                        imsiIndex = i;
                    }
                    if(cell != null && cell.getCellType() == 1 && cell.getStringCellValue().equals("INVOICE TO")){
                        invoiceIndex = i;
                    }
                }
                continue;
            }
            Cell imsiCell = row.getCell(imsiIndex);
            Cell invoiceCell = row.getCell(invoiceIndex);
            if(imsiCell != null && invoiceCell != null){
                String imsiValue = imsiCell.getCellType() == 1 ?
                        imsiCell.getStringCellValue() : Double.toString(imsiCell.getNumericCellValue());
                String invoiceValue = invoiceCell.getCellType() == 1 ?
                        invoiceCell.getStringCellValue() : Double.toString(invoiceCell.getNumericCellValue());
                result.put(imsiValue, invoiceValue);
            }
        }
        return result;
    }

    private void fillCells(Map<String, String> mapStg, Map<String, String> mapPrd, Map<String, String> mapInvoice, XSSFCellStyle style) {
        Iterator<Row> rowIterator = mySheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell imsiCell = row.getCell(imsiRowNum);
            Cell brandCellStg = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            Cell brandCellPrd = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            Cell invoiceCell = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            String valueStg = row.getRowNum() == 0?
                    SourceENUM.STG.toString() : mapStg.get(imsiCell.getStringCellValue());
            String valuePrd = row.getRowNum() == 0?
                    SourceENUM.PRD.toString() : mapPrd.get(imsiCell.getStringCellValue());
            String invoiceValue = row.getRowNum() == 0?
                    SourceENUM.INVOICE.toString() : mapInvoice.get(imsiCell.getStringCellValue());

            if (valueStg!= null && valuePrd!= null && !valueStg.equals(valuePrd)){
                style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                brandCellStg.setCellStyle(style);
                brandCellPrd.setCellStyle(style);
            }
            brandCellStg.setCellValue(valueStg);
            brandCellPrd.setCellValue(valuePrd);
            invoiceCell.setCellValue(invoiceValue);
        }
    }

}
