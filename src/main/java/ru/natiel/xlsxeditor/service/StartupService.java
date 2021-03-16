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
import ru.natiel.xlsxeditor.dto.ImsiData;

import java.io.*;
import java.text.DecimalFormat;
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
    String valueBrandStg = "";
    String valueBrandPrd = "";
    String valueChannelStg = "";
    String valueChannelPrd = "";
    String valueCostStg = "";
    String valueCostPrd = "";
    String invoiceValue = "";
    String keyValue = "";

    private static final String SCRIPT =
            "select de.imsi, de.brand_cd, de.request_channel_cd, di.cost_center from h_schema.device_mst de \n" +
            "JOIN h_schema.device_ifo di ON de.imei = di.imei\n" +
            "where de.imsi in (%s)\n" +
            "UNION all \n" +
            "select de.imsi, de.brand_cd, de.request_channel_cd, di.cost_center from k_schema.device_mst de \n" +
            "JOIN k_schema.device_ifo di ON de.imei = di.imei\n" +
            "where de.imsi in (%s)\n" +
            "UNION all \n" +
            "select de.imsi, de.brand_cd, de.request_channel_cd, di.cost_center from g_schema.device_mst de \n" +
            "JOIN g_schema.device_ifo di ON de.imei = di.imei\n" +
            "where de.imsi in (%s)";

    public StartupService(@Qualifier("jdbcTemplatePrd") JdbcTemplate jdbcTemplatePrd,
                          @Qualifier("jdbcTemplateStg") JdbcTemplate jdbcTemplateStg) {
        this.jdbcTemplatePrd = jdbcTemplatePrd;
        this.jdbcTemplateStg = jdbcTemplateStg;
    }

    public void start(String fileName) throws IOException {
        XSSFWorkbook myWorkBook = getXssfSheets(fileName);
        LOGGER.log(Level.INFO, fileName);
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
        Map<String, ImsiData> imsiMapStg = new HashMap<>();
        Map<String, ImsiData> imsiMapPrd = new HashMap<>();
        String imsiList = getImsiList(mySheet);

        List<Map<String, Object>> mapStg = jdbcTemplateStg.queryForList(String.format(SCRIPT, imsiList, imsiList, imsiList));
        List<Map<String, Object>> mapPrd = jdbcTemplatePrd.queryForList(String.format(SCRIPT, imsiList, imsiList, imsiList));

        mapStg.forEach(x -> imsiMapStg.put(x.get("imsi").toString(), getFromSQLselect(x)));
        mapPrd.forEach(x -> imsiMapPrd.put(x.get("imsi").toString(), getFromSQLselect(x)));

        File output = new File("output_" + myFile.getName());
        FileOutputStream os = new FileOutputStream(output);
        fillCells(imsiMapStg, imsiMapPrd, getInvoiceList(), myWorkBook.createCellStyle());
        myWorkBook.write(os);
    }

    private String getImsiList(XSSFSheet mySheet) {
        StringBuilder result = new StringBuilder();
        Iterator<Row> rowIterator = mySheet.iterator();
        boolean isFirst = true;
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
            if(cell != null){
                if(isFirst)
                    isFirst = false;
                else
                    result.append(",");
                String cellValue = getString(cell);
                result.append("\'").append(cellValue).append("\'");
            }
        }
        return result.toString();
    }

    private String getString(Cell cell) {
        int cellType = cell.getCellType();
        String cellValue = "";
        if(cellType == 0){
            double numericCellValue = cell.getNumericCellValue();
            DecimalFormat decimalFormat = new DecimalFormat("###0");
            cellValue = decimalFormat.format(numericCellValue);
        } else if(cellType == 1){
            cellValue = cell.getStringCellValue();
        }
        return cellValue;
    }

    private Map<String, String> getInvoiceList() throws IOException {
        Map<String, String> result = new HashMap<>();

        File dir = new File("./");
        FilenameFilter filter = (f, name) -> name.toLowerCase().endsWith(".xlsx") &&
                                            name.toLowerCase().startsWith("ru ccs sim");
        File[] files = dir.listFiles(filter);
        if(files!= null && files.length != 1){
            LOGGER.log(Level.WARNING, "You should delete other .xlsx files");
            return new HashMap<>();
        }
        XSSFSheet invoiceSheet = new XSSFWorkbook(new FileInputStream(files[0]))
                .getSheet("VIN LIST");

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

    private void fillCells(Map<String, ImsiData> mapStg, Map<String, ImsiData> mapPrd, Map<String, String> mapInvoice, XSSFCellStyle style) {
        Iterator<Row> rowIterator = mySheet.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell imsiCell = row.getCell(imsiRowNum);
            Cell brandCellStg = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            Cell brandCellPrd = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            Cell channelCellStg = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            Cell channelCellPrd = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            Cell costCellStg = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            Cell costCellPrd = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            Cell invoiceCell = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
            if(row.getRowNum() == 0){
                valueBrandStg = SourceENUM.STGBRAND.name();
                valueBrandPrd = SourceENUM.PRDBRAND.name();
                invoiceValue = SourceENUM.INVOICE.name();
                valueChannelStg = SourceENUM.STGREQUESTCHANNEL.name();
                valueChannelPrd = SourceENUM.PRDREQUESTCHANNEL.name();
                valueCostStg = SourceENUM.STGCOSTCENTER.name();
                valueCostPrd = SourceENUM.PRDCOSTCENTER.name();
            } else if(imsiCell != null){
                keyValue = getString(imsiCell);
                if(!keyValue.equals("")) {
                    ImsiData imsiDataStg = mapStg.get(keyValue);
                    ImsiData imsiDataPrd = mapPrd.get(keyValue);
                    if (imsiDataStg != null) {
                        valueBrandStg = imsiDataStg.getBrandCd();
                        valueChannelStg = imsiDataStg.getRequestChannelCd();
                        valueCostStg = imsiDataStg.getCostCenter();
                    } else {
                        valueBrandStg = "";
                        valueChannelStg = "";
                        valueCostStg = "";
                    }
                    if (imsiDataPrd != null) {
                        valueBrandPrd = imsiDataPrd.getBrandCd();
                        valueChannelPrd = imsiDataPrd.getRequestChannelCd();
                        valueCostPrd = imsiDataPrd.getCostCenter();
                    } else {
                        valueBrandPrd = "";
                        valueChannelPrd = "";
                        valueCostPrd = "";
                    }
                    invoiceValue = mapInvoice.get(keyValue);
                }
            }

            if (valueBrandStg!= null && valueBrandPrd!= null && !valueBrandStg.equals(valueBrandPrd)){
                style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                brandCellStg.setCellStyle(style);
                brandCellPrd.setCellStyle(style);
            }
            brandCellStg.setCellValue(valueBrandStg);
            brandCellPrd.setCellValue(valueBrandPrd);
            channelCellStg.setCellValue(valueChannelStg);
            channelCellPrd.setCellValue(valueChannelPrd);
            costCellStg.setCellValue(valueCostStg);
            costCellPrd.setCellValue(valueCostPrd);
            invoiceCell.setCellValue(invoiceValue);
        }
    }

    private ImsiData getFromSQLselect(Map<String, Object> map){
        String channel = map.get("request_channel_cd") == null ? "" : map.get("request_channel_cd").toString();
        String cost = map.get("cost_center") == null ? "" : map.get("cost_center").toString();
        String brand = map.get("brand_cd") == null ? "" : map.get("brand_cd").toString();
        return new ImsiData(brand, channel, cost);
    }

}
