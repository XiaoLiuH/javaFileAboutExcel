import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import cn.hutool.json.JSONObject;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import page.EnumInfo;
import utils.ExcelCopyUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class ExcelFileTest {

    private static Workbook copyWorkBook;

    private static List<JSONObject> postponeTableValue1 = new ArrayList<>(),
            postponeTableValue2 = new ArrayList<>(),
            postponeTableValue3 = new ArrayList<>();

    public static void main(String[] args){

        postponeTableValue1.clear();
        postponeTableValue2.clear();
        postponeTableValue3.clear();
        try {
//          获取源文件
            File file = new File(EnumInfo.filePath);
//          根据复制文件获取Workbook类
            Workbook workBook = WorkbookFactory.create(file);

            writeLogToConsole("\n*************************************************************************");
            writeLogToConsole("workBookSheetNumber:" + workBook.getNumberOfSheets());
//          根据Sheet名称查找数据
            Sheet modeWBSFileSheet = workBook.getSheet("追加開発_WBS");
            writeLogToConsole("*************************************************************************");
            writeLogToConsole("sheetName:" + modeWBSFileSheet.getSheetName());
            writeLogToConsole("Obtain the corresponding sheet ..........");
            writeLogToConsole("create new file ..........");
            String copyPathString = EnumInfo.copyFilePath + "_" + new Date().getTime() + "_" + "IES_IBM進捗報告.xlsx";
            Files.copy(Path.of(EnumInfo.modelFilePath), Path.of(copyPathString));
            writeLogToConsole("loading ..........");
            DataFormatter formatter = new DataFormatter();
            FileInputStream fis = new FileInputStream(copyPathString);
            copyWorkBook = new XSSFWorkbook(fis);
            Sheet copyWorkBookSheet = copyWorkBook.getSheetAt(0);

//          向表格中插入数值数据
            for (Row row : modeWBSFileSheet) {
                for (Cell cell : row) {
//                  将单元格下标映射成对应单元格名称
                    String columnName = getColumnName(cell.getColumnIndex());
//                  获取数值
                    String text = formatter.formatCellValue(cell);
//                  找到对应列
                    if (columnName.equals("C")) {
//                      计数判断
                        switch (text){
                            case "SAP S/4":
                                    setChildNumber1(
                                            6,
                                            26,
                                            text ,
                                            copyWorkBookSheet ,
                                            row);
                                    setChildNumber2(
                                            6,
                                            34,
                                            text,
                                            copyWorkBookSheet,
                                            row
                                    );
                                    setChildNumber3(
                                            6,
                                            52,
                                            text,
                                            copyWorkBookSheet,
                                            row
                                    );
                                    setChildNumber4(
                                            6,
                                            70,
                                            text,
                                            copyWorkBookSheet,
                                            row
                                    );
                                    break;
                            case "BTP":
                            case "IBP":
                                    setChildNumber1(
                                            6,
                                            27,
                                            "BTP",
                                            copyWorkBookSheet ,
                                            row);
                                    setChildNumber2(
                                            6,
                                            35,
                                            "BTP",
                                            copyWorkBookSheet,
                                            row
                                    );
                                    setChildNumber3(
                                            6,
                                            53,
                                            "BTP",
                                            copyWorkBookSheet,
                                            row
                                    );
                                    setChildNumber4(
                                            6,
                                            71,
                                            "BTP",
                                            copyWorkBookSheet,
                                            row
                                    );
                                    break;
                        }
                    }

                }
            }
//          获取每个遅延的数值
            Double sapY1 = copyWorkBookSheet.getRow(34).getCell(getColumnIndex("Y")).getNumericCellValue();
            Double btpY1 = copyWorkBookSheet.getRow(35).getCell(getColumnIndex("Y")).getNumericCellValue();
            Double sapY2 = copyWorkBookSheet.getRow(52).getCell(getColumnIndex("Y")).getNumericCellValue();
            Double btpY2 = copyWorkBookSheet.getRow(53).getCell(getColumnIndex("Y")).getNumericCellValue();
            Double sapY3 = copyWorkBookSheet.getRow(70).getCell(getColumnIndex("AA")).getNumericCellValue();
            Double btpY3 = copyWorkBookSheet.getRow(71).getCell(getColumnIndex("AA")).getNumericCellValue();

//          标记每个表格的初始行标
            int stateIndex1 = 0, stateIndex2 = 0, stateIndex3 = 0;

//          插入对应数的单元格
            ExcelWriter writer = ExcelUtil.getWriter(new File(copyPathString), modeWBSFileSheet.getSheetName());
            XSSFSheet sheet = (XSSFSheet)copyWorkBookSheet;

//          新增行数记录
            int newRowNum = 0;

            Double inRowsNum1 = sapY1 + btpY1;
            if (inRowsNum1.intValue() != 0) {
                ExcelCopyUtils.insertRow(writer, 44, inRowsNum1.intValue() - 3, sheet, false);
            }
//          表格第一行地址记录
            stateIndex1 = 41;

            newRowNum += inRowsNum1.intValue() - 3;

            Double inRowsNum2 = sapY2 + btpY2;
            if (inRowsNum2.intValue() != 0) {
                ExcelCopyUtils.insertRow(writer, 62 + newRowNum, inRowsNum2.intValue() - 3, sheet, false);
            }

            newRowNum += inRowsNum2.intValue() - 3;
            stateIndex2 = 59 + newRowNum;

            Double inRowsNum3 = sapY3 + btpY3;
            if (inRowsNum3.intValue() != 0) {
                ExcelCopyUtils.insertRow(writer, 80 + newRowNum, inRowsNum3.intValue() - 3, sheet, false);
            }
            stateIndex3 = 77 + newRowNum;

            setTableValue(copyWorkBookSheet, stateIndex1, postponeTableValue1, inRowsNum1.intValue());

            setTableValue(copyWorkBookSheet, stateIndex2, postponeTableValue2, inRowsNum2.intValue());

            setTableValue(copyWorkBookSheet, stateIndex3, postponeTableValue3, inRowsNum3.intValue());

            File copyFile = new File(copyPathString);
            FileOutputStream outFile = null;
            try {

                outFile = new FileOutputStream(copyFile);
                copyWorkBook.write(outFile);
                writeLogToConsole("file write successful.");
                writeLogToConsole("new file path: " + copyFile.toPath());

            } catch (IOException e) {

                writeLogToConsole("Error writing to file: " + e.getMessage());

            } finally {
                try {
                    if (outFile != null) {
                        outFile.close();
                    }
                } catch (IOException e) {
                    writeLogToConsole("Error closing file output stream: " + e.getMessage());
                }
            }

        } catch (FileNotFoundException e) {

            writeLogToConsole("file does not exist");
            e.printStackTrace();

        } catch (IOException e) {

            writeLogToConsole("file conversion exception");
            e.printStackTrace();

        }

    }

    /**
     * 向延期明细的表格中插入数据
     * @param copyWorkBookSheet
     * @param stateIndex2
     * @param postponeTableValue2
     * @param addTableNumber
     * */
    private static void setTableValue(Sheet copyWorkBookSheet,
                                      int stateIndex2,
                                      List<JSONObject> postponeTableValue2,
                                      Integer addTableNumber) {

        int i = 0;

        for (JSONObject jsonObject : postponeTableValue2) {

                Row copyRow = copyWorkBookSheet.getRow(stateIndex2 + i);
                Cell copyCellE = copyRow.getCell(getColumnIndex("E"));
                Cell copyCellJ = copyRow.getCell(getColumnIndex("J"));
                Cell copyCellT = copyRow.getCell(getColumnIndex("T"));
                copyCellE.setCellValue(String.valueOf(jsonObject.get("機能ID")));
                copyCellJ.setCellValue(String.valueOf(jsonObject.get("機能名")));
                copyCellT.setCellValue(String.valueOf(jsonObject.get("担当者")));

                if (i < addTableNumber){
                    i ++;
                }

        }
    }

    /**
     * 根据下标获取列名
     *
     * @param columnIndex 下标
     * */
    private static String getColumnName(int columnIndex) {
        StringBuilder columnName = new StringBuilder();
        while (columnIndex >= 0) {
            columnName.insert(0, (char) ('A' + columnIndex % 26));
            columnIndex = (columnIndex / 26) - 1;
        }
        return columnName.toString();
    }


    /**
     * 根据列名获取列下标
     * @param columnName 列名 例：R
     * */
    private static int getColumnIndex(String columnName) {
        int index = 0;
        int multiplier = 1;

        for (int i = columnName.length() - 1; i >= 0; i--) {
            char c = columnName.charAt(i);
            index += (c - 'A' + 1) * multiplier;
            multiplier *= 26;
        }

        return index - 1; // Adjusting to 0-based index
    }

    /**
     * 根据行数据判断所计算目标数据的数值 (1．設計引継ぎ)
     *
     * @param cellNum 写入目标数据所在列下标
     * @param rowNum 写入目标数据所在行下标
     * @param workBookSheet 目标表格数据
     * @param row 源表格行数据
     *
     * */
    private static void setChildNumber1(int cellNum,
                                       int rowNum,
                                       String childName,
                                       Sheet workBookSheet,
                                       Row row){

        Map<String, Integer> fileValuesMap = new HashMap<>();
        Row copyRow = workBookSheet.getRow(rowNum);
        Cell copyCell = copyRow.getCell(cellNum);
        if (copyCell == null){
            copyRow.createCell(cellNum);
            copyCell = copyRow.getCell(cellNum);
        }
        setCellValue(copyCell, childName + "総数", fileValuesMap);

        Cell rCell = row.getCell(getColumnIndex("R"));
        if (rCell.getCellType() != CellType.BLANK && rCell.getCellType() != CellType.ERROR && rCell.getCellType() != CellType.STRING) {
            copyCell = copyRow.getCell(8);
            setCellValue(copyCell, childName + "着手可", fileValuesMap);
            if (contrastDate(rCell)){
                copyCell = copyRow.getCell(10);
                setCellValue(copyCell, childName + "着手予", fileValuesMap);
            }
        }

        setValueByColName("T", 12, row, copyCell, copyRow, childName + "着手実", fileValuesMap);

        setValueByColName("S", 14, row, copyCell, copyRow, childName + "引継完予", fileValuesMap);

        setValueByColName("U", 16, row, copyCell, copyRow, childName + "引継完実", fileValuesMap);

        setPrereciprocalDataAndDelayNumber("Q", "M", "S", "U", copyRow);

    }

    /**
     * 根据行数据判断所计算目标数据的数值 (2．設計進捗)
     *
     * @param cellNum 写入目标数据所在列下标
     * @param rowNum 写入目标数据所在行下标
     * @param workBookSheet 目标表格数据
     * @param row 源表格行数据
     *
     * */
    private static void setChildNumber2(int cellNum,
                                        int rowNum,
                                        String childName,
                                        Sheet workBookSheet,
                                        Row row ){

        Map<String, Integer> fileValuesMap = new HashMap<>();
        Row copyRow = workBookSheet.getRow(rowNum);
        Cell copyCell = copyRow.getCell(cellNum);
        if (copyCell == null){
            copyRow.createCell(cellNum);
            copyCell = copyRow.getCell(cellNum);
        }
        setCellValue(copyCell, childName + "総数", fileValuesMap);

        Cell uCell = row.getCell(getColumnIndex("U"));
        if (uCell.getCellType() != CellType.BLANK && uCell.getCellType() != CellType.ERROR && uCell.getCellType() != CellType.STRING) {
            copyCell = copyRow.getCell(8);
            setCellValue(copyCell, childName + "着手可", fileValuesMap);
        }

        setValueByColName("W", 10, row, copyCell, copyRow, childName + "着手予", fileValuesMap);

        setValueByColName("Y", 12, row, copyCell, copyRow, childName + "着手実", fileValuesMap);

        setValueByColName("X", 14, row, copyCell, copyRow, childName + "設計完予", fileValuesMap);

        setValueByColName("Z", 16, row, copyCell, copyRow, childName + "設計完実", fileValuesMap);

        setValueByColName("Z", 18, row, copyCell, copyRow, childName + "引継完実", fileValuesMap);

        setValueByColName("AC", 20, row, copyCell, copyRow, childName + "ﾚﾋﾞｭｰ完予", fileValuesMap);

        setValueByColName("AE", 22, row, copyCell, copyRow, childName + "ﾚﾋﾞｭｰ完実", fileValuesMap);

        Cell y = row.getCell(getColumnIndex("Y"));
        Cell z = row.getCell(getColumnIndex("Z"));

        if (contrastDate(y) && !contrastDate(z)){
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("機能ID", row.getCell(getColumnIndex("B")).getStringCellValue());
            jsonObject.put("機能名", row.getCell(getColumnIndex("C")).getStringCellValue());
            jsonObject.put("担当者", row.getCell(getColumnIndex("AA")).getStringCellValue());
            postponeTableValue1.add(jsonObject);
        }

        setPrereciprocalDataAndDelayNumber("Q", "M", "W", "Y", copyRow);

    }

    /**
     * 根据行数据判断所计算目标数据的数值 (3．製造進捗)
     *
     * @param cellNum 写入目标数据所在列下标
     * @param rowNum 写入目标数据所在行下标
     * @param workBookSheet 目标表格数据
     * @param row 源表格行数据
     *
     * */
    private static void setChildNumber3(int cellNum,
                                        int rowNum,
                                        String childName,
                                        Sheet workBookSheet,
                                        Row row ){

        Map<String, Integer> fileValuesMap = new HashMap<>();
        Row copyRow = workBookSheet.getRow(rowNum);
        Cell copyCell = copyRow.getCell(cellNum);

        Cell uCell = row.getCell(getColumnIndex("U"));
        if (uCell.getCellType() != CellType.BLANK && uCell.getCellType() != CellType.STRING) {
            setCellValue(copyCell, childName + "総数", fileValuesMap);
        }

        setValueByColName("Z", 8, row, copyCell, copyRow, childName + "着手可", fileValuesMap);

        setValueByColName("AW", 10, row, copyCell, copyRow, childName + "着手予", fileValuesMap);

        setValueByColName("AY", 12, row, copyCell, copyRow, childName + "着手実", fileValuesMap);

        setValueByColName("AX", 14, row, copyCell, copyRow, childName + "製造完予", fileValuesMap);

        setValueByColName("AZ", 16, row, copyCell, copyRow, childName + "製造完実", fileValuesMap);

        setValueByColName("BB", 18, row, copyCell, copyRow, childName + "ﾚﾋﾞｭｰ完予", fileValuesMap);

        setValueByColName("BD", 20, row, copyCell, copyRow, childName + "ﾚﾋﾞｭｰ完実", fileValuesMap);

        Cell ay = row.getCell(getColumnIndex("AY"));
        Cell az = row.getCell(getColumnIndex("AZ"));

        if (contrastDate(ay) && !contrastDate(az)){
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("機能ID", row.getCell(getColumnIndex("B")).getStringCellValue());
            jsonObject.put("機能名", row.getCell(getColumnIndex("C")).getStringCellValue());
            jsonObject.put("担当者", row.getCell(getColumnIndex("BA")).getStringCellValue());
            postponeTableValue2.add(jsonObject);
        }

        setPrereciprocalDataAndDelayNumber("Q", "M", "W", "Y", copyRow);

    }

    /**
     * 根据行数据判断所计算目标数据的数值 (4．単体テスト進捗)
     *
     * @param cellNum 写入目标数据所在列下标
     * @param rowNum 写入目标数据所在行下标
     * @param workBookSheet 目标表格数据
     * @param row 源表格行数据
     *
     * */
    private static void setChildNumber4(int cellNum,
                                        int rowNum,
                                        String childName,
                                        Sheet workBookSheet,
                                        Row row ){

        Map<String, Integer> fileValuesMap = new HashMap<>();
        Row copyRow = workBookSheet.getRow(rowNum);
        Cell copyCell = copyRow.getCell(cellNum);
        if (copyCell == null){
            copyRow.createCell(cellNum);
            copyCell = copyRow.getCell(cellNum);
        }
        setCellValue(copyCell, childName + "機能数", fileValuesMap);

        setValueByColName("AK", getColumnIndex("I"), row, copyCell, copyRow, childName + "UTD完了予", fileValuesMap);

        setValueByColName("AM", getColumnIndex("K"), row, copyCell, copyRow, childName + "UTD完了実", fileValuesMap);

        setValueByColName("BJ", getColumnIndex("M"), row, copyCell, copyRow, childName + "ﾃｽﾄ着手予", fileValuesMap);

        setValueByColName("BL", getColumnIndex("O"), row, copyCell, copyRow, childName + "ﾃｽﾄ着手実", fileValuesMap);

        setValueByColName("BK", getColumnIndex("Q"), row, copyCell, copyRow, childName + "ﾃｽﾄ完予", fileValuesMap);

        setValueByColName("BW", getColumnIndex("S"), row, copyCell, copyRow, childName + "ﾃｽﾄ完実", fileValuesMap);

        setValueByColName("BP", getColumnIndex("U"), row, copyCell, copyRow, childName + "検収完予", fileValuesMap);

        setValueByColName("BR", getColumnIndex("W"), row, copyCell, copyRow, childName + "検収完実", fileValuesMap);

        Cell bw = row.getCell(getColumnIndex("BW"));
        Cell bl = row.getCell(getColumnIndex("BL"));

        if (contrastDate(bw) && !contrastDate(bl)){
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("機能ID", row.getCell(getColumnIndex("B")).getStringCellValue());
            jsonObject.put("機能名", row.getCell(getColumnIndex("C")).getStringCellValue());
            jsonObject.put("担当者", row.getCell(getColumnIndex("BN")).getStringCellValue());
            postponeTableValue3.add(jsonObject);
        }

        setPrereciprocalDataAndDelayNumber("S", "O", "Y", "AA", copyRow);

    }



    /**
     * 比时间赋值
     * */
    private static void setValueByColName(String colName,
                                          int copyColIndex,
                                          Row row,
                                          Cell copyCell,
                                          Row copyRow,
                                          String childName,
                                          Map<String, Integer> fileValuesMap){

        Cell cell = row.getCell(getColumnIndex(colName));
        if (cell.getCellType() != CellType.BLANK && cell.getCellType() != CellType.ERROR && cell.getCellType() != CellType.STRING) {
            if (contrastDate(cell)){
                copyCell = copyRow.getCell(copyColIndex);
                setCellValue(copyCell, childName, fileValuesMap);
            }
        }

    }

    /**
     * @param cellName1 计算1
     * @param cellName2 计算2
     * @param cellName3 结果1
     * @param cellName4 结果2
     * */
    private static void setPrereciprocalDataAndDelayNumber(String cellName1,
                                                           String cellName2,
                                                           String cellName3,
                                                           String cellName4,
                                                           Row copyRow){

        // 前倒数据
        Cell q = copyRow.getCell(getColumnIndex(cellName1));
        Cell m = copyRow.getCell(getColumnIndex(cellName2));
        Cell s = copyRow.getCell(getColumnIndex(cellName3));
        Cell u = copyRow.getCell(getColumnIndex(cellName4));
        Double qValue = q.getNumericCellValue();
        Double mValue = m.getNumericCellValue();
        Integer qValueInt = qValue.intValue();
        Integer mValueInt = mValue.intValue();
        s.setCellValue(qValueInt <= mValueInt ? 0: qValueInt-mValueInt);

        // 遅延数据
        int i = qValueInt <= mValueInt ? mValueInt - qValueInt : 0;
        u.setCellValue(i);
        if (i > 0){
            CellStyle cellStyle = copyWorkBook.createCellStyle();
            Font titleFont = copyWorkBook.createFont();
            titleFont.setBold(true);
            titleFont.setColor(IndexedColors.RED.getIndex());
            titleFont.setFontName("ＭＳ Ｐゴシック");
            cellStyle.setFont(titleFont);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            u.setCellStyle(cellStyle);
        }

    }

    /**
     * 向对应单元格中写入数据
     *
     * @param copyCell 输入值的列下标
     * @param childName 下标对应标头的名称 例：総数
     * @param fileValuesMap 计数映射
     * */
    private static void setCellValue(Cell copyCell,
                                     String childName,
                                     Map<String, Integer> fileValuesMap){

        Double cellValue = copyCell.getNumericCellValue();
//      获取映射 键值列表
        Set<String> strings = fileValuesMap.keySet();
        Integer cellValueNumber = strings.contains(childName) ?
                fileValuesMap.get(childName) :
                fileValuesMap.put(childName, cellValue.intValue());

        if (cellValueNumber == null) {
            cellValueNumber = fileValuesMap.get(childName);
        }

        fileValuesMap.put(childName, ++cellValueNumber);
        copyCell.setCellValue(fileValuesMap.get(childName));

    }

    /**
     * 对比时间
     * */
    private static Boolean contrastDate(Cell cell){

        // 获取当前系统时间
        Calendar calendar = Calendar.getInstance();
        SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
        String nowFormat = formatter.format(calendar.getTime());
        if(cell.getCellType() == CellType.NUMERIC) {
            Date dateCellValue = cell.getDateCellValue();
            String cellFormat = formatter.format(dateCellValue);
            return Integer.valueOf(nowFormat) > Integer.valueOf(cellFormat);
        }else {
            return false;
        }

    }

    /**
     * 向控制台输出消息
     * */
    private static void writeLogToConsole(String msg){
        System.out.println(String.format("\033[%dm%s\033[0m", 31, msg));
    }
}
