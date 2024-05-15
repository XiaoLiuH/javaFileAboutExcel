package utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

public class SelectColByName {

    /**
     * 根据单元格中的值 查找单元格地址
     * @param colName 目标数据
     * @param workbook
     * @param sheetName Sheet页名称
     * */
    public CellAddress selectCol(String colName, Workbook workbook, String sheetName){

        final CellAddress A1 = new CellAddress(0, 0);

        // 通过sheet页名称 查找sheet数据
        Sheet sheet = workbook.getSheet(sheetName);

        // 行计数器 (废弃)
        int rowNumber = 0;

        // 遍历行
        for (Row row : sheet) {

            // 遍历列
            for (Cell cell : row) {

                if (cell.getCellType() == CellType.BLANK || cell.getCellType() == CellType.ERROR){
                    continue;
                }
                // 数据为字符型
                if (cell.getCellType() == CellType.STRING) {
                    String stringCellValue = cell.getStringCellValue();
                    if (stringCellValue.trim().equals(colName.trim())){
                       return cell.getAddress();
                    }
                }
                // 数据为数值类型
                if (cell.getCellType() == CellType.FORMULA) {
                    double numericCellValue = cell.getNumericCellValue();
                    if (numericCellValue == Double.valueOf(colName)){
                        return cell.getAddress();
                    }
                }
                // 数据为布尔值
                if (cell.getCellType() == CellType.FORMULA) {
                    String stringCellValue = cell.getStringCellValue();
                    if (Boolean.getBoolean(stringCellValue) == Boolean.getBoolean(colName)) {
                        return cell.getAddress();
                    }
                }

            }
        }

        return A1;
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

}
