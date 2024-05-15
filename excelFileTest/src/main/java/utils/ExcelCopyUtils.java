package utils;


import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;


import java.util.ArrayList;
import java.util.List;

public class ExcelCopyUtils {

    /**
     *
     * @param writer ExcelWriter
     * @param startRow 插入行的行标,即在哪一行下插入
     * @param rows 插入多少行
     * @param sheet XSSFSheet
     * @param copyValue 新行复制(startRow-1)行的样式,而且在拷贝行的时候可以指定是否需要拷贝值
     */
    public static void insertRow(ExcelWriter writer, int startRow, int rows, XSSFSheet sheet, boolean copyValue) {
        if (sheet.getRow(startRow + 1) == null) {
            // 如果复制最后一行，首先需要创建最后一行的下一行，否则无法插入，Bug 2023/03/20修复
            sheet.createRow(startRow + 1);
        }

        // 先获取原始的合并单元格address集合
        List<CellRangeAddress> originMerged = new ArrayList<>(sheet.getMergedRegions());

        for (int i = originMerged.size() - 1; i >= 0; i--) {
            CellRangeAddress region = originMerged.get(i);
            // 判断移动的行数后重新拆分
            if (region.getFirstRow() > startRow) {
                sheet.removeMergedRegion(i);
            }
        }

        // 移动行
        sheet.shiftRows(startRow, sheet.getLastRowNum(), rows, true, false);
        sheet.createRow(startRow);

        for (CellRangeAddress cellRangeAddress : originMerged) {
            if (cellRangeAddress.getFirstRow() > startRow) {
                int firstRow = cellRangeAddress.getFirstRow() + rows;
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(
                        firstRow,
                        firstRow + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn()
                );
                sheet.addMergedRegion(newCellRangeAddress);
            }
        }

        CellCopyPolicy cellCopyPolicy = new CellCopyPolicy();
        cellCopyPolicy.setCopyCellValue(copyValue);

        for (int i = 0; i < rows; i++) {
            // 复制行
            sheet.copyRows(startRow - 1, startRow - 1, startRow + i, cellCopyPolicy);
        }

        // 刷新并关闭流
        writer.flush();
        writer.close();
    }

}