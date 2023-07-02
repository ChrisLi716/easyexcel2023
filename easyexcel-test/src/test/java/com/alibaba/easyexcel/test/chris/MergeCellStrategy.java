package com.alibaba.easyexcel.test.chris;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.Objects;

public class MergeCellStrategy implements CellWriteHandler {

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList,
                                 Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 如果当前是表头则不处理
        if (isHead) return;

        // 如果当前是第一行则不处理
        if (relativeRowIndex == 0) return;

        Sheet sheet = cell.getSheet();
        int currentRowInx = cell.getRowIndex();  // 当前行下标
        int preRowInx = currentRowInx - 1;       // 上一行下标
        Row currentRow = cell.getRow();             // 当前行
        Row preRow = sheet.getRow(preRowInx); // 上一行
        if (preRow == null) {
            // 当获取不到上一行数据时，使用缓存sheet中数据
            preRow = writeSheetHolder.getCachedSheet().getRow(preRowInx);
        }
        Cell preCell = preRow.getCell(cell.getColumnIndex()); // 上一个单元格

        Object cellValue = (cell.getCellType() == CellType.STRING) ? cell.getStringCellValue() : cell.getNumericCellValue();
        Object preCellValue = (preCell.getCellType() == CellType.STRING) ? preCell.getStringCellValue() : preCell.getNumericCellValue();

        if (!Objects.equals(cellValue, preCellValue)) return;

        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        boolean isMerged = false;
        for (int i = 0; i < mergedRegions.size(); i++) {
            CellRangeAddress cellRangeAddress = mergedRegions.get(i);

            // 如是上一行单元格已合并，则先删除再添加
            if (cellRangeAddress.isInRange(preRowInx, cell.getColumnIndex())) {
                sheet.removeMergedRegion(i); // 移除已合并单元格
                cellRangeAddress.setLastRow(currentRowInx);// 设置合并单元格
                sheet.addMergedRegion(cellRangeAddress);// 重新添加合并单元格
                isMerged = true;
                break;
            }
        }

        if (!isMerged) {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(preRowInx, currentRowInx, cell.getColumnIndex(), cell.getColumnIndex());
            sheet.addMergedRegion(cellRangeAddress);// 重新添加合并单元格
        }
    }
}
