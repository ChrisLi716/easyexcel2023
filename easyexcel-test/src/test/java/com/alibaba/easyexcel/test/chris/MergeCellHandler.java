package com.alibaba.easyexcel.test.chris;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

@Data
@Accessors(chain = true)
public class MergeCellHandler implements CellWriteHandler {

    // 数组存放这一行需要合并那几列  [0,1,2] 在这mergeRowIndex行中合并 0 、1、2列
    private int[] mergeColumnIndex;

    // 不要合并的列
    private Integer noMergeRowIndex;


    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 如果当前是表头则不处理
        if (isHead) return;

        // 如果当前是第一行则不处理
        if (relativeRowIndex == 0) return;

        int curRowIndex = cell.getRowIndex();
        int curColIndex = cell.getColumnIndex();
        for (int columnIndex : mergeColumnIndex) {
            if (curColIndex == columnIndex) {
                mergeWithPrevRow(writeSheetHolder, cell, curRowIndex, curColIndex);
                break;
            }
        }
    }


    /**
     * 当前单元格向上合并
     * .
     *
     * @param writeSheetHolder writeSheetHolder
     * @param cell             当前单元格
     * @param curRowIndex      当前行
     * @param curColIndex      当前列
     */
    private void mergeWithPrevRow(WriteSheetHolder writeSheetHolder, Cell cell, int curRowIndex, int curColIndex) {
        Object curData = cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : cell.getNumericCellValue();
        Row preRow = cell.getSheet().getRow(curRowIndex - 1);
        if (preRow == null) {
            // 当获取不到上一行数据时，使用缓存sheet中数据
            preRow = writeSheetHolder.getCachedSheet().getRow(curRowIndex - 1);
        }
        Cell preCell = preRow.getCell(curColIndex);
        Object preData = preCell.getCellType() == CellType.STRING ? preCell.getStringCellValue() : preCell.getNumericCellValue();
        // 不需要合并的列直接跳出
        if ((noMergeRowIndex != null) && noMergeRowIndex == (curRowIndex - 1)) {
            return;
        }
        // 将当前单元格数据与上一个单元格数据比较
        boolean dataBool = preData.equals(curData);

        // 此处需要注意：所以获取每一行第一列数据和上一行第一列数据进行比较，如果相等合并
        boolean equals = cell.getRow().getCell(0).getStringCellValue().equals(preRow.getCell(0).getStringCellValue());
        if (dataBool && equals) {
            Sheet sheet = writeSheetHolder.getSheet();
            List<CellRangeAddress> mergeRegions = sheet.getMergedRegions();
            boolean isMerged = false;
            for (int i = 0; i < mergeRegions.size() && !isMerged; i++) {
                CellRangeAddress cellRangeAddr = mergeRegions.get(i);
                // 若上一个单元格已经被合并，则先移出原有的合并单元，再重新添加合并单元
                if (cellRangeAddr.isInRange(curRowIndex - 1, curColIndex)) {
                    sheet.removeMergedRegion(i);
                    cellRangeAddr.setLastRow(curRowIndex);
                    sheet.addMergedRegion(cellRangeAddr);
                    isMerged = true;
                }
            }
            // 若上一个单元格未被合并，则新增合并单元
            if (!isMerged) {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(curRowIndex - 1, curRowIndex, curColIndex, curColIndex);
                sheet.addMergedRegion(cellRangeAddress);
            }
        }
    }

}
