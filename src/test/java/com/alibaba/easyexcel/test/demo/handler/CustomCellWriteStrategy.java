package com.alibaba.easyexcel.test.demo.handler;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.AbstractVerticalCellStyleStrategy;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class CustomCellWriteStrategy extends AbstractVerticalCellStyleStrategy {

    @Override
    protected WriteCellStyle contentCellStyle(CellWriteHandlerContext context) {
        WriteCellStyle writeCellStyle = new WriteCellStyle();
        WriteFont writeFont = new WriteFont();
        writeCellStyle.setWriteFont(writeFont);
        writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        writeFont.setBold(true);
        if (context.getColumnIndex() == 0) {
            writeCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            writeFont.setColor(IndexedColors.DARK_YELLOW.getIndex());
        }
        if (context.getColumnIndex() == 1) {
            writeCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            writeFont.setColor(IndexedColors.DARK_RED.getIndex());
        }
        if (context.getColumnIndex() == 2) {
            writeCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
            writeFont.setColor(IndexedColors.DARK_GREEN.getIndex());
        }
        if (context.getColumnIndex() == 3) {
            writeCellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
            writeFont.setColor(IndexedColors.DARK_BLUE.getIndex());
        }
        if (context.getColumnIndex() == 4) {
            writeCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            writeFont.setColor(IndexedColors.DARK_YELLOW.getIndex());
        }
        if (context.getColumnIndex() == 5) {
            writeCellStyle.setFillForegroundColor(IndexedColors.TEAL.getIndex());
            writeFont.setColor(IndexedColors.DARK_TEAL.getIndex());
        }
        return writeCellStyle;
    }

    @Override
    protected WriteCellStyle headCellStyle(Head head) {
        WriteCellStyle writeCellStyle = new WriteCellStyle();
        writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        if (head.getColumnIndex() == 0) {
            writeCellStyle.setFillForegroundColor(IndexedColors.PINK.getIndex());
        } else {
            writeCellStyle.setFillForegroundColor(IndexedColors.PINK.getIndex());
        }
        return writeCellStyle;
    }

}
