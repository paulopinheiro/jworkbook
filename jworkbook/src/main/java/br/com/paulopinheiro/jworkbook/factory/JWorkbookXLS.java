/*
 * JWorkbook for Microsoft Excel files
 */
package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.security.InvalidParameterException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.hssf.record.cf.BorderFormatting;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;

/**
 * JWorkbook for Microsoft Excel files
 *
 * @author usuario
 */
class JWorkbookXLS implements JWorkbook {

    private FileOutputStream fos;
    private Workbook workbook;
    private Sheet currentSheet;
    private Row currentRow;
    private CellStyle plainCellStyle;
    private List<Row> tittleTotalRows;

    protected JWorkbookXLS(File workbookFile) {
        try {
            setFos(new FileOutputStream(workbookFile));
        } catch (FileNotFoundException ex) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.parentNotFound", workbookFile.getParent()));
        }
        setWorkbook(new HSSFWorkbook());
        setPlainCellStyle(getWorkbook().getCellStyleAt((short) 0));
    }

    @Override
    public void addSheet(String sheetName) {
        finishCurrentSheet();
        if (!(getWorkbook().getSheet(sheetName) == null)) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("sheet.alreadyExists"));
        }
        setCurrentSheet(getWorkbook().createSheet(WorkbookUtil.createSafeSheetName(sheetName)));
        setCurrentRow(null);
    }

    @Override
    public void addSheet(String sheetName, String[] header, String[] footer) {
        this.addSheet(sheetName);
        if (header != null) {
            this.setSheetHeader(header);
        }
        if (footer != null) {
            this.setSheetFooter(footer);
        }
    }

    private void finishCurrentSheet() {
        if ((getCurrentSheet() != null) && (getCurrentRow() != null)) {
            formatRows();
            formatCols();
            setCurrentRow(null);
        }
    }

    private void formatCols() {
        int maxCol = 0;

        /*
         Look for the last col used for any row (maxCol)
         */
        for (Row row : getCurrentSheet()) {
            if (row.getLastCellNum() > maxCol) {
                maxCol = row.getLastCellNum();
            }
        }

        /*
         Starts an array of favourite alignments for each col
         The favourite alignment is the most frequent one for that col
         */
        short[] favAlign = this.getFavouriteAlignments(maxCol);

        for (int i = 0; i < maxCol; i++) {
            getCurrentSheet().autoSizeColumn(i);
            for (Row row : getCurrentSheet()) {
                if (row.getCell(i) != null) {
                    row.getCell(i).getCellStyle().setAlignment(favAlign[i]);
                }
            }
        }
    }

    private short[] getFavouriteAlignments(int numCols) {
        short[] resposta = allCenterlAligned(numCols);

        //Only detailRows will be analyzed
        //Tittle/total rows does not count
        List<Row> detailRows = this.getDetailRows();

        // No detail rows: all general aligned
        if (detailRows.isEmpty()) {
            return resposta;
        }

        for (int i = 0; i < numCols; i++) {
            List<Short> cellAlignments = new ArrayList<Short>();

            for (Row row : detailRows) {
                if (getCellAlignment(row,i)!=null)
                    cellAlignments.add(getCellAlignment(row, i));
            }
            if (!cellAlignments.isEmpty())
                resposta[i] = favouriteAlignment(cellAlignments);
        }

        return resposta;
    }

    private static short[] allCenterlAligned(int numCols) {
        short[] resposta = new short[numCols];
        for (int i = 0; i < numCols; i++) {
            resposta[i] = CellStyle.ALIGN_CENTER;
        }
        return resposta;
    }

    private static Short getCellAlignment(Row row, int index) {
        return row.getCell(index) != null ? row.getCell(index).getCellStyle().getAlignment() : null;
    }

    private List<Row> getDetailRows() {
        List<Row> detailRows = new ArrayList<Row>();
        for (Row row : getCurrentSheet()) {
            if (!this.getTittleTotalRows().contains(row)) {
                detailRows.add(row);
            }
        }
        return detailRows;
    }

    private static short favouriteAlignment(List<Short> colAlignments) {
        Map cardMap = (Map<Short, Integer>) CollectionUtils.getCardinalityMap(colAlignments);

        Comparator comp = new Comparator<Entry<Short, Integer>>() {
            @Override
            public int compare(Entry<Short, Integer> o1, Entry<Short, Integer> o2) {
                return o1.getValue() > o2.getValue() ? 1 : -1;
            }
        };

        Entry<Short, Integer> entry = (Entry<Short, Integer>) Collections.max(cardMap.entrySet(), comp);

        return entry.getKey();
    }

    private void formatRows() {
        for (Cell c : getCurrentRow()) {
            this.lastRowCellStyle(c.getCellStyle());
        }
    }

    private void setSheetHeader(String[] header) {
        Header sheetHeader = this.getCurrentSheet().getHeader();
        try {
            sheetHeader.setLeft(header[0]);
            sheetHeader.setCenter(header[1]);
            sheetHeader.setRight(header[2]);
        } catch (ArrayIndexOutOfBoundsException ex) {
        }

    }

    private void setSheetFooter(String[] footer) {
        Footer sheetFooter = this.getCurrentSheet().getFooter();
        try {
            sheetFooter.setLeft(footer[0]);
            sheetFooter.setCenter(footer[1]);
            sheetFooter.setRight(footer[2]);
        } catch (ArrayIndexOutOfBoundsException ex) {
        }
    }

    @Override
    public void addRow(Object[] cells) {
        this.addRow(cells, false);
    }

    @Override
    public void addRow(Object[] cells, boolean tittleTotal) {
        setCurrentRow(getCurrentSheet().createRow(getNewRowIndex()));

        if (cells != null) {
            for (int i = 0; i < cells.length; i++) {
                if (cells[i] == null) {
                    addCell("", i, tittleTotal);
                } else {
                    addCell(cells[i], i, tittleTotal);
                }
            }
        }
        if (tittleTotal) {
            this.addTittleTotalRows(this.getCurrentRow());
        }
    }

    /* Add a cell and set it with the appropriate type and alignment */
    private void addCell(Object o, int index, boolean tittleTotal) {
        Cell cell = getCurrentRow().createCell(index);
        short alignment;

        if (tittleTotal) {
            this.tittleTotalCellStyle(cell);
        } else {
            this.detailCellStyle(cell);
        }

        if (o instanceof Date) {
            cell.setCellValue((Date) o);
            this.dateCellStyle(cell.getCellStyle());
            alignment = CellStyle.ALIGN_RIGHT;
        } else {
            if (o instanceof Calendar) {
                cell.setCellValue((Calendar) o);
                this.calendarCellStyle(cell.getCellStyle());
                alignment = CellStyle.ALIGN_RIGHT;
            } else {
                if (o instanceof String) {
                    cell.setCellValue((String) o);
                    alignment = CellStyle.ALIGN_LEFT;
                } else {
                    if (o instanceof Boolean) {
                        cell.setCellValue((Boolean) o);
                        alignment = CellStyle.ALIGN_CENTER;
                    } else {
                        if (o instanceof Double) {
                            cell.setCellValue((Double) o);
                            alignment = CellStyle.ALIGN_RIGHT;
                        } else {
                            String arg = o.toString();

                            if (NumberUtils.isNumber(arg)) {
                                cell.setCellValue(Double.parseDouble(arg));
                                alignment = CellStyle.ALIGN_RIGHT;
                            } else {
                                cell.setCellValue(arg);
                                alignment = CellStyle.ALIGN_LEFT;
                            }
                        }
                    }
                }
            }
        }
        CellUtil.setAlignment(cell, getWorkbook(), alignment);

    }

    private void dateCellStyle(CellStyle style) {
        CreationHelper ch = getWorkbook().getCreationHelper();
        style.setDataFormat(ch.createDataFormat().getFormat("dd/MM/yyyy"));

    }

    private void calendarCellStyle(CellStyle style) {
        CreationHelper ch = getWorkbook().getCreationHelper();
        style.setDataFormat(ch.createDataFormat().getFormat("dd/MM/yyyy HH:mm"));
    }

    private void tittleTotalCellStyle(Cell cell) {
        cell.getCellStyle().cloneStyleFrom(getPlainCellStyle());
        CellStyle style = cell.getCellStyle();
        style.cloneStyleFrom(getPlainCellStyle());
        style.setBorderTop(BorderFormatting.BORDER_DOUBLE);
        style.setBorderBottom(BorderFormatting.BORDER_DOUBLE);

        Font font = getWorkbook().createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontName("Courier New");
        style.setFont(font);
        cell.setCellStyle(style);
    }

    private void detailCellStyle(Cell cell) {
        cell.getCellStyle().cloneStyleFrom(getPlainCellStyle());
        CellStyle style = cell.getCellStyle();

        style.setBorderBottom(BorderFormatting.BORDER_NONE);

        Font font = getWorkbook().createFont();
        font.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        font.setFontName("Courier New");
        style.setFont(font);
        cell.setCellStyle(style);
    }

    private void lastRowCellStyle(CellStyle style) {
        style.setBorderBottom(BorderFormatting.BORDER_DOUBLE);
    }

    /* Write to disk */
    @Override
    public void write() throws IOException {
        finishCurrentSheet();
        getWorkbook().write(getFos());
        getFos().close();
    }

    private FileOutputStream getFos() {
        return fos;
    }

    private void setFos(FileOutputStream fos) {
        finishCurrentSheet();
        this.fos = fos;
    }

    /**
     * @return the workbook
     */
    private Workbook getWorkbook() {
        return workbook;
    }

    /**
     * @param workbook the workbook to set
     */
    private void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * @return the currentSheet
     */
    private Sheet getCurrentSheet() {
        return currentSheet;
    }

    /**
     * @param currentSheet the currentSheet to set
     */
    private void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    /**
     * @return the currentRow
     */
    private Row getCurrentRow() {
        return currentRow;
    }

    private int getNewRowIndex() {
        if (getCurrentRow() == null) {
            return 0;
        }
        return this.getCurrentRow().getRowNum() + 1;
    }

    /**
     * @param currentRow the currentRow to set
     */
    private void setCurrentRow(Row currentRow) {
        this.currentRow = currentRow;
    }

    /**
     * @return the plainCellStyle
     */
    public CellStyle getPlainCellStyle() {
        return plainCellStyle;
    }

    /**
     * @param plainCellStyle the plainCellStyle to set
     */
    public void setPlainCellStyle(CellStyle plainCellStyle) {
        this.plainCellStyle = plainCellStyle;
    }

    /**
     * @return the tittleTotalRows
     */
    private List<Row> getTittleTotalRows() {
        if (this.tittleTotalRows == null) {
            this.tittleTotalRows = new ArrayList<Row>();
        }
        return tittleTotalRows;
    }

    /**
     * @param tittleTotalRows the tittleTotalRows to set
     */
    private void addTittleTotalRows(Row tittleTotalRow) {
        this.getTittleTotalRows().add(tittleTotalRow);
    }
}
