/*
 * JWorkbook for Microsoft Excel files
 */
package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.security.InvalidParameterException;
import java.util.Calendar;
import java.util.Date;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.hssf.record.cf.BorderFormatting;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import resources.MessagesBundle;

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

    protected JWorkbookXLS(File workbookFile) {
        try {
            setFos(new FileOutputStream(workbookFile));
        } catch (FileNotFoundException ex) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.parentNotFound", workbookFile.getParent()));
        }
        setWorkbook(new HSSFWorkbook());
    }

    @Override
    public void addSheet(String sheetName) {
        if ((getCurrentSheet()!=null)&&(getCurrentRow()!=null)) {
            getCurrentRow().setRowStyle(this.getLastRowCellStyle());
        }
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
        if (tittleTotal) getCurrentRow().setRowStyle(this.getTittleTotalCellStyle());
        else getCurrentRow().setRowStyle(this.getDetailCellStyle());

        if (cells != null) {
            if (cells.length > getCurrentRow().getLastCellNum()) {
                throw new InvalidParameterException(MessagesBundle.getExceptionMessage("row.tooMuchCells", cells.length));
            }
            for (int i = 0; i < cells.length; i++) {
                if (cells[i] == null) {
                    addCell("", i);
                } else {
                    addCell(cells[i], i);
                }
            }
        }
    }

    /* Add a cell and set it with the appropriate type */
    private void addCell(Object o, int index) {
        Cell cell = getCurrentRow().createCell(index);

        if (o instanceof Date) {
            cell.setCellValue((Date) o);
            CellUtil.setAlignment(cell, getWorkbook(), CellStyle.ALIGN_RIGHT);
        } else {
            if (o instanceof Calendar) {
                cell.setCellValue((Calendar) o);
                CellUtil.setAlignment(cell, getWorkbook(), CellStyle.ALIGN_RIGHT);
            } else {
                if (o instanceof String) {
                    cell.setCellValue((String) o);
                    CellUtil.setAlignment(cell, getWorkbook(), CellStyle.ALIGN_LEFT);
                } else {
                    if (o instanceof Boolean) {
                        cell.setCellValue((Boolean) o);
                        CellUtil.setAlignment(cell, getWorkbook(), CellStyle.ALIGN_CENTER);
                    } else {
                        if (o instanceof Double) {
                            cell.setCellValue((Double) o);
                            CellUtil.setAlignment(cell, getWorkbook(), CellStyle.ALIGN_RIGHT);
                        } else {
                            String arg = o.toString();
                            if (NumberUtils.isNumber(arg)) {
                                cell.setCellValue(Double.parseDouble(arg));
                                CellUtil.setAlignment(cell, getWorkbook(), CellStyle.ALIGN_RIGHT);
                            } else {
                                cell.setCellValue(arg);
                                CellUtil.setAlignment(cell, getWorkbook(), CellStyle.ALIGN_LEFT);
                            }
                        }
                    }
                }
            }
        }
    }

    private CellStyle getTittleTotalCellStyle() {
        CellStyle cellStyle = getWorkbook().createCellStyle();
        cellStyle.setBorderTop(BorderFormatting.BORDER_DOUBLE);
        cellStyle.setBorderBottom(BorderFormatting.BORDER_DOUBLE);

        Font font = getWorkbook().createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontName("Courier New");
        cellStyle.setFont(font);
        
        return cellStyle;
    }

    private CellStyle getDetailCellStyle() {
        CellStyle cellStyle = getWorkbook().createCellStyle();
        cellStyle.setBorderBottom(BorderFormatting.BORDER_NONE);

        Font font = getWorkbook().createFont();
        font.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        font.setFontName("Courier New");
        cellStyle.setFont(font);

        return cellStyle;
    }

    private CellStyle getLastRowCellStyle() {
        CellStyle cellStyle = getWorkbook().createCellStyle();
        cellStyle.setBorderBottom(BorderFormatting.BORDER_DOUBLE);

        return cellStyle;
    }

    /* Write to disk */
    @Override
    public void write() throws IOException {
        getWorkbook().write(getFos());
    }

    private FileOutputStream getFos() {
        return fos;
    }

    private void setFos(FileOutputStream fos) {
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

}
