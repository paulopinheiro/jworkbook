/*
 * Common methods to JWorbook classes
 */

package br.com.paulopinheiro.jworkbook.factory;

import java.io.IOException;

/**
 * Common methods to JWorbook classes
 * @author paulopinheiro
 */
public interface JWorkbook {
    /**
     * Adds a new sheet to the workbook.<br>
     * Throws InvalidParameterException if there is already a sheet with the same name.
     * @param sheetName The name of the new sheet
     */
    public void addSheet(String sheetName);

    /**
     * Adds a new sheet to workbook, with header and footer string arrays (left, center, right).<br>
     * If a given string array length is lesser than 3 the missing ones will be considered 0 lenght strings.<br>
     * If the length is greater than 3, then only first, second and third elements will be taken.<br>
     * Throws InvalidParameterException if there is already a sheet with the same name.<br>
     * @param sheetName The name of the sheet
     * @param header A String array with text to be put on left, center and right position of the header of the sheet
     * @param footer A String array with text to be put on left, center and right position of the header of the sheet
     */
    public void addSheet(String sheetName, String[] header, String[] footer);

    /**
     * Adds a new row to the current sheet.<br>
     * Throws InvalidParameterException if the array contains more elements than the row can support.<br>
     * @param cells An array containing the values of the cells of the row.
     */
    public void addRow(Object[] cells);

    /**
     * Adds a new row to the current sheet.<br>
     * Throws InvalidParameterException if the array contains more elements than the row can support.<br>
     * @param cells An array containing the values of the cells of the row.
     * @param tittleTotal boolean value, informing if the row is a tittle/total one
     */
    public void addRow(Object[] cells, boolean tittleTotal);

    /**
     * Write workbook file to disk
     * @throws IOException 
     */
    public void write() throws IOException;
}
