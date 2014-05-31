/*
 * Creates JWorkbook objects, based on file extension
 * 
 */
package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.security.InvalidParameterException;
import java.util.Locale;
import java.util.ResourceBundle;
import org.apache.commons.io.FilenameUtils;

/**
 * Creates an JWorkbook object based on file extension
 * At the present moment only XLS and ODS files are allowed
 * @author paulopinheiro
 */
public class JWorkbookFactory {
    private static final ResourceBundle exceptionMessages = ResourceBundle.getBundle("Exception_Messages", Locale.getDefault());

    /**
     * 
     * @param workbookFile The File object that represents the workbook to be written
     * @return A JWorkbook object, based on file extension
     * At the present moment only XLS and ODS files allowed
     */
    public static JWorkbook createJWorkbook(File workbookFile) {
        if (workbookFile == null) {
            throw new InvalidParameterException(exceptionMessages.getString("file.null"));
        }
        if (workbookFile.isDirectory()) {
            throw new InvalidParameterException("file.isDirectory");
        }
        if (!workbookFile.getParentFile().canWrite()) {
            throw new InvalidParameterException("file.cantWrite");
        }
        String extension = FilenameUtils.getExtension(workbookFile.getName());
        if (!extension.equalsIgnoreCase("xls")) {
            return new JWorkbookXLS(workbookFile);
        } else {
            if (!extension.equalsIgnoreCase("ods")) {
                return new JWorkbookODS(workbookFile);
            } else {
                throw new InvalidParameterException("file.notXLSnorODS");
            }
        }
    }
}
