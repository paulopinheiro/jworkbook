/*
 * Creates JWorkbook objects, based on file extension
 * 
 */
package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.security.InvalidParameterException;
import org.apache.commons.io.FilenameUtils;
import resources.MessagesBundle;

/**
 * Creates an JWorkbook object based on file extension.
 * At the present moment only XLS and ODS files are allowed
 * @author paulopinheiro
 */
public class JWorkbookFactory {

    /**
     * Creates a JWorkbook object, based on file extension (XLS for Microsoft Excel or ODS for OpenOffice.org).<br>
     * @param workbookFile The File object that represents the workbook to be written.<br>
     * Throws InvalidParameterException if:<br>
     *      - The param workbookFile is null;<br>
     *      - Parent directory does not exist;<br>
     *      - Parent directory is not writable;<br>
     *      - Filename extension is not XLS or ODS (case insensitive).<br>
     * IMPORTANT: It does not throw an exception if there is already a file with same name. The API client should take care about that.
     * @return A JWorkbook object, based on file extension.
     */
    public static JWorkbook createJWorkbook(File workbookFile) {
        if (workbookFile == null) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.null"));
        }
        if (workbookFile.isDirectory()) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.isDirectory",workbookFile.getName()));
        }
        if (!workbookFile.getParentFile().exists()) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.parentNotFound",workbookFile.getParent()));
        }
        if (!workbookFile.getParentFile().canWrite()) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.cantWrite",workbookFile.getParent()));
        }

        String extension = FilenameUtils.getExtension(workbookFile.getName());
        if (extension.equalsIgnoreCase("xls")) {
            return new JWorkbookXLS(workbookFile);
        } else {
            if (extension.equalsIgnoreCase("ods")) {
                return new JWorkbookODS(workbookFile);
            } else {
                throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.notSupported"));
            }
        }
    }
}
