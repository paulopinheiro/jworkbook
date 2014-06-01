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
 * Creates an JWorkbook object based on file extension
 * At the present moment only XLS and ODS files are allowed
 * @author paulopinheiro
 */
public class JWorkbookFactory {

    /**
     * 
     * @param workbookFile The File object that represents the workbook to be written
     * @return A JWorkbook object, based on file extension
     * At the present moment only XLS and ODS files allowed
     */
    public static JWorkbook createJWorkbook(File workbookFile) {
        if (workbookFile == null) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.null"));
        }
        if (workbookFile.isDirectory()) {
            throw new InvalidParameterException(MessagesBundle.getExceptionMessage("file.isDirectory",workbookFile.getName()));
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
