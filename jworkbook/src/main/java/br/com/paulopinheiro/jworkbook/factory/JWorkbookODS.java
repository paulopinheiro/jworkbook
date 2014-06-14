/*
 * JWorkbook for OpenOffice.org Calc files
 */

package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.io.IOException;
import java.security.InvalidParameterException;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.dom.element.office.OfficeSpreadsheetElement;
import org.odftoolkit.odfdom.pkg.OdfFileDom;

/**
 *
 * @author Paulo Pinheiro
 */
class JWorkbookODS implements JWorkbook {
    private OdfSpreadsheetDocument outputDocument;
    private OdfFileDom contentDom;
    private OdfFileDom stylesDom;
    private OdfFileDom metaDom;
    private OfficeSpreadsheetElement workbookContentElement;
    private File workbookFile;

    private OdfTable currentSheet;

    protected JWorkbookODS(File workbookFile) {
        this.setWorkbookFile(workbookFile);
        this.initDocumentElements();
    }

    private void initDocumentElements() {
        try {
            this.setOutputDocument(OdfSpreadsheetDocument.newSpreadsheetDocument());
            this.setContentDom(this.getOutputDocument().getContentDom());
            this.setStylesDom(this.getOutputDocument().getStylesDom());
            this.setMetaDom(this.getOutputDocument().getMetaDom());
            this.setWorkbookContentElement(this.getOutputDocument().getContentRoot());
        } catch (Exception ex) {
            throw new RuntimeException(MessagesBundle.getExceptionMessage("ods.creationError",ex.getMessage()));
        }        
    }

    @Override
    public void addSheet(String sheetName) {
        if (this.alreadyExists(sheetName)) throw new InvalidParameterException(MessagesBundle.getExceptionMessage("sheet.alreadyExists"));
        this.finishCurrentSheet();
        this.setCurrentSheet(OdfTable.newTable(this.getOutputDocument()));
        this.getCurrentSheet().setTableName(sheetName);
    }

    @Override
    public void addSheet(String sheetName, String[] header, String[] footer) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void finishCurrentSheet() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private boolean alreadyExists(String sheetName) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void addRow(Object[] cells) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void addRow(Object[] cells, boolean tittleTotal) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void write() throws IOException {
        try {
            this.getOutputDocument().save(this.getWorkbookFile());
        } catch (Exception ex) {
            throw new IOException(ex);
        }
    }

    /**
     * @return the outputDocument
     */
    private OdfSpreadsheetDocument getOutputDocument() {
        return outputDocument;
    }

    /**
     * @param outputDocument the outputDocument to set
     */
    private void setOutputDocument(OdfSpreadsheetDocument outputDocument) {
        this.outputDocument = outputDocument;
    }

    /**
     * @return the contentDom
     */
    private OdfFileDom getContentDom() {
        return contentDom;
    }

    /**
     * @param contentDom the contentDom to set
     */
    private void setContentDom(OdfFileDom contentDom) {
        this.contentDom = contentDom;
    }

    /**
     * @return the workbookFile
     */
    private File getWorkbookFile() {
        return workbookFile;
    }

    /**
     * @param workbookFile the workbookFile to set
     */
    private void setWorkbookFile(File workbookFile) {
        this.workbookFile = workbookFile;
    }

    /**
     * @return the stylesDom
     */
    private OdfFileDom getStylesDom() {
        return stylesDom;
    }

    /**
     * @param stylesDom the stylesDom to set
     */
    private void setStylesDom(OdfFileDom stylesDom) {
        this.stylesDom = stylesDom;
    }

    /**
     * @return the workbookContentElement
     */
    private OfficeSpreadsheetElement getWorkbookContentElement() {
        return workbookContentElement;
    }

    /**
     * @param workbookContentElement the workbookContentElement to set
     */
    private void setWorkbookContentElement(OfficeSpreadsheetElement workbookContentElement) {
        this.workbookContentElement = workbookContentElement;
    }

    /**
     * @return the currentSheet
     */
    private OdfTable getCurrentSheet() {
        return currentSheet;
    }

    /**
     * @param currentSheet the currentSheet to set
     */
    private void setCurrentSheet(OdfTable currentSheet) {
        this.currentSheet = currentSheet;
    }

    /**
     * @return the metaDom
     */
    private OdfFileDom getMetaDom() {
        return metaDom;
    }

    /**
     * @param metaDom the metaDom to set
     */
    private void setMetaDom(OdfFileDom metaDom) {
        this.metaDom = metaDom;
    }
}
