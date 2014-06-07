/*
 * JWorkbook for OpenOffice.org Calc files
 */

package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.io.IOException;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.pkg.OdfFileDom;

/**
 *
 * @author Paulo Pinheiro
 */
class JWorkbookODS implements JWorkbook {
    private OdfSpreadsheetDocument document;
    private OdfFileDom dom;
    private File workbookFile;
    
    protected JWorkbookODS(File workbookFile) {
        this.setWorkbookFile(workbookFile);
        try {
            this.setDocument(OdfSpreadsheetDocument.newSpreadsheetDocument());
            this.setDom(this.getDocument().getContentDom());
        } catch (Exception ex) {
            throw new RuntimeException(MessagesBundle.getExceptionMessage("ods.creationError",ex.getMessage()));
        }
    }

    @Override
    public void addSheet(String sheetName) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void addSheet(String sheetName, String[] header, String[] footer) {
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
            this.getDocument().save(this.getWorkbookFile());
        } catch (Exception ex) {
            throw new IOException(ex);
        }
    }

    /**
     * @return the document
     */
    private OdfSpreadsheetDocument getDocument() {
        return document;
    }

    /**
     * @param document the document to set
     */
    private void setDocument(OdfSpreadsheetDocument document) {
        this.document = document;
    }

    /**
     * @return the dom
     */
    private OdfFileDom getDom() {
        return dom;
    }

    /**
     * @param dom the dom to set
     */
    private void setDom(OdfFileDom dom) {
        this.dom = dom;
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
}
