/**
 * JWorkbook for OpenOffice.org Calc files
 */

package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.io.IOException;
import java.security.InvalidParameterException;
import java.util.ArrayList;
import java.util.List;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.dom.element.office.OfficeSpreadsheetElement;
import org.odftoolkit.odfdom.pkg.OdfFileDom;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

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
            this.setWorkbookContentElement(this.getOutputDocument().getContentRoot());
            this.setContentDom(this.getOutputDocument().getContentDom());
            this.setStylesDom(this.getOutputDocument().getStylesDom());
            this.setMetaDom(this.getOutputDocument().getMetaDom());
            this.cleanOutputDocument();
        } catch (Exception ex) {
            throw new RuntimeException(MessagesBundle.getExceptionMessage("ods.creationError",ex.getMessage()));
        }        
    }

    // By default ODFDOM creates a new workbook with one spreadsheet
    // This method is to clean it out
    // Credits to http://langintro.com/odfdom_tutorials/create_odt.html
    private void cleanOutputDocument() {
        Node childNode = this.getWorkbookContentElement().getFirstChild();

        while (childNode != null) {
            this.getWorkbookContentElement().removeChild(childNode);
            childNode = this.getWorkbookContentElement().getFirstChild();
        }
    }

    @Override
    public void addSheet(String sheetName) {
        if (this.alreadyExists(sheetName)) throw new InvalidParameterException(MessagesBundle.getExceptionMessage("sheet.alreadyExists"));
        finishCurrentSheet();
        this.setCurrentSheet(OdfTable.newTable(this.getOutputDocument()));
        this.getCurrentSheet().setTableName(sheetName);
    }

    private void finishCurrentSheet() {
        //use stylesDom to apply tittle and total row formats
    }

    @Override
    public void addSheet(String sheetName, String[] header, String[] footer) {
        //use stylesDom to add header and footer
    }

    private boolean alreadyExists(String sheetName) {
        List<Node> sheetList = this.getSheetList();
        for (Node sheet : sheetList) {
            if (getSheetName(sheet).trim().equals(sheetName.trim())) return true;
        }
        return false;
    }

    private List<Node> getSheetList() {
        List<Node> sheetList = new ArrayList<>();
        Node spreadSheetNode = this.getSpreadsheetNode();
        if (spreadSheetNode!=null) {
            NodeList spreadSheetChildren = spreadSheetNode.getChildNodes();
            for (int i=0;i<spreadSheetChildren.getLength();i++) {
                Node spreadSheetChild = spreadSheetChildren.item(i);
                if (spreadSheetChild.getNodeName().equals("table:table")) {
                    sheetList.add(spreadSheetChild);
                }
            }
        }
        return sheetList;
    }

    private static String getSheetName(Node tableNode) {
        NamedNodeMap map = tableNode.getAttributes();
        for (int i=0;i<map.getLength();i++) {
            Node attribute = map.item(i);
            if (attribute.getNodeName().equals("table:name")) return attribute.getNodeValue();
        }
        return null;
    }

    private Node getSpreadsheetNode() {
        NodeList rootChildren = getContentDom().getChildNodes();
        for (int i=0;i<rootChildren.getLength();i++) {
            Node rootChild = rootChildren.item(i);
            if (rootChild.getNodeName().equals("office:body")) {
                NodeList bodyChildren = rootChild.getChildNodes();
                for (int j=0;j<bodyChildren.getLength();j++) {
                    Node bodyChild = bodyChildren.item(j);
                    if (bodyChild.getNodeName().equals("office:spreadsheet")) return bodyChild;
                }
            }
        }
        return null;
    }

    @Override
    public void addRow(Object[] cells) {
        this.addRow(cells, false);
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
