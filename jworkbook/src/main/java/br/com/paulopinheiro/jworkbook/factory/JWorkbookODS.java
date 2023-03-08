/*
 * JWorkbook for OpenOffice.org Calc files
 */

package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.io.IOException;

/**
 *
 * @author usuario
 */
class JWorkbookODS implements JWorkbook {
    protected JWorkbookODS(File workbookFile) {
        
    }

    @Override
    public void write() throws IOException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
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
    
}
