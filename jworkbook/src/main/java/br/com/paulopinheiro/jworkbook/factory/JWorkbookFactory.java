/*
 * Creates JWorkbook objects, based on file extension
 */

package br.com.paulopinheiro.jworkbook.factory;

import java.io.File;
import java.security.InvalidParameterException;

/**
 *
 * @author paulopinheiro
 */
public class JWorkbookFactory {
    public static JWorkbook createJWorkbook(File parentDirectory, String fileName) {
        if (!parentDirectory.exists()) throw new InvalidParameterException("Mensagem internacionalizada");
        return null;
    }
}
