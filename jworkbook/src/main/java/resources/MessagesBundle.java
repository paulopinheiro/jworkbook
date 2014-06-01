/*
 * Helper class to return i18n messages
 * Adapted from Blog of Daniel Murygin
 * http://murygin.wordpress.com/2010/04/23/parameter-substitution-in-resource-bundles/
 */

package resources;

import java.text.MessageFormat;
import java.util.Locale;
import java.util.MissingResourceException;
import java.util.ResourceBundle;

/**
 * Helper class to return i18n messages
 * @author Paulo Pinheiro
 */
public class MessagesBundle {
    private static final String EXCEPTION_MESSAGES="Exception_Messages";
    private static final ResourceBundle ExceptionMessageBundle = ResourceBundle.getBundle(EXCEPTION_MESSAGES, Locale.getDefault());

    public static String getExceptionMessage(String key) {
        try {
            return ExceptionMessageBundle.getString(key);
        } catch (MissingResourceException e) {
            return '!' + key + '!';
        }
    }

    public static String getExceptionMessage(String key, Object... params  ) {
        try {
            return MessageFormat.format(ExceptionMessageBundle.getString(key), params);
        } catch (MissingResourceException e) {
            return '!' + key + '!';
        }
    }
}
