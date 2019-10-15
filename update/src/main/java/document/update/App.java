package document.update;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Program to apply custom property values on docx file.
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	String fileIn = "C:\\temp\\DocumentTemplateSimpleField.docx";
    	String fileOut = "C:\\temp\\DocumentTemplateOut.docx";
    			
    	DocPropertiesUpdater updater = new DocPropertiesUpdater(fileIn,fileOut);
        
        XWPFDocument updatedDocument = updater.updateFields();
        
        if(updater.documentUpdated) {
        	System.out.println("Successfully applied Document Properties in " + fileOut);
        }
        else {
        	System.out.println("Failed to apply Document Properties in " + fileOut);
        }
    }
}
