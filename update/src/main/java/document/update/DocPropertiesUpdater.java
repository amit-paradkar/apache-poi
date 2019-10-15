package document.update;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;

public class DocPropertiesUpdater {
	
	Map<String, String> customPropertiesMap = new HashMap<String, String>();
	XWPFDocument document;
	String FileIn, FileOut;
	boolean documentUpdated;
	
	public DocPropertiesUpdater(String fileIn, String fileOut) {
		this.FileIn = fileIn;
		this.FileOut = fileOut;
	};
	
	public boolean getDocumentUpdated() {
		return this.documentUpdated;
	}
	public XWPFDocument updateFields() {
	
		openDocument(this.FileIn);
		
		if(this.document != null) {
			populateCustomPropertiesMap();
			
			processDocument();
			
			if(!this.documentUpdated) {
				
			}
			
			closeDocument();
		}
		
		return this.document;
	}

	private void closeDocument() {
		FileOutputStream out = null;
		
		try {
			out = new FileOutputStream(this.FileOut);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			this.document.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
		try {
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void processDocument() {
		StringBuffer valueToSet = new StringBuffer();
		valueToSet.setLength(0);
		
		for(XWPFParagraph para: this.document.getParagraphs()) {
			boolean endFound = false;
			
			for(XWPFRun run:para.getRuns()) {
				for(CTFldChar fldChar:run.getCTR().getFldCharList()) {
					if(fldChar.getFldCharType() == STFldCharType.END) {
						endFound = true;
					}
					
					if(endFound == true) {
						CTText text = run.getCTR().addNewInstrText();
						if(valueToSet.length() > 0) {
							text.setStringValue(valueToSet.toString());
							this.documentUpdated = true;
						}
						else {
							text.setStringValue("ProcessingFailed");
						}
						
						fldChar.setFldData(text);
						
						endFound = false;
					}
				}
				
				for(CTText ctText:run.getCTR().getInstrTextList()) {
					String instruction = ctText.getStringValue();
					
					for (String prop: customPropertiesMap.keySet()) {
						if(instruction.contains(prop)) {
							valueToSet.setLength(0);
							valueToSet.append(customPropertiesMap.get(prop));
							ctText.setNil();
						}
					}
				}
			}
		}
		
	}

	private void populateCustomPropertiesMap() {
		POIXMLProperties properties = this.document.getProperties();
		POIXMLProperties.CustomProperties customProperties = properties.getCustomProperties();
		
		if(customProperties != null) {
			List<CTProperty> ctProperties = customProperties.getUnderlyingProperties()
					.getPropertyList();
			
			for(CTProperty ctp: ctProperties) {
				this.customPropertiesMap.put(ctp.getName(), ctp.getLpwstr());
			}
		}
	}

	private void openDocument(String file) {
		try {
			this.document = new XWPFDocument(new FileInputStream(file));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
