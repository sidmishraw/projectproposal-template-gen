/**
 * Project: ProjectProposalIEEE
 * Package: template.core.handler
 * File: ProposalHandler.java
 * 
 * @author sidmishraw
 *         Last modified: Oct 6, 2017 6:55:58 PM
 */
package template.core.handler;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.ICell;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * <p>
 * The {@link ProposalHandler} is a SAX parser for the proposal XML.
 * It reads the proposal and generates a word doc to be used as a template for
 * proposals.
 * 
 * <p>
 * I'm using this to document the template so that I don't have to go through
 * the pain all the time and this is going to be better than using Latex or Word
 * xD
 * 
 * <p>
 * I'm going to track 4 events:
 * startDocument, startElement, endElement and endDocument, and characterData
 * 
 * character data event is transmitted when content is encountered
 * 
 * @author sidmishraw
 *
 *         Qualified Name: template.core.handler.ProposalHandler
 *
 */
public class ProposalHandler extends DefaultHandler {
    
    /**
     * for logging
     */
    private static final Logger logger          = LoggerFactory.getLogger(ProposalHandler.class);
    
    /**** Word doc related ****/
    
    private XWPFDocument        wordDoc;
    
    // the name of the doc is obtained from the docName attribute of the
    // `proposal` tag
    private String              docName;
    
    /**
     * This represents the current text body
     */
    private XWPFParagraph       currentTextBody = null;
    
    /**** Word doc related ****/
    
    /*
     * <p>
     * Called at the start of the XML document
     * 
     * @see org.xml.sax.helpers.DefaultHandler#startDocument()
     */
    @Override
    public void startDocument() throws SAXException {
        
        logger.info(String.format("Starting document..."));
        
        // initiate the wordDocument that is going to have all the parsed
        // content
        try {
            
            // this.wordDoc = new XWPFDocument(new
            // FileInputStream(Paths.get("template.docx").toFile()));
            this.wordDoc = new XWPFDocument();
            
            // delete old runs
            // this.wordDoc.getParagraphs().forEach(p -> p.removeRun(0));
        } catch (Exception e) {
            
            logger.error(e.getMessage(), e);
        }
    }
    
    /*
     * <p>
     * Called at the end of the XML document
     * 
     * @see org.xml.sax.helpers.DefaultHandler#endDocument()
     */
    @Override
    public void endDocument() throws SAXException {
        
        try (FileOutputStream os = new FileOutputStream(Paths.get(this.docName).toFile())) {
            
            if (null != this.wordDoc) {
                
                // write the contents of the virtual POI document into a real
                // word document
                this.wordDoc.write(os);
                
                try {
                    
                    this.wordDoc.close();
                } catch (IOException e) {
                    
                    logger.error(e.getMessage(), e);
                }
            }
        } catch (Exception e) {
            
            logger.error(e.getMessage(), e);
        }
    }
    
    /*
     * <p>
     * Called at the end of the element
     * 
     * @see org.xml.sax.helpers.DefaultHandler#endElement(java.lang.String,
     * java.lang.String, java.lang.String)
     */
    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        
        logger.info(String.format("Element:: %s is going out of scope", qName));
    }
    
    /**
     * <p>
     * Gets the alignment of the paragraph, default being left
     * 
     * @param alignment
     *            The alignemnt mentioned in the XML
     * @return The paragraph alignment in POI
     */
    private ParagraphAlignment getAlignment(String alignment) {
        
        switch (alignment.toLowerCase()) {
            
            case "center":
                return ParagraphAlignment.CENTER;
            
            case "left":
                return ParagraphAlignment.LEFT;
            
            case "right":
                return ParagraphAlignment.RIGHT;
            
            default:
                return ParagraphAlignment.LEFT;
        }
    }
    
    /*
     * <p>
     * Called at the start of the element
     * 
     * @see org.xml.sax.helpers.DefaultHandler#startElement(java.lang.String,
     * java.lang.String, java.lang.String, org.xml.sax.Attributes)
     */
    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        
        switch (qName.toLowerCase()) {
            
            case "proposal": {
                
                // proposal handler, called when the <proposal> tag is
                // encountered
                this.docName = attributes.getValue("docName");
                
                XWPFHeader header = this.wordDoc.createHeader(HeaderFooterType.DEFAULT);
                
                // add header content
                XWPFRun run = header.createParagraph().createRun();
                run.setFontSize(7);
                run.setFontFamily("Times New Roman");
                run.setText(
                        attributes.getValue("headerValue") != null ? attributes.getValue("headerValue").toUpperCase()
                                : "");
                
                // add footer content
                XWPFFooter footer = this.wordDoc.createFooter(HeaderFooterType.DEFAULT);
                footer.createParagraph().createRun().setText("template text, replace with page nbr");
                footer.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                
                break;
            }
            
            case "title":
            case "sectiontitle":
            case "sectionparagraph": {
                
                // title handler, called when the tag title is encountered
                this.currentTextBody = this.wordDoc.createParagraph();
                this.currentTextBody.setAlignment(getAlignment(attributes.getValue("alignment")));
                
                // since the proposal is double spaced
                this.currentTextBody.setSpacingBetween(2.0);
                
                XWPFRun currentTextBodyRun = this.currentTextBody.createRun();
                currentTextBodyRun.setFontFamily(attributes.getValue("fontName"));
                currentTextBodyRun.setFontSize(Integer.parseInt(attributes.getValue("fontSize")));
                
                // for font styles
                switch (attributes.getValue("fontStyle").toLowerCase()) {
                    
                    case "bold": {
                        currentTextBodyRun.setBold(true);
                        break;
                    }
                    
                    case "italic": {
                        currentTextBodyRun.setItalic(true);
                        break;
                    }
                    
                    case "underlined": {
                        currentTextBodyRun.setUnderline(UnderlinePatterns.DASH);
                        break;
                    }
                }
                
                // if (qName.equalsIgnoreCase("sectiontitle")) {
                //
                // this.currentTextBody.setStyle("Heading 1");
                // }
                
                break;
            }
            
            case "section": {
                
                logger.info("beginning new section");
                break;
            }
            
            case "sectiontable": {
                
                // sectionTable handler
                // make a table and add the number of rows and columns as needed
                XWPFTable table = this.wordDoc.createTable();
                
                int nbrRows = Integer.parseInt(attributes.getValue("rows"));
                int nbrCols = Integer.parseInt(attributes.getValue("cols"));
                
                // since default table has 1 row and 1 column
                // add the rows to the table
                for (int i = 1; i < nbrRows; i++) {
                    
                    table.createRow();
                }
                
                // add columns to all the rows of the table
                for (int i = 1; i < nbrCols; i++) {
                    
                    table.addNewCol();
                }
                
                // add in place holder for the table
                for (XWPFTableRow row : table.getRows()) {
                    
                    for (ICell cell : row.getTableICells()) {
                        
                        ((XWPFTableCell) cell).setText(" cell template ");
                    }
                }
                
                break;
            }
            
            case "pagebreak": {
                // start new page
                this.wordDoc.createParagraph().setPageBreak(true);
                break;
            }
            
            case "linebreak": {
                
                if (null != this.currentTextBody) {
                    
                    this.currentTextBody.createRun().addCarriageReturn();
                }
                
                break;
            }
            // table of contents API not working properly
            // case "toc": {
            // this.wordDoc.createTOC();
            //
            // break;
            // }
        }
    }
    
    /*
     * <p>
     * Called when characters are encountered, like when content of the element
     * is encountered
     * 
     * @see org.xml.sax.helpers.DefaultHandler#characters(char[], int, int)
     */
    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        
        // content handler -- content of the element that is currently under
        // processing
        
        if (null != this.currentTextBody && null != this.currentTextBody.getIRuns()
                && this.currentTextBody.getIRuns().size() > 0) {
            
            ((XWPFRun) this.currentTextBody.getIRuns().get(0)).setText(new String(ch, start, length));
        }
        
        // logger.info(String.format("Text written :: %s", new String(ch, start,
        // length)));
    }
}
