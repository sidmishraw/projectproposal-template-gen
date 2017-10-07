/**
 * Project: ProjectProposalIEEE
 * Package: playground
 * File: PlaygroundInitial.java
 * 
 * @author sidmishraw
 *         Last modified: Oct 6, 2017 4:10:25 PM
 */
package playground;

import java.io.FileOutputStream;
import java.nio.file.Paths;

import org.apache.poi.xwpf.usermodel.ICell;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <p>
 * This is the initial playground I'll be using for testing out the various APIs
 * in the Apache POI's XWPF format.
 * 
 * Main objective: Create a word doc template that is ready for filling in stuff
 * 
 * @author sidmishraw
 *
 *         Qualified Name: playground.PlaygroundInitial
 *
 */
public class PlaygroundInitial {
    
    private static final Logger logger = LoggerFactory.getLogger(PlaygroundInitial.class);
    
    /**
     * <p>
     * This is the first test, I'm going to create a new Word Document
     * using the Apache POI API.
     * 
     * @param docName
     *            The name of the word document
     */
    private static void createBlankWordDoc(String docName) {
        
        // document : this is the blank document
        // document is a resource and needs to be in the try with resources
        // block just like the output stream
        try (XWPFDocument document = new XWPFDocument();
                FileOutputStream os = new FileOutputStream(
                        Paths.get("/", "Users", "sidmishraw", "Desktop", docName).toFile())) {
            
            // write the logic document you created into the actual document
            // using the
            // outout stream you opened
            document.write(os);
        } catch (Exception e) {
            
            logger.error(e.getMessage(), e);
        }
    }
    
    /**
     * This is the test for creating paragraph
     * 
     * @param docName
     *            The document name
     * @param text
     *            The text to be inserted into the paragraph
     */
    private static void createParagraph(String docName, String text) {
        
        try (XWPFDocument document = new XWPFDocument();
                FileOutputStream os = new FileOutputStream(
                        Paths.get("/", "Users", "sidmishraw", "Desktop", docName).toFile())) {
            
            logger.info("Creating document: " + docName + " with paragraph with text: " + text);
            
            XWPFParagraph p = document.createParagraph();
            
            XWPFRun pRun = p.createRun();
            
            pRun.setText(text);
            
            document.write(os);
            
            logger.info("Created document: " + docName);
        } catch (Exception e) {
            
            logger.error(e.getMessage(), e);
        }
    }
    
    /**
     * <p>
     * This tests the table creation APIs from apache POI
     * 
     * @param docName
     *            The document name
     */
    private static void createTable(String docName) {
        
        try (XWPFDocument document = new XWPFDocument();
                FileOutputStream os = new FileOutputStream(
                        Paths.get("/", "Users", "sidmishraw", "Desktop", docName).toFile())) {
            
            logger.info("Creating document: " + docName);
            
            // creates a table with 1 row and 1 col by default
            XWPFTable table = document.createTable();
            
            // add another row
            table.createRow();
            
            // adds a new col to each row of the table
            table.addNewCol();
            
            // now my table has 2 cols and 1 row
            for (XWPFTableRow row : table.getRows()) {
                
                // cells
                // the getTableCells and the getTableICells dont return the same
                // number of cells
                // don't know why, need to read up the DOcs to be clear.
                for (ICell cell : row.getTableICells()) {
                    
                    ((XWPFTableCell) cell)
                            .setText(String.format("This is row: %s and cell: %s", row.toString(), cell.toString()));
                }
            }
            
            document.write(os);
            
            logger.info("Created document: " + docName);
        } catch (Exception e) {
            
            logger.error(e.getMessage(), e);
        }
    }
    
    /**
     * @param args
     */
    public static void main(String[] args) {
        
        // createBlankWordDoc("first.docx");
        // createParagraph("first.docx", "Hey there, Hellow");
        createTable("first.docx");
    }
}
