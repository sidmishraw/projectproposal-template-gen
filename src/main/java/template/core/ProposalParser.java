/**
 * Project: ProjectProposalIEEE
 * Package: template.core
 * File: ProposalParser.java
 * 
 * @author sidmishraw
 *         Last modified: Oct 6, 2017 7:36:51 PM
 */
package template.core;

import java.io.IOException;
import java.nio.file.Paths;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import template.core.handler.ProposalHandler;

/**
 * <p>
 * Parsing driver
 * 
 * @author sidmishraw
 *
 *         Qualified Name: template.core.ProposalParser
 *
 */
public class ProposalParser {
    
    private static final Logger logger = LoggerFactory.getLogger(ProposalHandler.class);
    
    /**
     * Main thread for the parser
     * 
     * @param args
     */
    public static void main(String[] args) {
        
        try {
            
            SAXParserFactory factory = SAXParserFactory.newInstance();
            SAXParser parser = factory.newSAXParser();
            ProposalHandler handler = new ProposalHandler();
            parser.parse(Paths.get("proposal.xml").toFile(), handler);
        } catch (ParserConfigurationException | SAXException | IOException e) {
            
            logger.error(e.getMessage(), e);
        }
    }
}
