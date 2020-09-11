package xyz.anthony.fieldReview;

import java.math.BigInteger;

import javax.xml.bind.JAXBElement;

import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.Ftr;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Text;

public class FooterBuilder {
    public static Ftr createIt() {

        org.docx4j.wml.ObjectFactory wmlObjectFactory = new org.docx4j.wml.ObjectFactory();
        
        Ftr ftr = wmlObjectFactory.createFtr(); 
            // Create object for p
            P p = wmlObjectFactory.createP(); 
            RFonts rfonts = wmlObjectFactory.createRFonts();
            rfonts.setAscii("Times New Roman");
            rfonts.setHAnsi("Times New Roman");
            rfonts.setCs("Times New Roman");
            ftr.getContent().add( p); 
                // Create object for r
                R r = wmlObjectFactory.createR(); 
                p.getContent().add( r); 
                    // Create object for t (wrapped in JAXBElement) 
                    Text text = wmlObjectFactory.createText(); 
                    JAXBElement<org.docx4j.wml.Text> textWrapped = wmlObjectFactory.createRT(text); 
                    r.getContent().add( textWrapped); 
                        text.setValue( "Site Reports are valid to specific work and are subject to review should additional information become available. The company is not responsible for any views or opinions expressed by personnel performing work. The company will not be responsible for deviation beyond normal allowance limits."); 
                    // Create object for rPr
                    RPr rpr = wmlObjectFactory.createRPr(); 
                    r.setRPr(rpr); 
                        // Create object for sz
                        HpsMeasure hpsmeasure = wmlObjectFactory.createHpsMeasure(); 
                        rpr.setSz(hpsmeasure); 
                            hpsmeasure.setVal( BigInteger.valueOf( 14) ); 
                        // Create object for lang
                        CTLanguage language = wmlObjectFactory.createCTLanguage(); 
                        rpr.setLang(language); 
                            language.setVal( "en-CA"); 
                        rpr.setRFonts(rfonts);
                // Create object for pPr
                PPr ppr = wmlObjectFactory.createPPr(); 
                p.setPPr(ppr); 
                    // Create object for rPr
                    ParaRPr pararpr = wmlObjectFactory.createParaRPr(); 
                    ppr.setRPr(pararpr); 
                        // Create object for sz
                        HpsMeasure hpsmeasure2 = wmlObjectFactory.createHpsMeasure(); 
                        pararpr.setSz(hpsmeasure2); 
                            hpsmeasure2.setVal( BigInteger.valueOf( 14) ); 
                        // Create object for lang
                        CTLanguage language2 = wmlObjectFactory.createCTLanguage(); 
                        pararpr.setLang(language2); 
                            language2.setVal( "en-CA"); 
                    // Create object for pStyle
                    PPrBase.PStyle pprbasepstyle = wmlObjectFactory.createPPrBasePStyle(); 
                    ppr.setPStyle(pprbasepstyle); 
                        pprbasepstyle.setVal( "Footer"); 
        
        return ftr;
        }
}
