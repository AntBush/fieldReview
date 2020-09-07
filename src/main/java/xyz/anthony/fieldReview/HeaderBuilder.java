package xyz.anthony.fieldReview;

import java.math.BigInteger;

import javax.xml.bind.JAXBElement;

import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Color;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Text;

/*
    Header builder completely ripped off from the generator
*/

public class HeaderBuilder {
    public static Hdr createIt() {

        org.docx4j.wml.ObjectFactory wmlObjectFactory = new org.docx4j.wml.ObjectFactory();
        
        Hdr hdr = wmlObjectFactory.createHdr(); 
            // Create object for p
            P p = wmlObjectFactory.createP(); 
            hdr.getContent().add( p); 
                // Create object for r
                R r = wmlObjectFactory.createR(); 
                p.getContent().add( r); 
                    // Create object for t (wrapped in JAXBElement) 
                    Text text = wmlObjectFactory.createText(); 
                    JAXBElement<org.docx4j.wml.Text> textWrapped = wmlObjectFactory.createRT(text); 
                    r.getContent().add( textWrapped); 
                        text.setValue( "FIELD REVIEW CONSULTANTS LTD"); 
                    // Create object for rPr
                    RPr rpr = wmlObjectFactory.createRPr(); 
                    r.setRPr(rpr); 
                        // Create object for b
                        BooleanDefaultTrue booleandefaulttrue = wmlObjectFactory.createBooleanDefaultTrue(); 
                        rpr.setB(booleandefaulttrue); 
                        // Create object for bCs
                        BooleanDefaultTrue booleandefaulttrue2 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        rpr.setBCs(booleandefaulttrue2); 
                        // Create object for color
                        Color color = wmlObjectFactory.createColor(); 
                        rpr.setColor(color); 
                            color.setVal( "943634"); 
                            color.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_2);
                            color.setThemeShade( "BF"); 
                        // Create object for sz
                        HpsMeasure hpsmeasure = wmlObjectFactory.createHpsMeasure(); 
                        rpr.setSz(hpsmeasure); 
                            hpsmeasure.setVal( BigInteger.valueOf( 36) ); 
                        // Create object for szCs
                        HpsMeasure hpsmeasure2 = wmlObjectFactory.createHpsMeasure(); 
                        rpr.setSzCs(hpsmeasure2); 
                            hpsmeasure2.setVal( BigInteger.valueOf( 36) ); 
                // Create object for pPr
                PPr ppr = wmlObjectFactory.createPPr(); 
                p.setPPr(ppr); 
                    // Create object for rPr
                    ParaRPr pararpr = wmlObjectFactory.createParaRPr(); 
                    ppr.setRPr(pararpr); 
                        // Create object for b
                        BooleanDefaultTrue booleandefaulttrue3 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        pararpr.setB(booleandefaulttrue3); 
                        // Create object for bCs
                        BooleanDefaultTrue booleandefaulttrue4 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        pararpr.setBCs(booleandefaulttrue4); 
                        // Create object for color
                        Color color2 = wmlObjectFactory.createColor(); 
                        pararpr.setColor(color2); 
                            color2.setVal( "943634"); 
                            color2.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_2);
                            color2.setThemeShade( "BF"); 
                        // Create object for sz
                        HpsMeasure hpsmeasure3 = wmlObjectFactory.createHpsMeasure(); 
                        pararpr.setSz(hpsmeasure3); 
                            hpsmeasure3.setVal( BigInteger.valueOf( 36) ); 
                        // Create object for szCs
                        HpsMeasure hpsmeasure4 = wmlObjectFactory.createHpsMeasure(); 
                        pararpr.setSzCs(hpsmeasure4); 
                            hpsmeasure4.setVal( BigInteger.valueOf( 36) ); 
                    // Create object for pStyle
                    PPrBase.PStyle pprbasepstyle = wmlObjectFactory.createPPrBasePStyle(); 
                    ppr.setPStyle(pprbasepstyle); 
                        pprbasepstyle.setVal( "Header"); 
                    // Create object for jc
                    Jc jc = wmlObjectFactory.createJc(); 
                    ppr.setJc(jc); 
                        jc.setVal(org.docx4j.wml.JcEnumeration.CENTER);
            // Create object for p
            P p2 = wmlObjectFactory.createP(); 
            hdr.getContent().add( p2); 
                // Create object for r
                R r2 = wmlObjectFactory.createR(); 
                p2.getContent().add( r2); 
                    // Create object for t (wrapped in JAXBElement) 
                    Text text2 = wmlObjectFactory.createText(); 
                    JAXBElement<org.docx4j.wml.Text> textWrapped2 = wmlObjectFactory.createRT(text2); 
                    r2.getContent().add( textWrapped2); 
                        text2.setValue( "#104- 93 Dundas Street East, Mississauga, ON, L5A 1W7"); 
                    // Create object for rPr
                    RPr rpr2 = wmlObjectFactory.createRPr(); 
                    r2.setRPr(rpr2); 
                        // Create object for b
                        BooleanDefaultTrue booleandefaulttrue5 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        rpr2.setB(booleandefaulttrue5); 
                        // Create object for bCs
                        BooleanDefaultTrue booleandefaulttrue6 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        rpr2.setBCs(booleandefaulttrue6); 
                        // Create object for color
                        Color color3 = wmlObjectFactory.createColor(); 
                        rpr2.setColor(color3); 
                            color3.setVal( "943634"); 
                            color3.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_2);
                            color3.setThemeShade( "BF"); 
                // Create object for pPr
                PPr ppr2 = wmlObjectFactory.createPPr(); 
                p2.setPPr(ppr2); 
                    // Create object for rPr
                    ParaRPr pararpr2 = wmlObjectFactory.createParaRPr(); 
                    ppr2.setRPr(pararpr2); 
                        // Create object for b
                        BooleanDefaultTrue booleandefaulttrue7 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        pararpr2.setB(booleandefaulttrue7); 
                        // Create object for bCs
                        BooleanDefaultTrue booleandefaulttrue8 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        pararpr2.setBCs(booleandefaulttrue8); 
                        // Create object for color
                        Color color4 = wmlObjectFactory.createColor(); 
                        pararpr2.setColor(color4); 
                            color4.setVal( "943634"); 
                            color4.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_2);
                            color4.setThemeShade( "BF"); 
                    // Create object for pStyle
                    PPrBase.PStyle pprbasepstyle2 = wmlObjectFactory.createPPrBasePStyle(); 
                    ppr2.setPStyle(pprbasepstyle2); 
                        pprbasepstyle2.setVal( "Header"); 
                    // Create object for jc
                    Jc jc2 = wmlObjectFactory.createJc(); 
                    ppr2.setJc(jc2); 
                        jc2.setVal(org.docx4j.wml.JcEnumeration.CENTER);
            // Create object for p
            P p3 = wmlObjectFactory.createP(); 
            hdr.getContent().add( p3); 
                // Create object for r
                R r3 = wmlObjectFactory.createR(); 
                p3.getContent().add( r3); 
                    // Create object for t (wrapped in JAXBElement) 
                    Text text3 = wmlObjectFactory.createText(); 
                    JAXBElement<org.docx4j.wml.Text> textWrapped3 = wmlObjectFactory.createRT(text3); 
                    r3.getContent().add( textWrapped3); 
                        text3.setValue( "Tel: 647.977.4766, Fax: 647.9774765"); 
                    // Create object for rPr
                    RPr rpr3 = wmlObjectFactory.createRPr(); 
                    r3.setRPr(rpr3); 
                        // Create object for b
                        BooleanDefaultTrue booleandefaulttrue9 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        rpr3.setB(booleandefaulttrue9); 
                        // Create object for bCs
                        BooleanDefaultTrue booleandefaulttrue10 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        rpr3.setBCs(booleandefaulttrue10); 
                        // Create object for color
                        Color color5 = wmlObjectFactory.createColor(); 
                        rpr3.setColor(color5); 
                            color5.setVal( "943634"); 
                            color5.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_2);
                            color5.setThemeShade( "BF"); 
                // Create object for pPr
                PPr ppr3 = wmlObjectFactory.createPPr(); 
                p3.setPPr(ppr3); 
                    // Create object for rPr
                    ParaRPr pararpr3 = wmlObjectFactory.createParaRPr(); 
                    ppr3.setRPr(pararpr3); 
                        // Create object for b
                        BooleanDefaultTrue booleandefaulttrue11 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        pararpr3.setB(booleandefaulttrue11); 
                        // Create object for bCs
                        BooleanDefaultTrue booleandefaulttrue12 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        pararpr3.setBCs(booleandefaulttrue12); 
                        // Create object for color
                        Color color6 = wmlObjectFactory.createColor(); 
                        pararpr3.setColor(color6); 
                            color6.setVal( "943634"); 
                            color6.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_2);
                            color6.setThemeShade( "BF"); 
                    // Create object for pStyle
                    PPrBase.PStyle pprbasepstyle3 = wmlObjectFactory.createPPrBasePStyle(); 
                    ppr3.setPStyle(pprbasepstyle3); 
                        pprbasepstyle3.setVal( "Header"); 
                    // Create object for jc
                    Jc jc3 = wmlObjectFactory.createJc(); 
                    ppr3.setJc(jc3); 
                        jc3.setVal(org.docx4j.wml.JcEnumeration.CENTER);
            // Create object for p
            P p4 = wmlObjectFactory.createP(); 
            hdr.getContent().add( p4); 
                // Create object for r
                R r4 = wmlObjectFactory.createR(); 
                p4.getContent().add( r4); 
                    // Create object for t (wrapped in JAXBElement) 
                    Text text4 = wmlObjectFactory.createText(); 
                    JAXBElement<org.docx4j.wml.Text> textWrapped4 = wmlObjectFactory.createRT(text4); 
                    r4.getContent().add( textWrapped4); 
                        text4.setValue( "E-Mail Address: fmessiha@fieldreviewconsultants.com"); 
                    // Create object for rPr
                    RPr rpr4 = wmlObjectFactory.createRPr(); 
                    r4.setRPr(rpr4); 
                        // Create object for b
                        BooleanDefaultTrue booleandefaulttrue13 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        rpr4.setB(booleandefaulttrue13); 
                        // Create object for bCs
                        BooleanDefaultTrue booleandefaulttrue14 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        rpr4.setBCs(booleandefaulttrue14); 
                        // Create object for color
                        Color color7 = wmlObjectFactory.createColor(); 
                        rpr4.setColor(color7); 
                            color7.setVal( "943634"); 
                            color7.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_2);
                            color7.setThemeShade( "BF"); 
                // Create object for pPr
                PPr ppr4 = wmlObjectFactory.createPPr(); 
                p4.setPPr(ppr4); 
                    // Create object for rPr
                    ParaRPr pararpr4 = wmlObjectFactory.createParaRPr(); 
                    ppr4.setRPr(pararpr4); 
                        // Create object for b
                        BooleanDefaultTrue booleandefaulttrue15 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        pararpr4.setB(booleandefaulttrue15); 
                        // Create object for bCs
                        BooleanDefaultTrue booleandefaulttrue16 = wmlObjectFactory.createBooleanDefaultTrue(); 
                        pararpr4.setBCs(booleandefaulttrue16); 
                        // Create object for color
                        Color color8 = wmlObjectFactory.createColor(); 
                        pararpr4.setColor(color8); 
                            color8.setVal( "943634"); 
                            color8.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_2);
                            color8.setThemeShade( "BF"); 
                    // Create object for pStyle
                    PPrBase.PStyle pprbasepstyle4 = wmlObjectFactory.createPPrBasePStyle(); 
                    ppr4.setPStyle(pprbasepstyle4); 
                        pprbasepstyle4.setVal( "Header"); 
                    // Create object for jc
                    Jc jc4 = wmlObjectFactory.createJc(); 
                    ppr4.setJc(jc4); 
                        jc4.setVal(org.docx4j.wml.JcEnumeration.CENTER);
        
        return hdr;
        }
        
}
