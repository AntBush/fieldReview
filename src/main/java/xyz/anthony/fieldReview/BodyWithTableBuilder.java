package xyz.anthony.fieldReview;

import java.math.BigInteger;
import java.util.List;

import javax.xml.bind.JAXBElement;

import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTColumns;
import org.docx4j.wml.CTDocGrid;
import org.docx4j.wml.CTHeight;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CTTabStop;
import org.docx4j.wml.CTTblLook;
import org.docx4j.wml.Color;
import org.docx4j.wml.Document;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.R.Tab;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.Tabs;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblGrid;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.docx4j.wml.TcPrInner.TcBorders;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.TrPr;

public class BodyWithTableBuilder {
    public static Document createIt(FieldReviewData reviewData) {

        org.docx4j.wml.ObjectFactory wmlObjectFactory = new org.docx4j.wml.ObjectFactory();
        PPrBase.Spacing pprBaseSpacing = wmlObjectFactory.createPPrBaseSpacing();
        pprBaseSpacing.setAfter(BigInteger.valueOf(0));


        Document document = wmlObjectFactory.createDocument();
        // Create object for body
        Body body = wmlObjectFactory.createBody();
        document.setBody(body);
        // Create object for p
        P p = wmlObjectFactory.createP();
        body.getContent().add(p);
        // Create object for pPr
        PPr ppr = wmlObjectFactory.createPPr();
        p.setPPr(ppr);
        // Create object for rPr
        ParaRPr pararpr = wmlObjectFactory.createParaRPr();
        ppr.setRPr(pararpr);
        ppr.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts = wmlObjectFactory.createRFonts();
        pararpr.setRFonts(rfonts);
        rfonts.setAscii("Arial");
        rfonts.setHAnsi("Arial");
        rfonts.setCs("Arial");
        // Create object for tbl (wrapped in JAXBElement)
        Tbl tbl = wmlObjectFactory.createTbl();
        JAXBElement<org.docx4j.wml.Tbl> tblWrapped = wmlObjectFactory.createBodyTbl(tbl);
        body.getContent().add(tblWrapped);
        // Create object for tr
        Tr tr = wmlObjectFactory.createTr();
        tbl.getContent().add(tr);
        // Create object for tc (wrapped in JAXBElement)
        Tc tc = wmlObjectFactory.createTc();
        JAXBElement<org.docx4j.wml.Tc> tcWrapped = wmlObjectFactory.createTrTc(tc);
        tr.getContent().add(tcWrapped);
        // Create object for p
        P p2 = wmlObjectFactory.createP();
        tc.getContent().add(p2);
        // Create object for r
        R r = wmlObjectFactory.createR();
        p2.getContent().add(r);
        // Create object for t (wrapped in JAXBElement)
        Text text = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped = wmlObjectFactory.createRT(text);
        r.getContent().add(textWrapped);
        text.setValue("Field Review Reports:");
        // Create object for rPr
        RPr rpr = wmlObjectFactory.createRPr();
        r.setRPr(rpr);
        // Create object for rFonts
        RFonts rfonts2 = wmlObjectFactory.createRFonts();
        rpr.setRFonts(rfonts2);
        rfonts2.setAscii("Arial");
        rfonts2.setHAnsi("Arial");
        rfonts2.setCs("Arial");
        rpr.setB(wmlObjectFactory.createBooleanDefaultTrue());
        // Create object for color
        Color color = wmlObjectFactory.createColor();
        rpr.setColor(color);
        color.setVal("000000");
        color.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r2 = wmlObjectFactory.createR();
        p2.getContent().add(r2);
        // Create object for t (wrapped in JAXBElement)
        Text text2 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped2 = wmlObjectFactory.createRT(text2);
        r2.getContent().add(textWrapped2);
        text2.setValue(" ");
        text2.setSpace("preserve");
        // Create object for rPr
        RPr rpr2 = wmlObjectFactory.createRPr();
        r2.setRPr(rpr2);
        // Create object for rFonts
        RFonts rfonts3 = wmlObjectFactory.createRFonts();
        rpr2.setRFonts(rfonts3);
        rfonts3.setAscii("Arial");
        rfonts3.setHAnsi("Arial");
        rfonts3.setCs("Arial");
        // Create object for color
        Color color2 = wmlObjectFactory.createColor();
        rpr2.setColor(color2);
        color2.setVal("000000");
        color2.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r3 = wmlObjectFactory.createR();
        p2.getContent().add(r3);
        // Create object for t (wrapped in JAXBElement)
        Text text3 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped3 = wmlObjectFactory.createRT(text3);
        r3.getContent().add(textWrapped3);
        //TODO: Field Review Report Number variable
        text3.setValue(reviewData.getReportNumber());
        text3.setSpace("preserve");
        // Create object for rPr
        RPr rpr3 = wmlObjectFactory.createRPr();
        r3.setRPr(rpr3);
        // Create object for rFonts
        RFonts rfonts4 = wmlObjectFactory.createRFonts();
        rpr3.setRFonts(rfonts4);
        rfonts4.setAscii("Arial");
        rfonts4.setHAnsi("Arial");
        rfonts4.setCs("Arial");
        // Create object for b
        // BooleanDefaultTrue booleandefaulttrue = wmlObjectFactory.createBooleanDefaultTrue();
        // rpr3.setB(booleandefaulttrue);
        // Create object for color
        Color color3 = wmlObjectFactory.createColor();
        rpr3.setColor(color3);
        color3.setVal("1F497D");
        color3.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for r
        R r4 = wmlObjectFactory.createR();
        // p2.getContent().add(r4);
        // Create object for t (wrapped in JAXBElement)
        Text text4 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped4 = wmlObjectFactory.createRT(text4);
        r4.getContent().add(textWrapped4);
        text4.setValue("20");
        // Create object for rPr
        RPr rpr4 = wmlObjectFactory.createRPr();
        r4.setRPr(rpr4);
        // Create object for rFonts
        RFonts rfonts5 = wmlObjectFactory.createRFonts();
        rpr4.setRFonts(rfonts5);
        rfonts5.setAscii("Arial");
        rfonts5.setHAnsi("Arial");
        rfonts5.setCs("Arial");
        // Create object for b
        // BooleanDefaultTrue booleandefaulttrue2 = wmlObjectFactory.createBooleanDefaultTrue();
        // rpr4.setB(booleandefaulttrue2);
        // Create object for color
        Color color4 = wmlObjectFactory.createColor();
        rpr4.setColor(color4);
        color4.setVal("1F497D");
        color4.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for pPr
        PPr ppr2 = wmlObjectFactory.createPPr();
        p2.setPPr(ppr2);
        // Create object for rPr
        ParaRPr pararpr2 = wmlObjectFactory.createParaRPr();
        ppr2.setRPr(pararpr2);
        ppr2.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts6 = wmlObjectFactory.createRFonts();
        pararpr2.setRFonts(rfonts6);
        rfonts6.setAscii("Arial");
        rfonts6.setHAnsi("Arial");
        rfonts6.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue3 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr2.setB(booleandefaulttrue3);
        // Create object for color
        Color color5 = wmlObjectFactory.createColor();
        pararpr2.setColor(color5);
        color5.setVal("000000");
        color5.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle = wmlObjectFactory.createPPrBasePStyle();
        ppr2.setPStyle(pprbasepstyle);
        pprbasepstyle.setVal("BodyText");
        // Create object for tabs
        Tabs tabs = wmlObjectFactory.createTabs();
        ppr2.setTabs(tabs);
        // Create object for tab
        CTTabStop tabstop = wmlObjectFactory.createCTTabStop();
        tabs.getTab().add(tabstop);
        tabstop.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop.setPos(BigInteger.valueOf(3470));
        // Create object for tab
        CTTabStop tabstop2 = wmlObjectFactory.createCTTabStop();
        tabs.getTab().add(tabstop2);
        tabstop2.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop2.setPos(BigInteger.valueOf(3580));
        // Create object for tab
        CTTabStop tabstop3 = wmlObjectFactory.createCTTabStop();
        tabs.getTab().add(tabstop3);
        tabstop3.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop3.setPos(BigInteger.valueOf(5832));
        // Create object for tab
        CTTabStop tabstop4 = wmlObjectFactory.createCTTabStop();
        tabs.getTab().add(tabstop4);
        tabstop4.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop4.setPos(BigInteger.valueOf(6690));
        // Create object for p
        P p3 = wmlObjectFactory.createP();
        tc.getContent().add(p3);
        // Create object for r
        R r5 = wmlObjectFactory.createR();
        p3.getContent().add(r5);
        // Create object for t (wrapped in JAXBElement)
        Text text5 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped5 = wmlObjectFactory.createRT(text5);
        r5.getContent().add(textWrapped5);
        text5.setValue("Common Element No:");
        // Create object for rPr
        RPr rpr5 = wmlObjectFactory.createRPr();
        r5.setRPr(rpr5);
        // Create object for rFonts
        RFonts rfonts7 = wmlObjectFactory.createRFonts();
        rpr5.setRFonts(rfonts7);
        rfonts7.setAscii("Arial");
        rfonts7.setHAnsi("Arial");
        rfonts7.setCs("Arial");
        rpr5.setB(wmlObjectFactory.createBooleanDefaultTrue());
        // Create object for color
        Color color6 = wmlObjectFactory.createColor();
        rpr5.setColor(color6);
        color6.setVal("000000");
        color6.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r6 = wmlObjectFactory.createR();
        p3.getContent().add(r6);
        // Create object for t (wrapped in JAXBElement)
        Text text6 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped6 = wmlObjectFactory.createRT(text6);
        r6.getContent().add(textWrapped6);
        //TODO: Add variable value for Common Element Number
        text6.setValue(reviewData.getCommonElementNumber());
        // Create object for rPr
        RPr rpr6 = wmlObjectFactory.createRPr();
        r6.setRPr(rpr6);
        // Create object for b
        // BooleanDefaultTrue booleandefaulttrue4 = wmlObjectFactory.createBooleanDefaultTrue();
        // rpr6.setB(booleandefaulttrue4);
        // Create object for color
        Color color7 = wmlObjectFactory.createColor();
        rpr6.setColor(color7);
        color7.setVal("1F497D");
        color7.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for pPr
        PPr ppr3 = wmlObjectFactory.createPPr();
        p3.setPPr(ppr3);
        // Create object for rPr
        ParaRPr pararpr3 = wmlObjectFactory.createParaRPr();
        ppr3.setRPr(pararpr3);
        ppr3.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts8 = wmlObjectFactory.createRFonts();
        pararpr3.setRFonts(rfonts8);
        rfonts8.setAscii("Arial");
        rfonts8.setHAnsi("Arial");
        rfonts8.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue5 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr3.setB(booleandefaulttrue5);
        // Create object for color
        Color color8 = wmlObjectFactory.createColor();
        pararpr3.setColor(color8);
        color8.setVal("000000");
        color8.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle2 = wmlObjectFactory.createPPrBasePStyle();
        ppr3.setPStyle(pprbasepstyle2);
        pprbasepstyle2.setVal("BodyText");
        // Create object for tabs
        Tabs tabs2 = wmlObjectFactory.createTabs();
        ppr3.setTabs(tabs2);
        // Create object for tab
        CTTabStop tabstop5 = wmlObjectFactory.createCTTabStop();
        tabs2.getTab().add(tabstop5);
        tabstop5.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop5.setPos(BigInteger.valueOf(3470));
        // Create object for tab
        CTTabStop tabstop6 = wmlObjectFactory.createCTTabStop();
        tabs2.getTab().add(tabstop6);
        tabstop6.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop6.setPos(BigInteger.valueOf(3580));
        // Create object for tab
        CTTabStop tabstop7 = wmlObjectFactory.createCTTabStop();
        tabs2.getTab().add(tabstop7);
        tabstop7.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop7.setPos(BigInteger.valueOf(5832));
        // Create object for tab
        CTTabStop tabstop8 = wmlObjectFactory.createCTTabStop();
        tabs2.getTab().add(tabstop8);
        tabstop8.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop8.setPos(BigInteger.valueOf(6690));
        // Create object for tcPr
        TcPr tcpr = wmlObjectFactory.createTcPr();
        tc.setTcPr(tcpr);
        // Create object for tcW
        TblWidth tblwidth = wmlObjectFactory.createTblWidth();
        tcpr.setTcW(tblwidth);
        tblwidth.setW(BigInteger.valueOf(8199));
        tblwidth.setType("dxa");
        // Create object for gridSpan
        TcPrInner.GridSpan tcprinnergridspan = wmlObjectFactory.createTcPrInnerGridSpan();
        tcpr.setGridSpan(tcprinnergridspan);
        tcprinnergridspan.setVal(BigInteger.valueOf(2));
        // Create object for tcBorders
        TcPrInner.TcBorders tcprinnertcborders = wmlObjectFactory.createTcPrInnerTcBorders();
        tcpr.setTcBorders(tcprinnertcborders);
        // Create object for top
        CTBorder border = wmlObjectFactory.createCTBorder();
        tcprinnertcborders.setTop(border);
        border.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border.setColor("auto");
        border.setSz(BigInteger.valueOf(4));
        border.setSpace(BigInteger.valueOf(0));
        // Create object for left
        CTBorder border2 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders.setLeft(border2);
        border2.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border2.setColor("auto");
        border2.setSz(BigInteger.valueOf(4));
        border2.setSpace(BigInteger.valueOf(0));
        // Create object for right
        CTBorder border3 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders.setRight(border3);
        border3.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border3.setColor("auto");
        border3.setSz(BigInteger.valueOf(4));
        border3.setSpace(BigInteger.valueOf(0));
        // Create object for trPr
        TrPr trpr = wmlObjectFactory.createTrPr();
        tr.setTrPr(trpr);
        // Create object for trHeight (wrapped in JAXBElement)
        CTHeight height = wmlObjectFactory.createCTHeight();
        JAXBElement<org.docx4j.wml.CTHeight> heightWrapped = wmlObjectFactory.createCTTrPrBaseTrHeight(height);
        trpr.getCnfStyleOrDivIdOrGridBefore().add(heightWrapped);
        height.setVal(BigInteger.valueOf(576));
        // Create object for tr
        Tr tr2 = wmlObjectFactory.createTr();
        tbl.getContent().add(tr2);
        // Create object for tc (wrapped in JAXBElement)
        Tc tc2 = wmlObjectFactory.createTc();
        JAXBElement<org.docx4j.wml.Tc> tcWrapped2 = wmlObjectFactory.createTrTc(tc2);
        tr2.getContent().add(tcWrapped2);
        // Create object for p
        P p4 = wmlObjectFactory.createP();
        tc2.getContent().add(p4);
        // Create object for pPr
        PPr ppr4 = wmlObjectFactory.createPPr();
        p4.setPPr(ppr4);
        // Create object for rPr
        ParaRPr pararpr4 = wmlObjectFactory.createParaRPr();
        ppr4.setRPr(pararpr4);
        ppr4.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts9 = wmlObjectFactory.createRFonts();
        pararpr4.setRFonts(rfonts9);
        rfonts9.setAscii("Arial");
        rfonts9.setHAnsi("Arial");
        rfonts9.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue6 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr4.setB(booleandefaulttrue6);
        // Create object for color
        Color color9 = wmlObjectFactory.createColor();
        pararpr4.setColor(color9);
        color9.setVal("000000");
        color9.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle3 = wmlObjectFactory.createPPrBasePStyle();
        ppr4.setPStyle(pprbasepstyle3);
        pprbasepstyle3.setVal("BodyText");
        // Create object for p
        P p5 = wmlObjectFactory.createP();
        tc2.getContent().add(p5);
        // Create object for r
        R r7 = wmlObjectFactory.createR();
        p5.getContent().add(r7);
        // Create object for t (wrapped in JAXBElement)
        Text text7 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped7 = wmlObjectFactory.createRT(text7);
        r7.getContent().add(textWrapped7);
        text7.setValue("File No.:");
        // Create object for rPr
        RPr rpr7 = wmlObjectFactory.createRPr();
        r7.setRPr(rpr7);
        // Create object for rFonts
        RFonts rfonts10 = wmlObjectFactory.createRFonts();
        rpr7.setRFonts(rfonts10);
        rfonts10.setAscii("Arial");
        rfonts10.setHAnsi("Arial");
        rfonts10.setCs("Arial");
        rpr7.setB(wmlObjectFactory.createBooleanDefaultTrue());
        // Create object for color
        Color color10 = wmlObjectFactory.createColor();
        rpr7.setColor(color10);
        color10.setVal("000000");
        color10.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r8 = wmlObjectFactory.createR();
        p5.getContent().add(r8);
        // Create object for t (wrapped in JAXBElement)
        Text text8 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped8 = wmlObjectFactory.createRT(text8);
        r8.getContent().add(textWrapped8);
        text8.setValue(" ");
        text8.setSpace("preserve");
        // Create object for rPr
        RPr rpr8 = wmlObjectFactory.createRPr();
        r8.setRPr(rpr8);
        // Create object for rFonts
        RFonts rfonts11 = wmlObjectFactory.createRFonts();
        rpr8.setRFonts(rfonts11);
        rfonts11.setAscii("Arial");
        rfonts11.setHAnsi("Arial");
        rfonts11.setCs("Arial");
        // Create object for color
        Color color11 = wmlObjectFactory.createColor();
        rpr8.setColor(color11);
        color11.setVal("000000");
        color11.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r9 = wmlObjectFactory.createR();
        p5.getContent().add(r9);
        // Create object for t (wrapped in JAXBElement)
        Text text9 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped9 = wmlObjectFactory.createRT(text9);
        r9.getContent().add(textWrapped9);
        text9.setValue(reviewData.getFileNumber());
        // Create object for rPr
        RPr rpr9 = wmlObjectFactory.createRPr();
        r9.setRPr(rpr9);
        // Create object for rFonts
        RFonts rfonts12 = wmlObjectFactory.createRFonts();
        rpr9.setRFonts(rfonts12);
        rfonts12.setAscii("Arial");
        rfonts12.setHAnsi("Arial");
        rfonts12.setCs("Arial");
        // Create object for b
        // BooleanDefaultTrue booleandefaulttrue7 = wmlObjectFactory.createBooleanDefaultTrue(); remove bold for FRC 4915
        // rpr9.setB(booleandefaulttrue7);
        // Create object for color
        Color color12 = wmlObjectFactory.createColor();
        rpr9.setColor(color12);
        color12.setVal("1F497D");
        color12.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for pPr
        PPr ppr5 = wmlObjectFactory.createPPr();
        p5.setPPr(ppr5);
        // Create object for rPr
        ParaRPr pararpr5 = wmlObjectFactory.createParaRPr();
        ppr5.setRPr(pararpr5);
        ppr5.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts13 = wmlObjectFactory.createRFonts();
        pararpr5.setRFonts(rfonts13);
        rfonts13.setAscii("Arial");
        rfonts13.setHAnsi("Arial");
        rfonts13.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue8 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr5.setB(booleandefaulttrue8);
        // Create object for color
        Color color13 = wmlObjectFactory.createColor();
        pararpr5.setColor(color13);
        color13.setVal("000000");
        color13.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle4 = wmlObjectFactory.createPPrBasePStyle();
        ppr5.setPStyle(pprbasepstyle4);
        pprbasepstyle4.setVal("BodyText");
        // Create object for p
        P p6 = wmlObjectFactory.createP();
        tc2.getContent().add(p6);
        // Create object for pPr
        PPr ppr6 = wmlObjectFactory.createPPr();
        p6.setPPr(ppr6);
        // Create object for rPr
        ParaRPr pararpr6 = wmlObjectFactory.createParaRPr();
        ppr6.setRPr(pararpr6);
        ppr6.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts14 = wmlObjectFactory.createRFonts();
        pararpr6.setRFonts(rfonts14);
        rfonts14.setAscii("Arial");
        rfonts14.setHAnsi("Arial");
        rfonts14.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue9 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr6.setB(booleandefaulttrue9);
        // Create object for color
        Color color14 = wmlObjectFactory.createColor();
        pararpr6.setColor(color14);
        color14.setVal("000000");
        color14.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle5 = wmlObjectFactory.createPPrBasePStyle();
        ppr6.setPStyle(pprbasepstyle5);
        pprbasepstyle5.setVal("BodyText");
        // Create object for p
        P p7 = wmlObjectFactory.createP();
        tc2.getContent().add(p7);
        // Create object for r
        R r10 = wmlObjectFactory.createR();
        p7.getContent().add(r10);
        // Create object for t (wrapped in JAXBElement)
        Text text10 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped10 = wmlObjectFactory.createRT(text10);
        r10.getContent().add(textWrapped10);
        text10.setValue("Date:");
        // Create object for rPr
        RPr rpr10 = wmlObjectFactory.createRPr();
        r10.setRPr(rpr10);
        // Create object for rFonts
        RFonts rfonts15 = wmlObjectFactory.createRFonts();
        rpr10.setRFonts(rfonts15);
        rfonts15.setAscii("Arial");
        rfonts15.setHAnsi("Arial");
        rfonts15.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue10 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr10.setB(booleandefaulttrue10);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue11 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr10.setBCs(booleandefaulttrue11);
        // Create object for color
        Color color15 = wmlObjectFactory.createColor();
        rpr10.setColor(color15);
        color15.setVal("000000");
        color15.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r11 = wmlObjectFactory.createR();
        p7.getContent().add(r11);
        // Create object for t (wrapped in JAXBElement)
        Text text11 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped11 = wmlObjectFactory.createRT(text11);
        r11.getContent().add(textWrapped11);
        text11.setValue(" ");
        text11.setSpace("preserve");
        // Create object for rPr
        RPr rpr11 = wmlObjectFactory.createRPr();
        r11.setRPr(rpr11);
        // Create object for rFonts
        RFonts rfonts16 = wmlObjectFactory.createRFonts();
        rpr11.setRFonts(rfonts16);
        rfonts16.setAscii("Arial");
        rfonts16.setHAnsi("Arial");
        rfonts16.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue12 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr11.setB(booleandefaulttrue12);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue13 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr11.setBCs(booleandefaulttrue13);
        // Create object for color
        Color color16 = wmlObjectFactory.createColor();
        rpr11.setColor(color16);
        color16.setVal("000000");
        color16.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r12 = wmlObjectFactory.createR();
        p7.getContent().add(r12);
        // Create object for t (wrapped in JAXBElement)
        Text text12 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped12 = wmlObjectFactory.createRT(text12);
        r12.getContent().add(textWrapped12);
        text12.setValue(reviewData.getDate());
        // Create object for rPr
        RPr rpr12 = wmlObjectFactory.createRPr();
        r12.setRPr(rpr12);
        // Create object for rFonts
        RFonts rfonts17 = wmlObjectFactory.createRFonts();
        rpr12.setRFonts(rfonts17);
        rfonts17.setAscii("Arial");
        rfonts17.setHAnsi("Arial");
        rfonts17.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue14 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr12.setBCs(booleandefaulttrue14);
        // Create object for color
        Color color17 = wmlObjectFactory.createColor();
        rpr12.setColor(color17);
        color17.setVal("1F497D");
        color17.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for pPr
        PPr ppr7 = wmlObjectFactory.createPPr();
        p7.setPPr(ppr7);
        // Create object for rPr
        ParaRPr pararpr7 = wmlObjectFactory.createParaRPr();
        ppr7.setRPr(pararpr7);
        ppr7.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts18 = wmlObjectFactory.createRFonts();
        pararpr7.setRFonts(rfonts18);
        rfonts18.setAscii("Arial");
        rfonts18.setHAnsi("Arial");
        rfonts18.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue15 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr7.setBCs(booleandefaulttrue15);
        // Create object for color
        Color color18 = wmlObjectFactory.createColor();
        pararpr7.setColor(color18);
        color18.setVal("000000");
        color18.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p8 = wmlObjectFactory.createP();
        tc2.getContent().add(p8);
        // Create object for pPr
        PPr ppr8 = wmlObjectFactory.createPPr();
        p8.setPPr(ppr8);
        // Create object for rPr
        ParaRPr pararpr8 = wmlObjectFactory.createParaRPr();
        ppr8.setRPr(pararpr8);
        ppr8.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts19 = wmlObjectFactory.createRFonts();
        pararpr8.setRFonts(rfonts19);
        rfonts19.setAscii("Arial");
        rfonts19.setHAnsi("Arial");
        rfonts19.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue16 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr8.setBCs(booleandefaulttrue16);
        // Create object for color
        Color color19 = wmlObjectFactory.createColor();
        pararpr8.setColor(color19);
        color19.setVal("000000");
        color19.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p9 = wmlObjectFactory.createP();
        tc2.getContent().add(p9);
        // Create object for r
        R r13 = wmlObjectFactory.createR();
        p9.getContent().add(r13);
        // Create object for t (wrapped in JAXBElement)
        Text text13 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped13 = wmlObjectFactory.createRT(text13);
        r13.getContent().add(textWrapped13);
        text13.setValue("Project Address:");
        // Create object for rPr
        RPr rpr13 = wmlObjectFactory.createRPr();
        r13.setRPr(rpr13);
        // Create object for rFonts
        RFonts rfonts20 = wmlObjectFactory.createRFonts();
        rpr13.setRFonts(rfonts20);
        rfonts20.setAscii("Arial");
        rfonts20.setHAnsi("Arial");
        rfonts20.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue17 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr13.setB(booleandefaulttrue17);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue18 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr13.setBCs(booleandefaulttrue18);
        // Create object for color
        Color color20 = wmlObjectFactory.createColor();
        rpr13.setColor(color20);
        color20.setVal("000000");
        color20.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r14 = wmlObjectFactory.createR();
        p9.getContent().add(r14);
        // Create object for t (wrapped in JAXBElement)
        Text text14 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped14 = wmlObjectFactory.createRT(text14);
        r14.getContent().add(textWrapped14);
        text14.setValue(" ");
        text14.setSpace("preserve");
        // Create object for rPr
        RPr rpr14 = wmlObjectFactory.createRPr();
        r14.setRPr(rpr14);
        // Create object for rFonts
        RFonts rfonts21 = wmlObjectFactory.createRFonts();
        rpr14.setRFonts(rfonts21);
        rfonts21.setAscii("Arial");
        rfonts21.setHAnsi("Arial");
        rfonts21.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue19 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr14.setB(booleandefaulttrue19);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue20 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr14.setBCs(booleandefaulttrue20);
        // Create object for color
        Color color21 = wmlObjectFactory.createColor();
        rpr14.setColor(color21);
        color21.setVal("000000");
        color21.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r15 = wmlObjectFactory.createR();
        p9.getContent().add(r15);
        // Create object for t (wrapped in JAXBElement)
        Text text15 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped15 = wmlObjectFactory.createRT(text15);
        r15.getContent().add(textWrapped15);
        //TODO: Add variable Address
        text15.setValue(reviewData.getProjectAddress());
        // Create object for rPr
        RPr rpr15 = wmlObjectFactory.createRPr();
        r15.setRPr(rpr15);
        // Create object for rFonts
        RFonts rfonts22 = wmlObjectFactory.createRFonts();
        rpr15.setRFonts(rfonts22);
        rfonts22.setAscii("Arial");
        rfonts22.setHAnsi("Arial");
        rfonts22.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue21 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr15.setBCs(booleandefaulttrue21);
        // Create object for color
        Color color22 = wmlObjectFactory.createColor();
        rpr15.setColor(color22);
        color22.setVal("1F497D");
        color22.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for pPr
        PPr ppr9 = wmlObjectFactory.createPPr();
        p9.setPPr(ppr9);
        // Create object for rPr
        ParaRPr pararpr9 = wmlObjectFactory.createParaRPr();
        ppr9.setRPr(pararpr9);
        ppr9.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts23 = wmlObjectFactory.createRFonts();
        pararpr9.setRFonts(rfonts23);
        rfonts23.setAscii("Arial");
        rfonts23.setHAnsi("Arial");
        rfonts23.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue22 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr9.setBCs(booleandefaulttrue22);
        // Create object for color
        Color color23 = wmlObjectFactory.createColor();
        pararpr9.setColor(color23);
        color23.setVal("000000");
        color23.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p10 = wmlObjectFactory.createP();
        tc2.getContent().add(p10);
        // Create object for pPr
        PPr ppr10 = wmlObjectFactory.createPPr();
        p10.setPPr(ppr10);
        // Create object for rPr
        ParaRPr pararpr10 = wmlObjectFactory.createParaRPr();
        ppr10.setRPr(pararpr10);
        ppr10.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts24 = wmlObjectFactory.createRFonts();
        pararpr10.setRFonts(rfonts24);
        rfonts24.setAscii("Arial");
        rfonts24.setHAnsi("Arial");
        rfonts24.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue23 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr10.setBCs(booleandefaulttrue23);
        // Create object for color
        Color color24 = wmlObjectFactory.createColor();
        pararpr10.setColor(color24);
        color24.setVal("000000");
        color24.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p11 = wmlObjectFactory.createP();
        tc2.getContent().add(p11);
        // Create object for r
        R r16 = wmlObjectFactory.createR();
        p11.getContent().add(r16);
        // Create object for t (wrapped in JAXBElement)
        Text text16 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped16 = wmlObjectFactory.createRT(text16);
        r16.getContent().add(textWrapped16);
        //TODO: Add location variability
        addToP(p11, reviewData.getLocation()); //TextToAdd will be replaced with variable 
        text16.setValue("Location:");
        // Create object for rPr
        RPr rpr16 = wmlObjectFactory.createRPr();
        r16.setRPr(rpr16);
        // Create object for rFonts
        RFonts rfonts25 = wmlObjectFactory.createRFonts();
        rpr16.setRFonts(rfonts25);
        rfonts25.setAscii("Arial");
        rfonts25.setHAnsi("Arial");
        rfonts25.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue24 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr16.setB(booleandefaulttrue24);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue25 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr16.setBCs(booleandefaulttrue25);
        // Create object for color
        Color color25 = wmlObjectFactory.createColor();
        rpr16.setColor(color25);
        color25.setVal("000000");
        color25.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pPr
        PPr ppr11 = wmlObjectFactory.createPPr();
        p11.setPPr(ppr11);
        // Create object for rPr
        ParaRPr pararpr11 = wmlObjectFactory.createParaRPr();
        ppr11.setRPr(pararpr11);
        ppr11.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts26 = wmlObjectFactory.createRFonts();
        pararpr11.setRFonts(rfonts26);
        rfonts26.setAscii("Arial");
        rfonts26.setHAnsi("Arial");
        rfonts26.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue26 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr11.setB(booleandefaulttrue26);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue27 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr11.setBCs(booleandefaulttrue27);
        // Create object for color
        Color color26 = wmlObjectFactory.createColor();
        pararpr11.setColor(color26);
        color26.setVal("000000");
        color26.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for ind
        PPrBase.Ind pprbaseind = wmlObjectFactory.createPPrBaseInd();
        ppr11.setInd(pprbaseind);
        pprbaseind.setRight(BigInteger.valueOf(575));
        // Create object for p
        P p12 = wmlObjectFactory.createP();
        tc2.getContent().add(p12);
        // Create object for pPr
        PPr ppr12 = wmlObjectFactory.createPPr();
        p12.setPPr(ppr12);
        // Create object for rPr
        ParaRPr pararpr12 = wmlObjectFactory.createParaRPr();
        ppr12.setRPr(pararpr12);
        ppr12.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts27 = wmlObjectFactory.createRFonts();
        pararpr12.setRFonts(rfonts27);
        rfonts27.setAscii("Arial");
        rfonts27.setHAnsi("Arial");
        rfonts27.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue28 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr12.setBCs(booleandefaulttrue28);
        // Create object for color
        Color color27 = wmlObjectFactory.createColor();
        pararpr12.setColor(color27);
        color27.setVal("000000");
        color27.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for ind
        PPrBase.Ind pprbaseind2 = wmlObjectFactory.createPPrBaseInd();
        ppr12.setInd(pprbaseind2);
        pprbaseind2.setRight(BigInteger.valueOf(575));
        // Create object for p
        P p13 = wmlObjectFactory.createP();
        tc2.getContent().add(p13);
        // Create object for r
        R r17 = wmlObjectFactory.createR();
        p13.getContent().add(r17);
        // Create object for t (wrapped in JAXBElement)
        Text text17 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped17 = wmlObjectFactory.createRT(text17);
        r17.getContent().add(textWrapped17);
        //TODO: Add Reference Dwgs variability
        addToP(p13, reviewData.getReferenceDwgs());
        text17.setValue("Reference Dwgs:");
        // Create object for rPr
        RPr rpr17 = wmlObjectFactory.createRPr();
        r17.setRPr(rpr17);
        // Create object for rFonts
        RFonts rfonts28 = wmlObjectFactory.createRFonts();
        rpr17.setRFonts(rfonts28);
        rfonts28.setAscii("Arial");
        rfonts28.setHAnsi("Arial");
        rfonts28.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue29 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr17.setB(booleandefaulttrue29);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue30 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr17.setBCs(booleandefaulttrue30);
        // Create object for color
        Color color28 = wmlObjectFactory.createColor();
        rpr17.setColor(color28);
        color28.setVal("000000");
        color28.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pPr
        PPr ppr13 = wmlObjectFactory.createPPr();
        p13.setPPr(ppr13);
        // Create object for rPr
        ParaRPr pararpr13 = wmlObjectFactory.createParaRPr();
        ppr13.setRPr(pararpr13);
        ppr13.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts29 = wmlObjectFactory.createRFonts();
        pararpr13.setRFonts(rfonts29);
        rfonts29.setAscii("Arial");
        rfonts29.setHAnsi("Arial");
        rfonts29.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue31 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr13.setBCs(booleandefaulttrue31);
        // Create object for color
        Color color29 = wmlObjectFactory.createColor();
        pararpr13.setColor(color29);
        color29.setVal("000000");
        color29.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for ind
        PPrBase.Ind pprbaseind3 = wmlObjectFactory.createPPrBaseInd();
        ppr13.setInd(pprbaseind3);
        pprbaseind3.setRight(BigInteger.valueOf(575));
        // Create object for p
        P p14 = wmlObjectFactory.createP();
        tc2.getContent().add(p14);
        // Create object for pPr
        PPr ppr14 = wmlObjectFactory.createPPr();
        p14.setPPr(ppr14);
        // Create object for rPr
        ParaRPr pararpr14 = wmlObjectFactory.createParaRPr();
        ppr14.setRPr(pararpr14);
        ppr14.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts30 = wmlObjectFactory.createRFonts();
        pararpr14.setRFonts(rfonts30);
        rfonts30.setAscii("Arial");
        rfonts30.setHAnsi("Arial");
        rfonts30.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue32 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr14.setBCs(booleandefaulttrue32);
        // Create object for color
        Color color30 = wmlObjectFactory.createColor();
        pararpr14.setColor(color30);
        color30.setVal("000000");
        color30.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for ind
        PPrBase.Ind pprbaseind4 = wmlObjectFactory.createPPrBaseInd();
        ppr14.setInd(pprbaseind4);
        pprbaseind4.setRight(BigInteger.valueOf(575));
        // Create object for tcPr
        TcPr tcpr2 = wmlObjectFactory.createTcPr();
        tc2.setTcPr(tcpr2);
        // Create object for tcW
        TblWidth tblwidth2 = wmlObjectFactory.createTblWidth();
        tcpr2.setTcW(tblwidth2);
        tblwidth2.setW(BigInteger.valueOf(4260));
        tblwidth2.setType("dxa");
        // Create object for tcBorders
        TcPrInner.TcBorders tcprinnertcborders2 = wmlObjectFactory.createTcPrInnerTcBorders();
        tcpr2.setTcBorders(tcprinnertcborders2);
        // Create object for top
        CTBorder border4 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders2.setTop(border4);
        border4.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border4.setColor("auto");
        border4.setSz(BigInteger.valueOf(4));
        border4.setSpace(BigInteger.valueOf(0));
        // Create object for left
        CTBorder border5 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders2.setLeft(border5);
        border5.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border5.setColor("auto");
        border5.setSz(BigInteger.valueOf(4));
        border5.setSpace(BigInteger.valueOf(0));
        // Create object for bottom
        CTBorder border6 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders2.setBottom(border6);
        border6.setVal(org.docx4j.wml.STBorder.TRIPLE);
        border6.setColor("auto");
        border6.setSz(BigInteger.valueOf(4));
        border6.setSpace(BigInteger.valueOf(0));
        // Create object for right
        CTBorder border7 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders2.setRight(border7);
        border7.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border7.setColor("auto");
        border7.setSz(BigInteger.valueOf(4));
        border7.setSpace(BigInteger.valueOf(0));
        // Create object for tc (wrapped in JAXBElement)
        Tc tc3 = wmlObjectFactory.createTc();
        JAXBElement<org.docx4j.wml.Tc> tcWrapped3 = wmlObjectFactory.createTrTc(tc3);
        tr2.getContent().add(tcWrapped3);
        // Create object for p
        P p15 = wmlObjectFactory.createP();
        tc3.getContent().add(p15);
        // Create object for pPr
        PPr ppr15 = wmlObjectFactory.createPPr();
        p15.setPPr(ppr15);
        // Create object for rPr
        ParaRPr pararpr15 = wmlObjectFactory.createParaRPr();
        ppr15.setRPr(pararpr15);
        ppr15.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts31 = wmlObjectFactory.createRFonts();
        pararpr15.setRFonts(rfonts31);
        rfonts31.setAscii("Arial");
        rfonts31.setHAnsi("Arial");
        rfonts31.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue33 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr15.setB(booleandefaulttrue33);
        // Create object for color
        Color color31 = wmlObjectFactory.createColor();
        pararpr15.setColor(color31);
        color31.setVal("000000");
        color31.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle6 = wmlObjectFactory.createPPrBasePStyle();
        ppr15.setPStyle(pprbasepstyle6);
        pprbasepstyle6.setVal("BodyText");
        // Create object for p
        P p16 = wmlObjectFactory.createP();
        tc3.getContent().add(p16);
        // Create object for r
        R r18 = wmlObjectFactory.createR();
        p16.getContent().add(r18);
        // Create object for t (wrapped in JAXBElement)
        Text text18 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped18 = wmlObjectFactory.createRT(text18);
        r18.getContent().add(textWrapped18);
        text18.setValue("Project:");
        // Create object for rPr
        RPr rpr18 = wmlObjectFactory.createRPr();
        r18.setRPr(rpr18);
        // Create object for rFonts
        RFonts rfonts32 = wmlObjectFactory.createRFonts();
        rpr18.setRFonts(rfonts32);
        rfonts32.setAscii("Arial");
        rfonts32.setHAnsi("Arial");
        rfonts32.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue34 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr18.setB(booleandefaulttrue34);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue35 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr18.setBCs(booleandefaulttrue35);
        // Create object for color
        Color color32 = wmlObjectFactory.createColor();
        rpr18.setColor(color32);
        color32.setVal("000000");
        color32.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r19 = wmlObjectFactory.createR();
        p16.getContent().add(r19);
        // Create object for t (wrapped in JAXBElement)
        Text text19 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped19 = wmlObjectFactory.createRT(text19);
        r19.getContent().add(textWrapped19);
        text19.setValue(" ");
        text19.setSpace("preserve");
        // Create object for rPr
        RPr rpr19 = wmlObjectFactory.createRPr();
        r19.setRPr(rpr19);
        // Create object for rFonts
        RFonts rfonts33 = wmlObjectFactory.createRFonts();
        rpr19.setRFonts(rfonts33);
        rfonts33.setAscii("Arial");
        rfonts33.setHAnsi("Arial");
        rfonts33.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue36 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr19.setB(booleandefaulttrue36);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue37 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr19.setBCs(booleandefaulttrue37);
        // Create object for color
        Color color33 = wmlObjectFactory.createColor();
        rpr19.setColor(color33);
        color33.setVal("000000");
        color33.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r20 = wmlObjectFactory.createR();
        p16.getContent().add(r20);
        // Create object for t (wrapped in JAXBElement)
        Text text20 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped20 = wmlObjectFactory.createRT(text20);
        r20.getContent().add(textWrapped20);
        //TODO: Add variable Project name
        text20.setValue(reviewData.getProjectName());
        // Create object for rPr
        RPr rpr20 = wmlObjectFactory.createRPr();
        r20.setRPr(rpr20);
        // Create object for rFonts
        RFonts rfonts34 = wmlObjectFactory.createRFonts();
        rpr20.setRFonts(rfonts34);
        rfonts34.setAscii("Arial");
        rfonts34.setHAnsi("Arial");
        rfonts34.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue38 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr20.setBCs(booleandefaulttrue38);
        // Create object for color
        Color color34 = wmlObjectFactory.createColor();
        rpr20.setColor(color34);
        color34.setVal("1F497D");
        color34.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for pPr
        PPr ppr16 = wmlObjectFactory.createPPr();
        p16.setPPr(ppr16);
        // Create object for rPr
        ParaRPr pararpr16 = wmlObjectFactory.createParaRPr();
        ppr16.setRPr(pararpr16);
        // Create object for rFonts
        RFonts rfonts35 = wmlObjectFactory.createRFonts();
        pararpr16.setRFonts(rfonts35);
        rfonts35.setAscii("Arial");
        rfonts35.setHAnsi("Arial");
        rfonts35.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue39 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr16.setBCs(booleandefaulttrue39);
        // Create object for color
        Color color35 = wmlObjectFactory.createColor();
        pararpr16.setColor(color35);
        color35.setVal("000000");
        color35.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p17 = wmlObjectFactory.createP();
        tc3.getContent().add(p17);
        // Create object for pPr
        PPr ppr17 = wmlObjectFactory.createPPr();
        p17.setPPr(ppr17);
        // Create object for rPr
        ParaRPr pararpr17 = wmlObjectFactory.createParaRPr();
        ppr17.setRPr(pararpr17);
        ppr17.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts36 = wmlObjectFactory.createRFonts();
        pararpr17.setRFonts(rfonts36);
        rfonts36.setAscii("Arial");
        rfonts36.setHAnsi("Arial");
        rfonts36.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue40 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr17.setBCs(booleandefaulttrue40);
        // Create object for color
        Color color36 = wmlObjectFactory.createColor();
        pararpr17.setColor(color36);
        color36.setVal("000000");
        color36.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p18 = wmlObjectFactory.createP();
        tc3.getContent().add(p18);
        // Create object for r
        R r21 = wmlObjectFactory.createR();
        p18.getContent().add(r21);
        // Create object for t (wrapped in JAXBElement)
        Text text21 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped21 = wmlObjectFactory.createRT(text21);
        r21.getContent().add(textWrapped21);
        text21.setValue("Builder:");
        // Create object for rPr
        RPr rpr21 = wmlObjectFactory.createRPr();
        r21.setRPr(rpr21);
        // Create object for rFonts
        RFonts rfonts37 = wmlObjectFactory.createRFonts();
        rpr21.setRFonts(rfonts37);
        rfonts37.setAscii("Arial");
        rfonts37.setHAnsi("Arial");
        rfonts37.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue41 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr21.setB(booleandefaulttrue41);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue42 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr21.setBCs(booleandefaulttrue42);
        // Create object for color
        Color color37 = wmlObjectFactory.createColor();
        rpr21.setColor(color37);
        color37.setVal("000000");
        color37.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r22 = wmlObjectFactory.createR();
        p18.getContent().add(r22);
        // Create object for t (wrapped in JAXBElement)
        Text text22 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped22 = wmlObjectFactory.createRT(text22);
        r22.getContent().add(textWrapped22);
        text22.setValue(" ");
        text22.setSpace("preserve");
        // Create object for rPr
        RPr rpr22 = wmlObjectFactory.createRPr();
        r22.setRPr(rpr22);
        // Create object for rFonts
        RFonts rfonts38 = wmlObjectFactory.createRFonts();
        rpr22.setRFonts(rfonts38);
        rfonts38.setAscii("Arial");
        rfonts38.setHAnsi("Arial");
        rfonts38.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue43 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr22.setB(booleandefaulttrue43);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue44 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr22.setBCs(booleandefaulttrue44);
        // Create object for color
        Color color38 = wmlObjectFactory.createColor();
        rpr22.setColor(color38);
        color38.setVal("000000");
        color38.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for r
        R r23 = wmlObjectFactory.createR();
        p18.getContent().add(r23);
        // Create object for t (wrapped in JAXBElement)
        Text text23 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped23 = wmlObjectFactory.createRT(text23);
        r23.getContent().add(textWrapped23);
        //TODO: Add variability to Builder name
        text23.setValue(reviewData.getBuilder());
        // Create object for rPr
        RPr rpr23 = wmlObjectFactory.createRPr();
        r23.setRPr(rpr23);
        // Create object for rFonts
        RFonts rfonts39 = wmlObjectFactory.createRFonts();
        rpr23.setRFonts(rfonts39);
        rfonts39.setAscii("Arial");
        rfonts39.setHAnsi("Arial");
        rfonts39.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue45 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr23.setBCs(booleandefaulttrue45);
        // Create object for color
        Color color39 = wmlObjectFactory.createColor();
        rpr23.setColor(color39);
        color39.setVal("1F497D");
        color39.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for pPr
        PPr ppr18 = wmlObjectFactory.createPPr();
        p18.setPPr(ppr18);
        // Create object for rPr
        ParaRPr pararpr18 = wmlObjectFactory.createParaRPr();
        ppr18.setRPr(pararpr18);
        ppr18.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts40 = wmlObjectFactory.createRFonts();
        pararpr18.setRFonts(rfonts40);
        rfonts40.setAscii("Arial");
        rfonts40.setHAnsi("Arial");
        rfonts40.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue46 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr18.setBCs(booleandefaulttrue46);
        // Create object for color
        Color color40 = wmlObjectFactory.createColor();
        pararpr18.setColor(color40);
        color40.setVal("000000");
        color40.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p19 = wmlObjectFactory.createP();
        tc3.getContent().add(p19);
        // Create object for pPr
        PPr ppr19 = wmlObjectFactory.createPPr();
        p19.setPPr(ppr19);
        // Create object for rPr
        ParaRPr pararpr19 = wmlObjectFactory.createParaRPr();
        ppr19.setRPr(pararpr19);
        ppr19.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts41 = wmlObjectFactory.createRFonts();
        pararpr19.setRFonts(rfonts41);
        rfonts41.setAscii("Arial");
        rfonts41.setHAnsi("Arial");
        rfonts41.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue47 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr19.setBCs(booleandefaulttrue47);
        // Create object for color
        Color color41 = wmlObjectFactory.createColor();
        pararpr19.setColor(color41);
        color41.setVal("000000");
        color41.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p20 = wmlObjectFactory.createP();
        tc3.getContent().add(p20);
        // Create object for r
        R r24 = wmlObjectFactory.createR();
        p20.getContent().add(r24);
        // Create object for t (wrapped in JAXBElement)
        Text text24 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped24 = wmlObjectFactory.createRT(text24);
        r24.getContent().add(textWrapped24);
        //TODO: Add variability to weather conditions
        addToP(p20, reviewData.getWeatherCondition()); //change textToAdd with variable
        text24.setValue("Weather Condition:");
        // Create object for rPr
        RPr rpr24 = wmlObjectFactory.createRPr();
        r24.setRPr(rpr24);
        // Create object for rFonts
        RFonts rfonts42 = wmlObjectFactory.createRFonts();
        rpr24.setRFonts(rfonts42);
        rfonts42.setAscii("Arial");
        rfonts42.setHAnsi("Arial");
        rfonts42.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue48 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr24.setB(booleandefaulttrue48);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue49 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr24.setBCs(booleandefaulttrue49);
        // Create object for color
        Color color42 = wmlObjectFactory.createColor();
        rpr24.setColor(color42);
        color42.setVal("000000");
        color42.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pPr
        PPr ppr20 = wmlObjectFactory.createPPr();
        p20.setPPr(ppr20);
        // Create object for rPr
        ParaRPr pararpr20 = wmlObjectFactory.createParaRPr();
        ppr20.setRPr(pararpr20);
        ppr20.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts43 = wmlObjectFactory.createRFonts();
        pararpr20.setRFonts(rfonts43);
        rfonts43.setAscii("Arial");
        rfonts43.setHAnsi("Arial");
        rfonts43.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue50 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr20.setBCs(booleandefaulttrue50);
        // Create object for color
        Color color43 = wmlObjectFactory.createColor();
        pararpr20.setColor(color43);
        color43.setVal("000000");
        color43.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p21 = wmlObjectFactory.createP();
        tc3.getContent().add(p21);
        // Create object for pPr
        PPr ppr21 = wmlObjectFactory.createPPr();
        p21.setPPr(ppr21);
        // Create object for rPr
        ParaRPr pararpr21 = wmlObjectFactory.createParaRPr();
        ppr21.setRPr(pararpr21);
        ppr21.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts44 = wmlObjectFactory.createRFonts();
        pararpr21.setRFonts(rfonts44);
        rfonts44.setAscii("Arial");
        rfonts44.setHAnsi("Arial");
        rfonts44.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue51 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr21.setBCs(booleandefaulttrue51);
        // Create object for color
        Color color44 = wmlObjectFactory.createColor();
        pararpr21.setColor(color44);
        color44.setVal("000000");
        color44.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p22 = wmlObjectFactory.createP();
        tc3.getContent().add(p22);
        // Create object for r
        R r25 = wmlObjectFactory.createR();
        p22.getContent().add(r25);
        // Create object for t (wrapped in JAXBElement)
        Text text25 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped25 = wmlObjectFactory.createRT(text25);
        r25.getContent().add(textWrapped25);
        text25.setValue("Inspection Category: ");
        text25.setSpace("preserve");
        // Create object for rPr
        RPr rpr25 = wmlObjectFactory.createRPr();
        r25.setRPr(rpr25);
        // Create object for rFonts
        RFonts rfonts45 = wmlObjectFactory.createRFonts();
        rpr25.setRFonts(rfonts45);
        rfonts45.setAscii("Arial");
        rfonts45.setHAnsi("Arial");
        rfonts45.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue52 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr25.setB(booleandefaulttrue52);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue53 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr25.setBCs(booleandefaulttrue53);
        // Create object for color
        Color color45 = wmlObjectFactory.createColor();
        rpr25.setColor(color45);
        color45.setVal("000000");
        color45.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for pPr
        PPr ppr22 = wmlObjectFactory.createPPr();
        p22.setPPr(ppr22);
        // Create object for rPr
        ParaRPr pararpr22 = wmlObjectFactory.createParaRPr();
        ppr22.setRPr(pararpr22);
        ppr22.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts46 = wmlObjectFactory.createRFonts();
        pararpr22.setRFonts(rfonts46);
        rfonts46.setAscii("Arial");
        rfonts46.setHAnsi("Arial");
        rfonts46.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue54 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr22.setB(booleandefaulttrue54);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue55 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr22.setBCs(booleandefaulttrue55);
        // Create object for color
        Color color46 = wmlObjectFactory.createColor();
        pararpr22.setColor(color46);
        color46.setVal("000000");
        color46.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_1);
        // Create object for p
        P p23 = wmlObjectFactory.createP();
        tc3.getContent().add(p23);
        // Create object for r
        R r26 = wmlObjectFactory.createR();
        p23.getContent().add(r26);
        // Create object for t (wrapped in JAXBElement)
        Text text26 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped26 = wmlObjectFactory.createRT(text26);
        r26.getContent().add(textWrapped26);
        //TODO: Add variability to Inspection Category
        text26.setValue(reviewData.getInspectionCategory());
        // Create object for rPr
        RPr rpr26 = wmlObjectFactory.createRPr();
        r26.setRPr(rpr26);
        // Create object for rFonts
        RFonts rfonts47 = wmlObjectFactory.createRFonts();
        rpr26.setRFonts(rfonts47);
        rfonts47.setAscii("Arial");
        rfonts47.setHAnsi("Arial");
        rfonts47.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue56 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr26.setBCs(booleandefaulttrue56);
        // Create object for color
        Color color47 = wmlObjectFactory.createColor();
        rpr26.setColor(color47);
        color47.setVal("365F91");
        color47.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color47.setThemeShade("BF");
        // Create object for pPr
        PPr ppr23 = wmlObjectFactory.createPPr();
        p23.setPPr(ppr23);
        // Create object for rPr
        ParaRPr pararpr23 = wmlObjectFactory.createParaRPr();
        ppr23.setRPr(pararpr23);
        ppr23.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts48 = wmlObjectFactory.createRFonts();
        pararpr23.setRFonts(rfonts48);
        rfonts48.setAscii("Arial");
        rfonts48.setHAnsi("Arial");
        rfonts48.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue57 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr23.setBCs(booleandefaulttrue57);
        // Create object for color
        Color color48 = wmlObjectFactory.createColor();
        pararpr23.setColor(color48);
        color48.setVal("365F91");
        color48.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color48.setThemeShade("BF");
        // Create object for tcPr
        TcPr tcpr3 = wmlObjectFactory.createTcPr();
        tc3.setTcPr(tcpr3);
        // Create object for tcW
        TblWidth tblwidth3 = wmlObjectFactory.createTblWidth();
        tcpr3.setTcW(tblwidth3);
        tblwidth3.setW(BigInteger.valueOf(3939));
        tblwidth3.setType("dxa");
        // Create object for tcBorders
        TcPrInner.TcBorders tcprinnertcborders3 = wmlObjectFactory.createTcPrInnerTcBorders();
        tcpr3.setTcBorders(tcprinnertcborders3);
        // Create object for top
        CTBorder border8 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders3.setTop(border8);
        border8.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border8.setColor("auto");
        border8.setSz(BigInteger.valueOf(4));
        border8.setSpace(BigInteger.valueOf(0));
        // Create object for left
        CTBorder border9 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders3.setLeft(border9);
        border9.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border9.setColor("auto");
        border9.setSz(BigInteger.valueOf(4));
        border9.setSpace(BigInteger.valueOf(0));
        // Create object for bottom
        CTBorder border10 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders3.setBottom(border10);
        border10.setVal(org.docx4j.wml.STBorder.TRIPLE);
        border10.setColor("auto");
        border10.setSz(BigInteger.valueOf(4));
        border10.setSpace(BigInteger.valueOf(0));
        // Create object for right
        CTBorder border11 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders3.setRight(border11);
        border11.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border11.setColor("auto");
        border11.setSz(BigInteger.valueOf(4));
        border11.setSpace(BigInteger.valueOf(0));
        // Create object for trPr
        TrPr trpr2 = wmlObjectFactory.createTrPr();
        tr2.setTrPr(trpr2);
        // Create object for trHeight (wrapped in JAXBElement)
        CTHeight height2 = wmlObjectFactory.createCTHeight();
        JAXBElement<org.docx4j.wml.CTHeight> heightWrapped2 = wmlObjectFactory.createCTTrPrBaseTrHeight(height2);
        trpr2.getCnfStyleOrDivIdOrGridBefore().add(heightWrapped2);
        height2.setVal(BigInteger.valueOf(2870));
        // Create object for tr
        Tr tr3 = wmlObjectFactory.createTr();
        tbl.getContent().add(tr3);
        // Create object for tc (wrapped in JAXBElement)
        Tc tc4 = wmlObjectFactory.createTc();
        JAXBElement<org.docx4j.wml.Tc> tcWrapped4 = wmlObjectFactory.createTrTc(tc4);
        tr3.getContent().add(tcWrapped4);
        // Create object for p
        P p24 = wmlObjectFactory.createP();
        tc4.getContent().add(p24);
        // Create object for pPr
        PPr ppr24 = wmlObjectFactory.createPPr();
        p24.setPPr(ppr24);
        // Create object for rPr
        ParaRPr pararpr24 = wmlObjectFactory.createParaRPr();
        ppr24.setRPr(pararpr24);
        ppr24.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts49 = wmlObjectFactory.createRFonts();
        pararpr24.setRFonts(rfonts49);
        rfonts49.setAscii("Arial");
        rfonts49.setHAnsi("Arial");
        rfonts49.setCs("Arial");
        // Create object for color
        Color color49 = wmlObjectFactory.createColor();
        pararpr24.setColor(color49);
        color49.setVal("365F91");
        color49.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color49.setThemeShade("BF");
        // Create object for p
        P p25 = wmlObjectFactory.createP();
        tc4.getContent().add(p25);
        // Create object for r
        R r27 = wmlObjectFactory.createR();
        p25.getContent().add(r27);
        // Create object for t (wrapped in JAXBElement)
        Text text27 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped27 = wmlObjectFactory.createRT(text27);
        r27.getContent().add(textWrapped27);
        text27.setValue("We have visited the above noted site on");
        // Create object for rPr
        RPr rpr27 = wmlObjectFactory.createRPr();
        r27.setRPr(rpr27);
        // Create object for rFonts
        RFonts rfonts50 = wmlObjectFactory.createRFonts();
        rpr27.setRFonts(rfonts50);
        rfonts50.setAscii("Arial");
        rfonts50.setHAnsi("Arial");
        rfonts50.setCs("Arial");
        // Create object for color
        Color color50 = wmlObjectFactory.createColor();
        rpr27.setColor(color50);
        color50.setVal("365F91");
        color50.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color50.setThemeShade("BF");
        // Create object for r
        R r28 = wmlObjectFactory.createR();
        p25.getContent().add(r28);
        // Create object for t (wrapped in JAXBElement)
        Text text28 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped28 = wmlObjectFactory.createRT(text28);
        r28.getContent().add(textWrapped28);
        //TODO: Add date list functionality
        text28.setValue(" July 13 and August 11, 2020 ");
        text28.setSpace("preserve");
        // Create object for rPr
        RPr rpr28 = wmlObjectFactory.createRPr();
        r28.setRPr(rpr28);
        // Create object for rFonts
        RFonts rfonts51 = wmlObjectFactory.createRFonts();
        rpr28.setRFonts(rfonts51);
        rfonts51.setAscii("Arial");
        rfonts51.setHAnsi("Arial");
        rfonts51.setCs("Arial");
        // Create object for color
        Color color51 = wmlObjectFactory.createColor();
        rpr28.setColor(color51);
        color51.setVal("365F91");
        color51.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color51.setThemeShade("BF");
        // Create object for r
        R r29 = wmlObjectFactory.createR();
        p25.getContent().add(r29);
        // Create object for t (wrapped in JAXBElement)
        Text text29 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped29 = wmlObjectFactory.createRT(text29);
        r29.getContent().add(textWrapped29);
        text29.setValue("and our observations were as follow");
        // Create object for rPr
        RPr rpr29 = wmlObjectFactory.createRPr();
        r29.setRPr(rpr29);
        // Create object for rFonts
        RFonts rfonts52 = wmlObjectFactory.createRFonts();
        rpr29.setRFonts(rfonts52);
        rfonts52.setAscii("Arial");
        rfonts52.setHAnsi("Arial");
        rfonts52.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue58 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr29.setBCs(booleandefaulttrue58);
        // Create object for color
        Color color52 = wmlObjectFactory.createColor();
        rpr29.setColor(color52);
        color52.setVal("365F91");
        color52.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color52.setThemeShade("BF");
        // Create object for r
        R r30 = wmlObjectFactory.createR();
        p25.getContent().add(r30);
        // Create object for t (wrapped in JAXBElement)
        Text text30 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped30 = wmlObjectFactory.createRT(text30);
        r30.getContent().add(textWrapped30);
        text30.setValue(":");
        // Create object for rPr
        RPr rpr30 = wmlObjectFactory.createRPr();
        r30.setRPr(rpr30);
        // Create object for rFonts
        RFonts rfonts53 = wmlObjectFactory.createRFonts();
        rpr30.setRFonts(rfonts53);
        rfonts53.setAscii("Arial");
        rfonts53.setHAnsi("Arial");
        rfonts53.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue59 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr30.setBCs(booleandefaulttrue59);
        // Create object for color
        Color color53 = wmlObjectFactory.createColor();
        rpr30.setColor(color53);
        color53.setVal("365F91");
        color53.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color53.setThemeShade("BF");
        // Create object for pPr
        PPr ppr25 = wmlObjectFactory.createPPr();
        p25.setPPr(ppr25);
        // Create object for rPr
        ParaRPr pararpr25 = wmlObjectFactory.createParaRPr();
        ppr25.setRPr(pararpr25);
        ppr25.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts54 = wmlObjectFactory.createRFonts();
        pararpr25.setRFonts(rfonts54);
        rfonts54.setAscii("Arial");
        rfonts54.setHAnsi("Arial");
        rfonts54.setCs("Arial");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue60 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr25.setBCs(booleandefaulttrue60);
        // Create object for color
        Color color54 = wmlObjectFactory.createColor();
        pararpr25.setColor(color54);
        color54.setVal("365F91");
        color54.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color54.setThemeShade("BF");
        // Create object for p
        P p26 = wmlObjectFactory.createP();
        tc4.getContent().add(p26);
        // Create object for pPr
        PPr ppr26 = wmlObjectFactory.createPPr();
        p26.setPPr(ppr26);
        // Create object for rPr
        ParaRPr pararpr26 = wmlObjectFactory.createParaRPr();
        ppr26.setRPr(pararpr26);
        ppr26.setSpacing(pprBaseSpacing);
        // Create object for rFonts
        RFonts rfonts55 = wmlObjectFactory.createRFonts();
        pararpr26.setRFonts(rfonts55);
        rfonts55.setAscii("Arial");
        rfonts55.setHAnsi("Arial");
        rfonts55.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue61 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr26.setB(booleandefaulttrue61);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue62 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr26.setBCs(booleandefaulttrue62);
        // Create object for color
        Color color55 = wmlObjectFactory.createColor();
        pararpr26.setColor(color55);
        color55.setVal("365F91");
        color55.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color55.setThemeShade("BF");
        // Create object for p
        addNumberedList(reviewData.getNumberedList(), tc4);
        // Create object for p
        P p46 = wmlObjectFactory.createP();
        tc4.getContent().add(p46);
        // Create object for r
        R r45 = wmlObjectFactory.createR();
        p46.getContent().add(r45);
        // Create object for t (wrapped in JAXBElement)
        Text text45 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped45 = wmlObjectFactory.createRT(text45);
        r45.getContent().add(textWrapped45);
        text45.setValue(" ");
        text45.setSpace("preserve");
        // Create object for rPr
        RPr rpr45 = wmlObjectFactory.createRPr();
        r45.setRPr(rpr45);
        // Create object for rFonts
        RFonts rfonts89 = wmlObjectFactory.createRFonts();
        rpr45.setRFonts(rfonts89);
        rfonts89.setAscii("Arial");
        rfonts89.setHAnsi("Arial");
        rfonts89.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue103 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr45.setB(booleandefaulttrue103);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue104 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr45.setBCs(booleandefaulttrue104);
        // Create object for color
        Color color89 = wmlObjectFactory.createColor();
        rpr45.setColor(color89);
        color89.setVal("365F91");
        color89.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color89.setThemeShade("BF");
        // Create object for pPr
        PPr ppr46 = wmlObjectFactory.createPPr();
        p46.setPPr(ppr46);
        // Create object for rPr
        ParaRPr pararpr46 = wmlObjectFactory.createParaRPr();
        ppr46.setRPr(pararpr46);
        // Create object for rFonts
        RFonts rfonts90 = wmlObjectFactory.createRFonts();
        pararpr46.setRFonts(rfonts90);
        rfonts90.setAscii("Arial");
        rfonts90.setHAnsi("Arial");
        rfonts90.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue105 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr46.setB(booleandefaulttrue105);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue106 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr46.setBCs(booleandefaulttrue106);
        // Create object for color
        Color color90 = wmlObjectFactory.createColor();
        pararpr46.setColor(color90);
        color90.setVal("365F91");
        color90.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color90.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle23 = wmlObjectFactory.createPPrBasePStyle();
        ppr46.setPStyle(pprbasepstyle23);
        pprbasepstyle23.setVal("ListParagraph");
        // Create object for p
        P p47 = wmlObjectFactory.createP();
        tc4.getContent().add(p47);
        // Create object for r
        R r46 = wmlObjectFactory.createR();
        p47.getContent().add(r46);
        // Create object for t (wrapped in JAXBElement)
        Text text46 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped46 = wmlObjectFactory.createRT(text46);
        r46.getContent().add(textWrapped46);
        text46.setValue(
                "Work was in progress and was reviewed on a random sampling basis and was found to be in general conformance with Contract Documents unless otherwise ");
        text46.setSpace("preserve");
        // Create object for rPr
        RPr rpr46 = wmlObjectFactory.createRPr();
        r46.setRPr(rpr46);
        // Create object for rFonts
        RFonts rfonts91 = wmlObjectFactory.createRFonts();
        rpr46.setRFonts(rfonts91);
        rfonts91.setAscii("Arial");
        rfonts91.setHAnsi("Arial");
        rfonts91.setCs("Arial");
        // Create object for color
        Color color91 = wmlObjectFactory.createColor();
        rpr46.setColor(color91);
        color91.setVal("365F91");
        color91.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color91.setThemeShade("BF");
        // Create object for r
        R r47 = wmlObjectFactory.createR();
        p47.getContent().add(r47);
        // Create object for t (wrapped in JAXBElement)
        Text text47 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped47 = wmlObjectFactory.createRT(text47);
        r47.getContent().add(textWrapped47);
        text47.setValue("noted above");
        // Create object for rPr
        RPr rpr47 = wmlObjectFactory.createRPr();
        r47.setRPr(rpr47);
        // Create object for rFonts
        RFonts rfonts92 = wmlObjectFactory.createRFonts();
        rpr47.setRFonts(rfonts92);
        rfonts92.setAscii("Arial");
        rfonts92.setHAnsi("Arial");
        rfonts92.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue107 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr47.setB(booleandefaulttrue107);
        // Create object for i
        BooleanDefaultTrue booleandefaulttrue108 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr47.setI(booleandefaulttrue108);
        // Create object for color
        Color color92 = wmlObjectFactory.createColor();
        rpr47.setColor(color92);
        color92.setVal("365F91");
        color92.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color92.setThemeShade("BF");
        // Create object for r
        R r48 = wmlObjectFactory.createR();
        p47.getContent().add(r48);
        // Create object for t (wrapped in JAXBElement)
        Text text48 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped48 = wmlObjectFactory.createRT(text48);
        r48.getContent().add(textWrapped48);
        text48.setValue(
                ". Site was clean and free of debris. Photos index enclosed is an essential part of our field Report.");
        // Create object for rPr
        RPr rpr48 = wmlObjectFactory.createRPr();
        r48.setRPr(rpr48);
        // Create object for rFonts
        RFonts rfonts93 = wmlObjectFactory.createRFonts();
        rpr48.setRFonts(rfonts93);
        rfonts93.setAscii("Arial");
        rfonts93.setHAnsi("Arial");
        rfonts93.setCs("Arial");
        // Create object for color
        Color color93 = wmlObjectFactory.createColor();
        rpr48.setColor(color93);
        color93.setVal("365F91");
        color93.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color93.setThemeShade("BF");
        // Create object for pPr
        PPr ppr47 = wmlObjectFactory.createPPr();
        p47.setPPr(ppr47);
        // Create object for rPr
        ParaRPr pararpr47 = wmlObjectFactory.createParaRPr();
        ppr47.setRPr(pararpr47);
        // Create object for rFonts
        RFonts rfonts94 = wmlObjectFactory.createRFonts();
        pararpr47.setRFonts(rfonts94);
        rfonts94.setAscii("Arial");
        rfonts94.setHAnsi("Arial");
        rfonts94.setCs("Arial");
        // Create object for color
        Color color94 = wmlObjectFactory.createColor();
        pararpr47.setColor(color94);
        color94.setVal("365F91");
        color94.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color94.setThemeShade("BF");
        // Create object for jc
        Jc jc = wmlObjectFactory.createJc();
        ppr47.setJc(jc);
        jc.setVal(org.docx4j.wml.JcEnumeration.BOTH);
        // Create object for tcPr
        TcPr tcpr4 = wmlObjectFactory.createTcPr();
        tc4.setTcPr(tcpr4);
        // Create object for tcW
        TblWidth tblwidth4 = wmlObjectFactory.createTblWidth();
        tcpr4.setTcW(tblwidth4);
        tblwidth4.setW(BigInteger.valueOf(8199));
        tblwidth4.setType("dxa");
        // Create object for gridSpan
        TcPrInner.GridSpan tcprinnergridspan2 = wmlObjectFactory.createTcPrInnerGridSpan();
        tcpr4.setGridSpan(tcprinnergridspan2);
        tcprinnergridspan2.setVal(BigInteger.valueOf(2));
        // Create object for tcBorders
        TcPrInner.TcBorders tcprinnertcborders4 = wmlObjectFactory.createTcPrInnerTcBorders();
        tcpr4.setTcBorders(tcprinnertcborders4);
        // Create object for top
        CTBorder border12 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders4.setTop(border12);
        border12.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border12.setColor("auto");
        border12.setSz(BigInteger.valueOf(4));
        border12.setSpace(BigInteger.valueOf(0));
        // Create object for left
        CTBorder border13 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders4.setLeft(border13);
        border13.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border13.setColor("auto");
        border13.setSz(BigInteger.valueOf(4));
        border13.setSpace(BigInteger.valueOf(0));
        // Create object for bottom
        CTBorder border14 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders4.setBottom(border14);
        border14.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border14.setColor("auto");
        border14.setSz(BigInteger.valueOf(4));
        border14.setSpace(BigInteger.valueOf(0));
        // Create object for right
        CTBorder border15 = wmlObjectFactory.createCTBorder();
        tcprinnertcborders4.setRight(border15);
        border15.setVal(org.docx4j.wml.STBorder.DOUBLE);
        border15.setColor("auto");
        border15.setSz(BigInteger.valueOf(4));
        border15.setSpace(BigInteger.valueOf(0));
        // Create object for trPr
        TrPr trpr3 = wmlObjectFactory.createTrPr();
        tr3.setTrPr(trpr3);
        // Create object for trHeight (wrapped in JAXBElement)
        CTHeight height3 = wmlObjectFactory.createCTHeight();
        JAXBElement<org.docx4j.wml.CTHeight> heightWrapped3 = wmlObjectFactory.createCTTrPrBaseTrHeight(height3);
        trpr3.getCnfStyleOrDivIdOrGridBefore().add(heightWrapped3);
        height3.setVal(BigInteger.valueOf(1180));
        // Create object for tblPr
        TblPr tblpr = wmlObjectFactory.createTblPr();
        tbl.setTblPr(tblpr);
        // Create object for tblW
        TblWidth tblwidth5 = wmlObjectFactory.createTblWidth();
        tblpr.setTblW(tblwidth5);
        tblwidth5.setW(BigInteger.valueOf(0));
        tblwidth5.setType("auto");
        // Create object for tblInd
        TblWidth tblwidth6 = wmlObjectFactory.createTblWidth();
        tblpr.setTblInd(tblwidth6);
        tblwidth6.setW(BigInteger.valueOf(-432));
        tblwidth6.setType("dxa");
        // Create object for tblBorders
        TblBorders tblborders = wmlObjectFactory.createTblBorders();
        tblpr.setTblBorders(tblborders);
        // Create object for top
        CTBorder border16 = wmlObjectFactory.createCTBorder();
        tblborders.setTop(border16);
        border16.setVal(org.docx4j.wml.STBorder.SINGLE);
        border16.setColor("auto");
        border16.setSz(BigInteger.valueOf(4));
        border16.setSpace(BigInteger.valueOf(0));
        // Create object for left
        CTBorder border17 = wmlObjectFactory.createCTBorder();
        tblborders.setLeft(border17);
        border17.setVal(org.docx4j.wml.STBorder.SINGLE);
        border17.setColor("auto");
        border17.setSz(BigInteger.valueOf(4));
        border17.setSpace(BigInteger.valueOf(0));
        // Create object for bottom
        CTBorder border18 = wmlObjectFactory.createCTBorder();
        tblborders.setBottom(border18);
        border18.setVal(org.docx4j.wml.STBorder.SINGLE);
        border18.setColor("auto");
        border18.setSz(BigInteger.valueOf(4));
        border18.setSpace(BigInteger.valueOf(0));
        // Create object for right
        CTBorder border19 = wmlObjectFactory.createCTBorder();
        tblborders.setRight(border19);
        border19.setVal(org.docx4j.wml.STBorder.SINGLE);
        border19.setColor("auto");
        border19.setSz(BigInteger.valueOf(4));
        border19.setSpace(BigInteger.valueOf(0));
        // Create object for insideH
        CTBorder border20 = wmlObjectFactory.createCTBorder();
        tblborders.setInsideH(border20);
        border20.setVal(org.docx4j.wml.STBorder.SINGLE);
        border20.setColor("auto");
        border20.setSz(BigInteger.valueOf(4));
        border20.setSpace(BigInteger.valueOf(0));
        // Create object for insideV
        CTBorder border21 = wmlObjectFactory.createCTBorder();
        tblborders.setInsideV(border21);
        border21.setVal(org.docx4j.wml.STBorder.SINGLE);
        border21.setColor("auto");
        border21.setSz(BigInteger.valueOf(4));
        border21.setSpace(BigInteger.valueOf(0));
        // Create object for tblLook
        CTTblLook tbllook = wmlObjectFactory.createCTTblLook();
        tblpr.setTblLook(tbllook);
        tbllook.setFirstRow(org.docx4j.sharedtypes.STOnOff.ZERO);
        tbllook.setLastRow(org.docx4j.sharedtypes.STOnOff.ZERO);
        tbllook.setFirstColumn(org.docx4j.sharedtypes.STOnOff.ZERO);
        tbllook.setLastColumn(org.docx4j.sharedtypes.STOnOff.ZERO);
        tbllook.setNoHBand(org.docx4j.sharedtypes.STOnOff.ZERO);
        tbllook.setNoVBand(org.docx4j.sharedtypes.STOnOff.ZERO);
        tbllook.setVal("0000");
        // Create object for tblGrid
        TblGrid tblgrid = wmlObjectFactory.createTblGrid();
        tbl.setTblGrid(tblgrid);
        // Create object for gridCol
        TblGridCol tblgridcol = wmlObjectFactory.createTblGridCol();
        tblgrid.getGridCol().add(tblgridcol);
        tblgridcol.setW(BigInteger.valueOf(4260));
        // Create object for gridCol
        TblGridCol tblgridcol2 = wmlObjectFactory.createTblGridCol();
        tblgrid.getGridCol().add(tblgridcol2);
        tblgridcol2.setW(BigInteger.valueOf(3939));
        // Create object for p
        P p48 = wmlObjectFactory.createP();
        body.getContent().add(p48);
        // Create object for pPr
        PPr ppr48 = wmlObjectFactory.createPPr();
        p48.setPPr(ppr48);
        // Create object for rPr
        ParaRPr pararpr48 = wmlObjectFactory.createParaRPr();
        ppr48.setRPr(pararpr48);
        // Create object for rFonts
        RFonts rfonts95 = wmlObjectFactory.createRFonts();
        pararpr48.setRFonts(rfonts95);
        rfonts95.setAscii("Arial");
        rfonts95.setHAnsi("Arial");
        rfonts95.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue109 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr48.setB(booleandefaulttrue109);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue110 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr48.setBCs(booleandefaulttrue110);
        // Create object for color
        Color color95 = wmlObjectFactory.createColor();
        pararpr48.setColor(color95);
        color95.setVal("365F91");
        color95.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color95.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle24 = wmlObjectFactory.createPPrBasePStyle();
        ppr48.setPStyle(pprbasepstyle24);
        pprbasepstyle24.setVal("BodyText");
        // Create object for p
        P p49 = wmlObjectFactory.createP();
        body.getContent().add(p49);
        // Create object for r
        R r49 = wmlObjectFactory.createR();
        p49.getContent().add(r49);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped = wmlObjectFactory.createRTab(rtab);
        r49.getContent().add(rtabWrapped);
        // Create object for rPr
        RPr rpr49 = wmlObjectFactory.createRPr();
        r49.setRPr(rpr49);
        // Create object for rFonts
        RFonts rfonts96 = wmlObjectFactory.createRFonts();
        rpr49.setRFonts(rfonts96);
        rfonts96.setAscii("Arial");
        rfonts96.setHAnsi("Arial");
        rfonts96.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue111 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr49.setB(booleandefaulttrue111);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue112 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr49.setBCs(booleandefaulttrue112);
        // Create object for color
        Color color96 = wmlObjectFactory.createColor();
        rpr49.setColor(color96);
        color96.setVal("365F91");
        color96.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color96.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure = wmlObjectFactory.createHpsMeasure();
        rpr49.setSz(hpsmeasure);
        hpsmeasure.setVal(BigInteger.valueOf(20));
        // Create object for szCs
        HpsMeasure hpsmeasure2 = wmlObjectFactory.createHpsMeasure();
        rpr49.setSzCs(hpsmeasure2);
        hpsmeasure2.setVal(BigInteger.valueOf(20));
        // Create object for r
        R r50 = wmlObjectFactory.createR();
        p49.getContent().add(r50);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab2 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped2 = wmlObjectFactory.createRTab(rtab2);
        r50.getContent().add(rtabWrapped2);
        // Create object for rPr
        RPr rpr50 = wmlObjectFactory.createRPr();
        r50.setRPr(rpr50);
        // Create object for rFonts
        RFonts rfonts97 = wmlObjectFactory.createRFonts();
        rpr50.setRFonts(rfonts97);
        rfonts97.setAscii("Arial");
        rfonts97.setHAnsi("Arial");
        rfonts97.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue113 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr50.setB(booleandefaulttrue113);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue114 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr50.setBCs(booleandefaulttrue114);
        // Create object for color
        Color color97 = wmlObjectFactory.createColor();
        rpr50.setColor(color97);
        color97.setVal("365F91");
        color97.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color97.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure3 = wmlObjectFactory.createHpsMeasure();
        rpr50.setSz(hpsmeasure3);
        hpsmeasure3.setVal(BigInteger.valueOf(20));
        // Create object for szCs
        HpsMeasure hpsmeasure4 = wmlObjectFactory.createHpsMeasure();
        rpr50.setSzCs(hpsmeasure4);
        hpsmeasure4.setVal(BigInteger.valueOf(20));
        // Create object for r
        R r51 = wmlObjectFactory.createR();
        p49.getContent().add(r51);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab3 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped3 = wmlObjectFactory.createRTab(rtab3);
        r51.getContent().add(rtabWrapped3);
        // Create object for rPr
        RPr rpr51 = wmlObjectFactory.createRPr();
        r51.setRPr(rpr51);
        // Create object for rFonts
        RFonts rfonts98 = wmlObjectFactory.createRFonts();
        rpr51.setRFonts(rfonts98);
        rfonts98.setAscii("Arial");
        rfonts98.setHAnsi("Arial");
        rfonts98.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue115 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr51.setB(booleandefaulttrue115);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue116 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr51.setBCs(booleandefaulttrue116);
        // Create object for color
        Color color98 = wmlObjectFactory.createColor();
        rpr51.setColor(color98);
        color98.setVal("365F91");
        color98.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color98.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure5 = wmlObjectFactory.createHpsMeasure();
        rpr51.setSz(hpsmeasure5);
        hpsmeasure5.setVal(BigInteger.valueOf(20));
        // Create object for szCs
        HpsMeasure hpsmeasure6 = wmlObjectFactory.createHpsMeasure();
        rpr51.setSzCs(hpsmeasure6);
        hpsmeasure6.setVal(BigInteger.valueOf(20));
        // Create object for r
        R r52 = wmlObjectFactory.createR();
        p49.getContent().add(r52);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab4 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped4 = wmlObjectFactory.createRTab(rtab4);
        r52.getContent().add(rtabWrapped4);
        // Create object for rPr
        RPr rpr52 = wmlObjectFactory.createRPr();
        r52.setRPr(rpr52);
        // Create object for rFonts
        RFonts rfonts99 = wmlObjectFactory.createRFonts();
        rpr52.setRFonts(rfonts99);
        rfonts99.setAscii("Arial");
        rfonts99.setHAnsi("Arial");
        rfonts99.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue117 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr52.setB(booleandefaulttrue117);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue118 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr52.setBCs(booleandefaulttrue118);
        // Create object for color
        Color color99 = wmlObjectFactory.createColor();
        rpr52.setColor(color99);
        color99.setVal("365F91");
        color99.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color99.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure7 = wmlObjectFactory.createHpsMeasure();
        rpr52.setSz(hpsmeasure7);
        hpsmeasure7.setVal(BigInteger.valueOf(20));
        // Create object for szCs
        HpsMeasure hpsmeasure8 = wmlObjectFactory.createHpsMeasure();
        rpr52.setSzCs(hpsmeasure8);
        hpsmeasure8.setVal(BigInteger.valueOf(20));
        // Create object for r
        R r53 = wmlObjectFactory.createR();
        p49.getContent().add(r53);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab5 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped5 = wmlObjectFactory.createRTab(rtab5);
        r53.getContent().add(rtabWrapped5);
        // Create object for rPr
        RPr rpr53 = wmlObjectFactory.createRPr();
        r53.setRPr(rpr53);
        // Create object for rFonts
        RFonts rfonts100 = wmlObjectFactory.createRFonts();
        rpr53.setRFonts(rfonts100);
        rfonts100.setAscii("Arial");
        rfonts100.setHAnsi("Arial");
        rfonts100.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue119 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr53.setB(booleandefaulttrue119);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue120 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr53.setBCs(booleandefaulttrue120);
        // Create object for color
        Color color100 = wmlObjectFactory.createColor();
        rpr53.setColor(color100);
        color100.setVal("365F91");
        color100.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color100.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure9 = wmlObjectFactory.createHpsMeasure();
        rpr53.setSz(hpsmeasure9);
        hpsmeasure9.setVal(BigInteger.valueOf(20));
        // Create object for szCs
        HpsMeasure hpsmeasure10 = wmlObjectFactory.createHpsMeasure();
        rpr53.setSzCs(hpsmeasure10);
        hpsmeasure10.setVal(BigInteger.valueOf(20));
        // Create object for r
        R r54 = wmlObjectFactory.createR();
        p49.getContent().add(r54);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab6 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped6 = wmlObjectFactory.createRTab(rtab6);
        r54.getContent().add(rtabWrapped6);
        // Create object for rPr
        RPr rpr54 = wmlObjectFactory.createRPr();
        r54.setRPr(rpr54);
        // Create object for rFonts
        RFonts rfonts101 = wmlObjectFactory.createRFonts();
        rpr54.setRFonts(rfonts101);
        rfonts101.setAscii("Arial");
        rfonts101.setHAnsi("Arial");
        rfonts101.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue121 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr54.setB(booleandefaulttrue121);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue122 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr54.setBCs(booleandefaulttrue122);
        // Create object for color
        Color color101 = wmlObjectFactory.createColor();
        rpr54.setColor(color101);
        color101.setVal("365F91");
        color101.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color101.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure11 = wmlObjectFactory.createHpsMeasure();
        rpr54.setSz(hpsmeasure11);
        hpsmeasure11.setVal(BigInteger.valueOf(20));
        // Create object for szCs
        HpsMeasure hpsmeasure12 = wmlObjectFactory.createHpsMeasure();
        rpr54.setSzCs(hpsmeasure12);
        hpsmeasure12.setVal(BigInteger.valueOf(20));
        // Create object for r
        R r55 = wmlObjectFactory.createR();
        p49.getContent().add(r55);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab7 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped7 = wmlObjectFactory.createRTab(rtab7);
        r55.getContent().add(rtabWrapped7);
        // Create object for rPr
        RPr rpr55 = wmlObjectFactory.createRPr();
        r55.setRPr(rpr55);
        // Create object for rFonts
        RFonts rfonts102 = wmlObjectFactory.createRFonts();
        rpr55.setRFonts(rfonts102);
        rfonts102.setAscii("Arial");
        rfonts102.setHAnsi("Arial");
        rfonts102.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue123 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr55.setB(booleandefaulttrue123);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue124 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr55.setBCs(booleandefaulttrue124);
        // Create object for color
        Color color102 = wmlObjectFactory.createColor();
        rpr55.setColor(color102);
        color102.setVal("365F91");
        color102.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color102.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure13 = wmlObjectFactory.createHpsMeasure();
        rpr55.setSz(hpsmeasure13);
        hpsmeasure13.setVal(BigInteger.valueOf(20));
        // Create object for szCs
        HpsMeasure hpsmeasure14 = wmlObjectFactory.createHpsMeasure();
        rpr55.setSzCs(hpsmeasure14);
        hpsmeasure14.setVal(BigInteger.valueOf(20));
        // Create object for r
        R r56 = wmlObjectFactory.createR();
        p49.getContent().add(r56);
        // Create object for t (wrapped in JAXBElement)
        Text text49 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped49 = wmlObjectFactory.createRT(text49);
        r56.getContent().add(textWrapped49);
        text49.setValue(
                "Signature                                                                                                                                                             ");
        text49.setSpace("preserve");
        // Create object for rPr
        RPr rpr56 = wmlObjectFactory.createRPr();
        r56.setRPr(rpr56);
        // Create object for rFonts
        RFonts rfonts103 = wmlObjectFactory.createRFonts();
        rpr56.setRFonts(rfonts103);
        rfonts103.setAscii("Arial");
        rfonts103.setHAnsi("Arial");
        rfonts103.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue125 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr56.setB(booleandefaulttrue125);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue126 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr56.setBCs(booleandefaulttrue126);
        // Create object for color
        Color color103 = wmlObjectFactory.createColor();
        rpr56.setColor(color103);
        color103.setVal("365F91");
        color103.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color103.setThemeShade("BF");
        // Create object for pPr
        PPr ppr49 = wmlObjectFactory.createPPr();
        p49.setPPr(ppr49);
        // Create object for rPr
        ParaRPr pararpr49 = wmlObjectFactory.createParaRPr();
        ppr49.setRPr(pararpr49);
        // Create object for rFonts
        RFonts rfonts104 = wmlObjectFactory.createRFonts();
        pararpr49.setRFonts(rfonts104);
        rfonts104.setAscii("Arial");
        rfonts104.setHAnsi("Arial");
        rfonts104.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue127 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr49.setB(booleandefaulttrue127);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue128 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr49.setBCs(booleandefaulttrue128);
        // Create object for color
        Color color104 = wmlObjectFactory.createColor();
        pararpr49.setColor(color104);
        color104.setVal("365F91");
        color104.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color104.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure15 = wmlObjectFactory.createHpsMeasure();
        pararpr49.setSz(hpsmeasure15);
        hpsmeasure15.setVal(BigInteger.valueOf(20));
        // Create object for szCs
        HpsMeasure hpsmeasure16 = wmlObjectFactory.createHpsMeasure();
        pararpr49.setSzCs(hpsmeasure16);
        hpsmeasure16.setVal(BigInteger.valueOf(20));
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle25 = wmlObjectFactory.createPPrBasePStyle();
        ppr49.setPStyle(pprbasepstyle25);
        pprbasepstyle25.setVal("BodyText");
        // Create object for p
        P p50 = wmlObjectFactory.createP();
        body.getContent().add(p50);
        // Create object for r
        R r57 = wmlObjectFactory.createR();
        p50.getContent().add(r57);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab8 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped8 = wmlObjectFactory.createRTab(rtab8);
        r57.getContent().add(rtabWrapped8);
        // Create object for rPr
        RPr rpr57 = wmlObjectFactory.createRPr();
        r57.setRPr(rpr57);
        // Create object for rFonts
        RFonts rfonts105 = wmlObjectFactory.createRFonts();
        rpr57.setRFonts(rfonts105);
        rfonts105.setAscii("Brush Script MT");
        rfonts105.setHAnsi("Brush Script MT");
        rfonts105.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue129 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr57.setB(booleandefaulttrue129);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue130 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr57.setBCs(booleandefaulttrue130);
        // Create object for color
        Color color105 = wmlObjectFactory.createColor();
        rpr57.setColor(color105);
        color105.setVal("365F91");
        color105.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color105.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure17 = wmlObjectFactory.createHpsMeasure();
        rpr57.setSz(hpsmeasure17);
        hpsmeasure17.setVal(BigInteger.valueOf(28));
        // Create object for szCs
        HpsMeasure hpsmeasure18 = wmlObjectFactory.createHpsMeasure();
        rpr57.setSzCs(hpsmeasure18);
        hpsmeasure18.setVal(BigInteger.valueOf(28));
        // Create object for r
        R r58 = wmlObjectFactory.createR();
        p50.getContent().add(r58);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab9 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped9 = wmlObjectFactory.createRTab(rtab9);
        r58.getContent().add(rtabWrapped9);
        // Create object for rPr
        RPr rpr58 = wmlObjectFactory.createRPr();
        r58.setRPr(rpr58);
        // Create object for rFonts
        RFonts rfonts106 = wmlObjectFactory.createRFonts();
        rpr58.setRFonts(rfonts106);
        rfonts106.setAscii("Brush Script MT");
        rfonts106.setHAnsi("Brush Script MT");
        rfonts106.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue131 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr58.setB(booleandefaulttrue131);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue132 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr58.setBCs(booleandefaulttrue132);
        // Create object for color
        Color color106 = wmlObjectFactory.createColor();
        rpr58.setColor(color106);
        color106.setVal("365F91");
        color106.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color106.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure19 = wmlObjectFactory.createHpsMeasure();
        rpr58.setSz(hpsmeasure19);
        hpsmeasure19.setVal(BigInteger.valueOf(28));
        // Create object for szCs
        HpsMeasure hpsmeasure20 = wmlObjectFactory.createHpsMeasure();
        rpr58.setSzCs(hpsmeasure20);
        hpsmeasure20.setVal(BigInteger.valueOf(28));
        // Create object for r
        R r59 = wmlObjectFactory.createR();
        p50.getContent().add(r59);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab10 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped10 = wmlObjectFactory.createRTab(rtab10);
        r59.getContent().add(rtabWrapped10);
        // Create object for rPr
        RPr rpr59 = wmlObjectFactory.createRPr();
        r59.setRPr(rpr59);
        // Create object for rFonts
        RFonts rfonts107 = wmlObjectFactory.createRFonts();
        rpr59.setRFonts(rfonts107);
        rfonts107.setAscii("Brush Script MT");
        rfonts107.setHAnsi("Brush Script MT");
        rfonts107.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue133 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr59.setB(booleandefaulttrue133);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue134 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr59.setBCs(booleandefaulttrue134);
        // Create object for color
        Color color107 = wmlObjectFactory.createColor();
        rpr59.setColor(color107);
        color107.setVal("365F91");
        color107.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color107.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure21 = wmlObjectFactory.createHpsMeasure();
        rpr59.setSz(hpsmeasure21);
        hpsmeasure21.setVal(BigInteger.valueOf(28));
        // Create object for szCs
        HpsMeasure hpsmeasure22 = wmlObjectFactory.createHpsMeasure();
        rpr59.setSzCs(hpsmeasure22);
        hpsmeasure22.setVal(BigInteger.valueOf(28));
        // Create object for r
        R r60 = wmlObjectFactory.createR();
        p50.getContent().add(r60);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab11 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped11 = wmlObjectFactory.createRTab(rtab11);
        r60.getContent().add(rtabWrapped11);
        // Create object for rPr
        RPr rpr60 = wmlObjectFactory.createRPr();
        r60.setRPr(rpr60);
        // Create object for rFonts
        RFonts rfonts108 = wmlObjectFactory.createRFonts();
        rpr60.setRFonts(rfonts108);
        rfonts108.setAscii("Brush Script MT");
        rfonts108.setHAnsi("Brush Script MT");
        rfonts108.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue135 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr60.setB(booleandefaulttrue135);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue136 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr60.setBCs(booleandefaulttrue136);
        // Create object for color
        Color color108 = wmlObjectFactory.createColor();
        rpr60.setColor(color108);
        color108.setVal("365F91");
        color108.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color108.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure23 = wmlObjectFactory.createHpsMeasure();
        rpr60.setSz(hpsmeasure23);
        hpsmeasure23.setVal(BigInteger.valueOf(28));
        // Create object for szCs
        HpsMeasure hpsmeasure24 = wmlObjectFactory.createHpsMeasure();
        rpr60.setSzCs(hpsmeasure24);
        hpsmeasure24.setVal(BigInteger.valueOf(28));
        // Create object for r
        R r61 = wmlObjectFactory.createR();
        p50.getContent().add(r61);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab12 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped12 = wmlObjectFactory.createRTab(rtab12);
        r61.getContent().add(rtabWrapped12);
        // Create object for rPr
        RPr rpr61 = wmlObjectFactory.createRPr();
        r61.setRPr(rpr61);
        // Create object for rFonts
        RFonts rfonts109 = wmlObjectFactory.createRFonts();
        rpr61.setRFonts(rfonts109);
        rfonts109.setAscii("Brush Script MT");
        rfonts109.setHAnsi("Brush Script MT");
        rfonts109.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue137 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr61.setB(booleandefaulttrue137);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue138 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr61.setBCs(booleandefaulttrue138);
        // Create object for color
        Color color109 = wmlObjectFactory.createColor();
        rpr61.setColor(color109);
        color109.setVal("365F91");
        color109.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color109.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure25 = wmlObjectFactory.createHpsMeasure();
        rpr61.setSz(hpsmeasure25);
        hpsmeasure25.setVal(BigInteger.valueOf(28));
        // Create object for szCs
        HpsMeasure hpsmeasure26 = wmlObjectFactory.createHpsMeasure();
        rpr61.setSzCs(hpsmeasure26);
        hpsmeasure26.setVal(BigInteger.valueOf(28));
        // Create object for r
        R r62 = wmlObjectFactory.createR();
        p50.getContent().add(r62);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab13 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped13 = wmlObjectFactory.createRTab(rtab13);
        r62.getContent().add(rtabWrapped13);
        // Create object for rPr
        RPr rpr62 = wmlObjectFactory.createRPr();
        r62.setRPr(rpr62);
        // Create object for rFonts
        RFonts rfonts110 = wmlObjectFactory.createRFonts();
        rpr62.setRFonts(rfonts110);
        rfonts110.setAscii("Brush Script MT");
        rfonts110.setHAnsi("Brush Script MT");
        rfonts110.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue139 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr62.setB(booleandefaulttrue139);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue140 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr62.setBCs(booleandefaulttrue140);
        // Create object for color
        Color color110 = wmlObjectFactory.createColor();
        rpr62.setColor(color110);
        color110.setVal("365F91");
        color110.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color110.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure27 = wmlObjectFactory.createHpsMeasure();
        rpr62.setSz(hpsmeasure27);
        hpsmeasure27.setVal(BigInteger.valueOf(28));
        // Create object for szCs
        HpsMeasure hpsmeasure28 = wmlObjectFactory.createHpsMeasure();
        rpr62.setSzCs(hpsmeasure28);
        hpsmeasure28.setVal(BigInteger.valueOf(28));
        // Create object for r
        R r63 = wmlObjectFactory.createR();
        p50.getContent().add(r63);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab14 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped14 = wmlObjectFactory.createRTab(rtab14);
        r63.getContent().add(rtabWrapped14);
        // Create object for rPr
        RPr rpr63 = wmlObjectFactory.createRPr();
        r63.setRPr(rpr63);
        // Create object for rFonts
        RFonts rfonts111 = wmlObjectFactory.createRFonts();
        rpr63.setRFonts(rfonts111);
        rfonts111.setAscii("Brush Script MT");
        rfonts111.setHAnsi("Brush Script MT");
        rfonts111.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue141 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr63.setB(booleandefaulttrue141);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue142 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr63.setBCs(booleandefaulttrue142);
        // Create object for color
        Color color111 = wmlObjectFactory.createColor();
        rpr63.setColor(color111);
        color111.setVal("365F91");
        color111.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color111.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure29 = wmlObjectFactory.createHpsMeasure();
        rpr63.setSz(hpsmeasure29);
        hpsmeasure29.setVal(BigInteger.valueOf(28));
        // Create object for szCs
        HpsMeasure hpsmeasure30 = wmlObjectFactory.createHpsMeasure();
        rpr63.setSzCs(hpsmeasure30);
        hpsmeasure30.setVal(BigInteger.valueOf(28));
        // Create object for r
        R r64 = wmlObjectFactory.createR();
        p50.getContent().add(r64);
        // Create object for t (wrapped in JAXBElement)
        Text text50 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped50 = wmlObjectFactory.createRT(text50);
        r64.getContent().add(textWrapped50);
        text50.setValue("A. Lawandy");
        // Create object for rPr
        RPr rpr64 = wmlObjectFactory.createRPr();
        r64.setRPr(rpr64);
        // Create object for rFonts
        RFonts rfonts112 = wmlObjectFactory.createRFonts();
        rpr64.setRFonts(rfonts112);
        rfonts112.setAscii("MV Boli");
        rfonts112.setHAnsi("MV Boli");
        rfonts112.setCs("MV Boli");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue143 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr64.setBCs(booleandefaulttrue143);
        // Create object for i
        BooleanDefaultTrue booleandefaulttrue144 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr64.setI(booleandefaulttrue144);
        // Create object for color
        Color color112 = wmlObjectFactory.createColor();
        rpr64.setColor(color112);
        color112.setVal("1F497D");
        // Create object for sz
        HpsMeasure hpsmeasure31 = wmlObjectFactory.createHpsMeasure();
        rpr64.setSz(hpsmeasure31);
        hpsmeasure31.setVal(BigInteger.valueOf(32));
        // Create object for szCs
        HpsMeasure hpsmeasure32 = wmlObjectFactory.createHpsMeasure();
        rpr64.setSzCs(hpsmeasure32);
        hpsmeasure32.setVal(BigInteger.valueOf(32));
        // Create object for pPr
        PPr ppr50 = wmlObjectFactory.createPPr();
        p50.setPPr(ppr50);
        // Create object for rPr
        ParaRPr pararpr50 = wmlObjectFactory.createParaRPr();
        ppr50.setRPr(pararpr50);
        // Create object for rFonts
        RFonts rfonts113 = wmlObjectFactory.createRFonts();
        pararpr50.setRFonts(rfonts113);
        rfonts113.setAscii("MV Boli");
        rfonts113.setHAnsi("MV Boli");
        rfonts113.setCs("MV Boli");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue145 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr50.setBCs(booleandefaulttrue145);
        // Create object for i
        BooleanDefaultTrue booleandefaulttrue146 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr50.setI(booleandefaulttrue146);
        // Create object for color
        Color color113 = wmlObjectFactory.createColor();
        pararpr50.setColor(color113);
        color113.setVal("365F91");
        color113.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color113.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure33 = wmlObjectFactory.createHpsMeasure();
        pararpr50.setSz(hpsmeasure33);
        hpsmeasure33.setVal(BigInteger.valueOf(28));
        // Create object for szCs
        HpsMeasure hpsmeasure34 = wmlObjectFactory.createHpsMeasure();
        pararpr50.setSzCs(hpsmeasure34);
        hpsmeasure34.setVal(BigInteger.valueOf(28));
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle26 = wmlObjectFactory.createPPrBasePStyle();
        ppr50.setPStyle(pprbasepstyle26);
        pprbasepstyle26.setVal("BodyText");
        // Create object for p
        P p51 = wmlObjectFactory.createP();
        body.getContent().add(p51);
        // Create object for r
        R r65 = wmlObjectFactory.createR();
        p51.getContent().add(r65);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab15 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped15 = wmlObjectFactory.createRTab(rtab15);
        r65.getContent().add(rtabWrapped15);
        // Create object for rPr
        RPr rpr65 = wmlObjectFactory.createRPr();
        r65.setRPr(rpr65);
        // Create object for rFonts
        RFonts rfonts114 = wmlObjectFactory.createRFonts();
        rpr65.setRFonts(rfonts114);
        rfonts114.setAscii("CommercialScript BT");
        rfonts114.setHAnsi("CommercialScript BT");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue147 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr65.setBCs(booleandefaulttrue147);
        // Create object for color
        Color color114 = wmlObjectFactory.createColor();
        rpr65.setColor(color114);
        color114.setVal("365F91");
        color114.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color114.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure35 = wmlObjectFactory.createHpsMeasure();
        rpr65.setSz(hpsmeasure35);
        hpsmeasure35.setVal(BigInteger.valueOf(32));
        // Create object for szCs
        HpsMeasure hpsmeasure36 = wmlObjectFactory.createHpsMeasure();
        rpr65.setSzCs(hpsmeasure36);
        hpsmeasure36.setVal(BigInteger.valueOf(32));
        // Create object for r
        R r66 = wmlObjectFactory.createR();
        p51.getContent().add(r66);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab16 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped16 = wmlObjectFactory.createRTab(rtab16);
        r66.getContent().add(rtabWrapped16);
        // Create object for rPr
        RPr rpr66 = wmlObjectFactory.createRPr();
        r66.setRPr(rpr66);
        // Create object for rFonts
        RFonts rfonts115 = wmlObjectFactory.createRFonts();
        rpr66.setRFonts(rfonts115);
        rfonts115.setAscii("CommercialScript BT");
        rfonts115.setHAnsi("CommercialScript BT");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue148 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr66.setBCs(booleandefaulttrue148);
        // Create object for color
        Color color115 = wmlObjectFactory.createColor();
        rpr66.setColor(color115);
        color115.setVal("365F91");
        color115.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color115.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure37 = wmlObjectFactory.createHpsMeasure();
        rpr66.setSz(hpsmeasure37);
        hpsmeasure37.setVal(BigInteger.valueOf(32));
        // Create object for szCs
        HpsMeasure hpsmeasure38 = wmlObjectFactory.createHpsMeasure();
        rpr66.setSzCs(hpsmeasure38);
        hpsmeasure38.setVal(BigInteger.valueOf(32));
        // Create object for r
        R r67 = wmlObjectFactory.createR();
        p51.getContent().add(r67);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab17 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped17 = wmlObjectFactory.createRTab(rtab17);
        r67.getContent().add(rtabWrapped17);
        // Create object for rPr
        RPr rpr67 = wmlObjectFactory.createRPr();
        r67.setRPr(rpr67);
        // Create object for rFonts
        RFonts rfonts116 = wmlObjectFactory.createRFonts();
        rpr67.setRFonts(rfonts116);
        rfonts116.setAscii("CommercialScript BT");
        rfonts116.setHAnsi("CommercialScript BT");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue149 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr67.setBCs(booleandefaulttrue149);
        // Create object for color
        Color color116 = wmlObjectFactory.createColor();
        rpr67.setColor(color116);
        color116.setVal("365F91");
        color116.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color116.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure39 = wmlObjectFactory.createHpsMeasure();
        rpr67.setSz(hpsmeasure39);
        hpsmeasure39.setVal(BigInteger.valueOf(32));
        // Create object for szCs
        HpsMeasure hpsmeasure40 = wmlObjectFactory.createHpsMeasure();
        rpr67.setSzCs(hpsmeasure40);
        hpsmeasure40.setVal(BigInteger.valueOf(32));
        // Create object for r
        R r68 = wmlObjectFactory.createR();
        p51.getContent().add(r68);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab18 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped18 = wmlObjectFactory.createRTab(rtab18);
        r68.getContent().add(rtabWrapped18);
        // Create object for rPr
        RPr rpr68 = wmlObjectFactory.createRPr();
        r68.setRPr(rpr68);
        // Create object for rFonts
        RFonts rfonts117 = wmlObjectFactory.createRFonts();
        rpr68.setRFonts(rfonts117);
        rfonts117.setAscii("CommercialScript BT");
        rfonts117.setHAnsi("CommercialScript BT");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue150 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr68.setBCs(booleandefaulttrue150);
        // Create object for color
        Color color117 = wmlObjectFactory.createColor();
        rpr68.setColor(color117);
        color117.setVal("365F91");
        color117.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color117.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure41 = wmlObjectFactory.createHpsMeasure();
        rpr68.setSz(hpsmeasure41);
        hpsmeasure41.setVal(BigInteger.valueOf(32));
        // Create object for szCs
        HpsMeasure hpsmeasure42 = wmlObjectFactory.createHpsMeasure();
        rpr68.setSzCs(hpsmeasure42);
        hpsmeasure42.setVal(BigInteger.valueOf(32));
        // Create object for r
        R r69 = wmlObjectFactory.createR();
        p51.getContent().add(r69);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab19 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped19 = wmlObjectFactory.createRTab(rtab19);
        r69.getContent().add(rtabWrapped19);
        // Create object for rPr
        RPr rpr69 = wmlObjectFactory.createRPr();
        r69.setRPr(rpr69);
        // Create object for rFonts
        RFonts rfonts118 = wmlObjectFactory.createRFonts();
        rpr69.setRFonts(rfonts118);
        rfonts118.setAscii("CommercialScript BT");
        rfonts118.setHAnsi("CommercialScript BT");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue151 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr69.setBCs(booleandefaulttrue151);
        // Create object for color
        Color color118 = wmlObjectFactory.createColor();
        rpr69.setColor(color118);
        color118.setVal("365F91");
        color118.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color118.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure43 = wmlObjectFactory.createHpsMeasure();
        rpr69.setSz(hpsmeasure43);
        hpsmeasure43.setVal(BigInteger.valueOf(32));
        // Create object for szCs
        HpsMeasure hpsmeasure44 = wmlObjectFactory.createHpsMeasure();
        rpr69.setSzCs(hpsmeasure44);
        hpsmeasure44.setVal(BigInteger.valueOf(32));
        // Create object for r
        R r70 = wmlObjectFactory.createR();
        p51.getContent().add(r70);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab20 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped20 = wmlObjectFactory.createRTab(rtab20);
        r70.getContent().add(rtabWrapped20);
        // Create object for rPr
        RPr rpr70 = wmlObjectFactory.createRPr();
        r70.setRPr(rpr70);
        // Create object for rFonts
        RFonts rfonts119 = wmlObjectFactory.createRFonts();
        rpr70.setRFonts(rfonts119);
        rfonts119.setAscii("CommercialScript BT");
        rfonts119.setHAnsi("CommercialScript BT");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue152 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr70.setBCs(booleandefaulttrue152);
        // Create object for color
        Color color119 = wmlObjectFactory.createColor();
        rpr70.setColor(color119);
        color119.setVal("365F91");
        color119.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color119.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure45 = wmlObjectFactory.createHpsMeasure();
        rpr70.setSz(hpsmeasure45);
        hpsmeasure45.setVal(BigInteger.valueOf(32));
        // Create object for szCs
        HpsMeasure hpsmeasure46 = wmlObjectFactory.createHpsMeasure();
        rpr70.setSzCs(hpsmeasure46);
        hpsmeasure46.setVal(BigInteger.valueOf(32));
        // Create object for r
        R r71 = wmlObjectFactory.createR();
        p51.getContent().add(r71);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab21 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped21 = wmlObjectFactory.createRTab(rtab21);
        r71.getContent().add(rtabWrapped21);
        // Create object for rPr
        RPr rpr71 = wmlObjectFactory.createRPr();
        r71.setRPr(rpr71);
        // Create object for rFonts
        RFonts rfonts120 = wmlObjectFactory.createRFonts();
        rpr71.setRFonts(rfonts120);
        rfonts120.setAscii("CommercialScript BT");
        rfonts120.setHAnsi("CommercialScript BT");
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue153 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr71.setBCs(booleandefaulttrue153);
        // Create object for color
        Color color120 = wmlObjectFactory.createColor();
        rpr71.setColor(color120);
        color120.setVal("365F91");
        color120.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color120.setThemeShade("BF");
        // Create object for sz
        HpsMeasure hpsmeasure47 = wmlObjectFactory.createHpsMeasure();
        rpr71.setSz(hpsmeasure47);
        hpsmeasure47.setVal(BigInteger.valueOf(32));
        // Create object for szCs
        HpsMeasure hpsmeasure48 = wmlObjectFactory.createHpsMeasure();
        rpr71.setSzCs(hpsmeasure48);
        hpsmeasure48.setVal(BigInteger.valueOf(32));
        // Create object for r
        R r72 = wmlObjectFactory.createR();
        p51.getContent().add(r72);
        // Create object for t (wrapped in JAXBElement)
        Text text51 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped51 = wmlObjectFactory.createRT(text51);
        r72.getContent().add(textWrapped51);
        text51.setValue("Anthony Lawandy");
        // Create object for rPr
        RPr rpr72 = wmlObjectFactory.createRPr();
        r72.setRPr(rpr72);
        // Create object for rFonts
        RFonts rfonts121 = wmlObjectFactory.createRFonts();
        rpr72.setRFonts(rfonts121);
        rfonts121.setAscii("Arial");
        rfonts121.setHAnsi("Arial");
        rfonts121.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue154 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr72.setB(booleandefaulttrue154);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue155 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr72.setBCs(booleandefaulttrue155);
        // Create object for color
        Color color121 = wmlObjectFactory.createColor();
        rpr72.setColor(color121);
        color121.setVal("365F91");
        color121.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color121.setThemeShade("BF");
        // Create object for r
        R r73 = wmlObjectFactory.createR();
        p51.getContent().add(r73);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab22 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped22 = wmlObjectFactory.createRTab(rtab22);
        r73.getContent().add(rtabWrapped22);
        // Create object for rPr
        RPr rpr73 = wmlObjectFactory.createRPr();
        r73.setRPr(rpr73);
        // Create object for rFonts
        RFonts rfonts122 = wmlObjectFactory.createRFonts();
        rpr73.setRFonts(rfonts122);
        rfonts122.setAscii("Arial");
        rfonts122.setHAnsi("Arial");
        rfonts122.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue156 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr73.setB(booleandefaulttrue156);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue157 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr73.setBCs(booleandefaulttrue157);
        // Create object for color
        Color color122 = wmlObjectFactory.createColor();
        rpr73.setColor(color122);
        color122.setVal("365F91");
        color122.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color122.setThemeShade("BF");
        // Create object for r
        R r74 = wmlObjectFactory.createR();
        p51.getContent().add(r74);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab23 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped23 = wmlObjectFactory.createRTab(rtab23);
        r74.getContent().add(rtabWrapped23);
        // Create object for rPr
        RPr rpr74 = wmlObjectFactory.createRPr();
        r74.setRPr(rpr74);
        // Create object for rFonts
        RFonts rfonts123 = wmlObjectFactory.createRFonts();
        rpr74.setRFonts(rfonts123);
        rfonts123.setAscii("Arial");
        rfonts123.setHAnsi("Arial");
        rfonts123.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue158 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr74.setB(booleandefaulttrue158);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue159 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr74.setBCs(booleandefaulttrue159);
        // Create object for color
        Color color123 = wmlObjectFactory.createColor();
        rpr74.setColor(color123);
        color123.setVal("365F91");
        color123.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color123.setThemeShade("BF");
        // Create object for r
        R r75 = wmlObjectFactory.createR();
        p51.getContent().add(r75);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab24 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped24 = wmlObjectFactory.createRTab(rtab24);
        r75.getContent().add(rtabWrapped24);
        // Create object for rPr
        RPr rpr75 = wmlObjectFactory.createRPr();
        r75.setRPr(rpr75);
        // Create object for rFonts
        RFonts rfonts124 = wmlObjectFactory.createRFonts();
        rpr75.setRFonts(rfonts124);
        rfonts124.setAscii("Arial");
        rfonts124.setHAnsi("Arial");
        rfonts124.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue160 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr75.setB(booleandefaulttrue160);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue161 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr75.setBCs(booleandefaulttrue161);
        // Create object for color
        Color color124 = wmlObjectFactory.createColor();
        rpr75.setColor(color124);
        color124.setVal("365F91");
        color124.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color124.setThemeShade("BF");
        // Create object for r
        R r76 = wmlObjectFactory.createR();
        p51.getContent().add(r76);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab25 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped25 = wmlObjectFactory.createRTab(rtab25);
        r76.getContent().add(rtabWrapped25);
        // Create object for rPr
        RPr rpr76 = wmlObjectFactory.createRPr();
        r76.setRPr(rpr76);
        // Create object for rFonts
        RFonts rfonts125 = wmlObjectFactory.createRFonts();
        rpr76.setRFonts(rfonts125);
        rfonts125.setAscii("Arial");
        rfonts125.setHAnsi("Arial");
        rfonts125.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue162 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr76.setB(booleandefaulttrue162);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue163 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr76.setBCs(booleandefaulttrue163);
        // Create object for color
        Color color125 = wmlObjectFactory.createColor();
        rpr76.setColor(color125);
        color125.setVal("365F91");
        color125.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color125.setThemeShade("BF");
        // Create object for r
        R r77 = wmlObjectFactory.createR();
        p51.getContent().add(r77);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab26 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped26 = wmlObjectFactory.createRTab(rtab26);
        r77.getContent().add(rtabWrapped26);
        // Create object for rPr
        RPr rpr77 = wmlObjectFactory.createRPr();
        r77.setRPr(rpr77);
        // Create object for rFonts
        RFonts rfonts126 = wmlObjectFactory.createRFonts();
        rpr77.setRFonts(rfonts126);
        rfonts126.setAscii("Arial");
        rfonts126.setHAnsi("Arial");
        rfonts126.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue164 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr77.setB(booleandefaulttrue164);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue165 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr77.setBCs(booleandefaulttrue165);
        // Create object for color
        Color color126 = wmlObjectFactory.createColor();
        rpr77.setColor(color126);
        color126.setVal("365F91");
        color126.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color126.setThemeShade("BF");
        // Create object for r
        R r78 = wmlObjectFactory.createR();
        p51.getContent().add(r78);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab27 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped27 = wmlObjectFactory.createRTab(rtab27);
        r78.getContent().add(rtabWrapped27);
        // Create object for rPr
        RPr rpr78 = wmlObjectFactory.createRPr();
        r78.setRPr(rpr78);
        // Create object for rFonts
        RFonts rfonts127 = wmlObjectFactory.createRFonts();
        rpr78.setRFonts(rfonts127);
        rfonts127.setAscii("Arial");
        rfonts127.setHAnsi("Arial");
        rfonts127.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue166 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr78.setB(booleandefaulttrue166);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue167 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr78.setBCs(booleandefaulttrue167);
        // Create object for color
        Color color127 = wmlObjectFactory.createColor();
        rpr78.setColor(color127);
        color127.setVal("365F91");
        color127.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color127.setThemeShade("BF");
        // Create object for r
        R r79 = wmlObjectFactory.createR();
        p51.getContent().add(r79);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab28 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped28 = wmlObjectFactory.createRTab(rtab28);
        r79.getContent().add(rtabWrapped28);
        // Create object for rPr
        RPr rpr79 = wmlObjectFactory.createRPr();
        r79.setRPr(rpr79);
        // Create object for rFonts
        RFonts rfonts128 = wmlObjectFactory.createRFonts();
        rpr79.setRFonts(rfonts128);
        rfonts128.setAscii("Arial");
        rfonts128.setHAnsi("Arial");
        rfonts128.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue168 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr79.setB(booleandefaulttrue168);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue169 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr79.setBCs(booleandefaulttrue169);
        // Create object for color
        Color color128 = wmlObjectFactory.createColor();
        rpr79.setColor(color128);
        color128.setVal("365F91");
        color128.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color128.setThemeShade("BF");
        // Create object for r
        R r80 = wmlObjectFactory.createR();
        p51.getContent().add(r80);
        // Create object for tab (wrapped in JAXBElement)
        R.Tab rtab29 = wmlObjectFactory.createRTab();
        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped29 = wmlObjectFactory.createRTab(rtab29);
        r80.getContent().add(rtabWrapped29);
        // Create object for t (wrapped in JAXBElement)
        Text text52 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped52 = wmlObjectFactory.createRT(text52);
        r80.getContent().add(textWrapped52);
        text52.setValue("Field Inspector");
        // Create object for rPr
        RPr rpr80 = wmlObjectFactory.createRPr();
        r80.setRPr(rpr80);
        Br pageBreak = wmlObjectFactory.createBr();
        pageBreak.setType(STBrType.PAGE);
        r80.getContent().add(pageBreak);
        // Create object for rFonts
        RFonts rfonts129 = wmlObjectFactory.createRFonts();
        rpr80.setRFonts(rfonts129);
        rfonts129.setAscii("Arial");
        rfonts129.setHAnsi("Arial");
        rfonts129.setCs("Arial");
        // Create object for b
        BooleanDefaultTrue booleandefaulttrue170 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr80.setB(booleandefaulttrue170);
        // Create object for bCs
        BooleanDefaultTrue booleandefaulttrue171 = wmlObjectFactory.createBooleanDefaultTrue();
        rpr80.setBCs(booleandefaulttrue171);
        // Create object for color
        Color color129 = wmlObjectFactory.createColor();
        rpr80.setColor(color129);
        color129.setVal("365F91");
        color129.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color129.setThemeShade("BF");
        // Create object for pPr
        PPr ppr51 = wmlObjectFactory.createPPr();
        p51.setPPr(ppr51);
        // Create object for rPr
        ParaRPr pararpr51 = wmlObjectFactory.createParaRPr();
        ppr51.setRPr(pararpr51);
        // Create object for color
        Color color130 = wmlObjectFactory.createColor();
        pararpr51.setColor(color130);
        color130.setVal("365F91");
        color130.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color130.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle27 = wmlObjectFactory.createPPrBasePStyle();
        ppr51.setPStyle(pprbasepstyle27);
        pprbasepstyle27.setVal("BodyText");
        // Create object for p
        P p52 = wmlObjectFactory.createP();
        body.getContent().add(p52);
        // Create object for pPr
        PPr ppr52 = wmlObjectFactory.createPPr();
        p52.setPPr(ppr52);
        // Create object for rPr
        ParaRPr pararpr52 = wmlObjectFactory.createParaRPr();
        ppr52.setRPr(pararpr52);
        // Create object for noProof
        BooleanDefaultTrue booleandefaulttrue172 = wmlObjectFactory.createBooleanDefaultTrue();
        pararpr52.setNoProof(booleandefaulttrue172);
        // Create object for color
        Color color131 = wmlObjectFactory.createColor();
        pararpr52.setColor(color131);
        color131.setVal("365F91");
        color131.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color131.setThemeShade("BF");
        // Create object for lang
        CTLanguage language = wmlObjectFactory.createCTLanguage();
        pararpr52.setLang(language);
        language.setVal("en-CA");
        language.setEastAsia("en-CA");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle28 = wmlObjectFactory.createPPrBasePStyle();
        ppr52.setPStyle(pprbasepstyle28);
        pprbasepstyle28.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind5 = wmlObjectFactory.createPPrBaseInd();
        ppr52.setInd(pprbaseind5);
        pprbaseind5.setLeft(BigInteger.valueOf(-2268));
        pprbaseind5.setRight(BigInteger.valueOf(-1238));
        // Create object for jc
        Jc jc2 = wmlObjectFactory.createJc();
        ppr52.setJc(jc2);
        jc2.setVal(org.docx4j.wml.JcEnumeration.CENTER);
        // Create object for p
        P p53 = wmlObjectFactory.createP();
        body.getContent().add(p53);
        // Create object for pPr
        PPr ppr53 = wmlObjectFactory.createPPr();
        p53.setPPr(ppr53);
        // Create object for rPr
        ParaRPr pararpr53 = wmlObjectFactory.createParaRPr();
        ppr53.setRPr(pararpr53);
        // Create object for color
        Color color132 = wmlObjectFactory.createColor();
        pararpr53.setColor(color132);
        color132.setVal("365F91");
        color132.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color132.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle29 = wmlObjectFactory.createPPrBasePStyle();
        ppr53.setPStyle(pprbasepstyle29);
        pprbasepstyle29.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind6 = wmlObjectFactory.createPPrBaseInd();
        ppr53.setInd(pprbaseind6);
        pprbaseind6.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p54 = wmlObjectFactory.createP();
        body.getContent().add(p54);
        // Create object for pPr
        PPr ppr54 = wmlObjectFactory.createPPr();
        p54.setPPr(ppr54);
        // Create object for rPr
        ParaRPr pararpr54 = wmlObjectFactory.createParaRPr();
        ppr54.setRPr(pararpr54);
        // Create object for color
        Color color133 = wmlObjectFactory.createColor();
        pararpr54.setColor(color133);
        color133.setVal("365F91");
        color133.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color133.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle30 = wmlObjectFactory.createPPrBasePStyle();
        ppr54.setPStyle(pprbasepstyle30);
        pprbasepstyle30.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind7 = wmlObjectFactory.createPPrBaseInd();
        ppr54.setInd(pprbaseind7);
        pprbaseind7.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p55 = wmlObjectFactory.createP();
        body.getContent().add(p55);
        // Create object for pPr
        PPr ppr55 = wmlObjectFactory.createPPr();
        p55.setPPr(ppr55);
        // Create object for rPr
        ParaRPr pararpr55 = wmlObjectFactory.createParaRPr();
        ppr55.setRPr(pararpr55);
        // Create object for color
        Color color134 = wmlObjectFactory.createColor();
        pararpr55.setColor(color134);
        color134.setVal("365F91");
        color134.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color134.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle31 = wmlObjectFactory.createPPrBasePStyle();
        ppr55.setPStyle(pprbasepstyle31);
        pprbasepstyle31.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind8 = wmlObjectFactory.createPPrBaseInd();
        ppr55.setInd(pprbaseind8);
        pprbaseind8.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p56 = wmlObjectFactory.createP();
        body.getContent().add(p56);
        // Create object for pPr
        PPr ppr56 = wmlObjectFactory.createPPr();
        p56.setPPr(ppr56);
        // Create object for rPr
        ParaRPr pararpr56 = wmlObjectFactory.createParaRPr();
        ppr56.setRPr(pararpr56);
        // Create object for color
        Color color135 = wmlObjectFactory.createColor();
        pararpr56.setColor(color135);
        color135.setVal("365F91");
        color135.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color135.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle32 = wmlObjectFactory.createPPrBasePStyle();
        ppr56.setPStyle(pprbasepstyle32);
        pprbasepstyle32.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind9 = wmlObjectFactory.createPPrBaseInd();
        ppr56.setInd(pprbaseind9);
        pprbaseind9.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p57 = wmlObjectFactory.createP();
        body.getContent().add(p57);
        // Create object for pPr
        PPr ppr57 = wmlObjectFactory.createPPr();
        p57.setPPr(ppr57);
        // Create object for rPr
        ParaRPr pararpr57 = wmlObjectFactory.createParaRPr();
        ppr57.setRPr(pararpr57);
        // Create object for color
        Color color136 = wmlObjectFactory.createColor();
        pararpr57.setColor(color136);
        color136.setVal("365F91");
        color136.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color136.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle33 = wmlObjectFactory.createPPrBasePStyle();
        ppr57.setPStyle(pprbasepstyle33);
        pprbasepstyle33.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind10 = wmlObjectFactory.createPPrBaseInd();
        ppr57.setInd(pprbaseind10);
        pprbaseind10.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p58 = wmlObjectFactory.createP();
        body.getContent().add(p58);
        // Create object for pPr
        PPr ppr58 = wmlObjectFactory.createPPr();
        p58.setPPr(ppr58);
        // Create object for rPr
        ParaRPr pararpr58 = wmlObjectFactory.createParaRPr();
        ppr58.setRPr(pararpr58);
        // Create object for color
        Color color137 = wmlObjectFactory.createColor();
        pararpr58.setColor(color137);
        color137.setVal("365F91");
        color137.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color137.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle34 = wmlObjectFactory.createPPrBasePStyle();
        ppr58.setPStyle(pprbasepstyle34);
        pprbasepstyle34.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind11 = wmlObjectFactory.createPPrBaseInd();
        ppr58.setInd(pprbaseind11);
        pprbaseind11.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p59 = wmlObjectFactory.createP();
        body.getContent().add(p59);
        // Create object for pPr
        PPr ppr59 = wmlObjectFactory.createPPr();
        p59.setPPr(ppr59);
        // Create object for rPr
        ParaRPr pararpr59 = wmlObjectFactory.createParaRPr();
        ppr59.setRPr(pararpr59);
        // Create object for color
        Color color138 = wmlObjectFactory.createColor();
        pararpr59.setColor(color138);
        color138.setVal("365F91");
        color138.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color138.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle35 = wmlObjectFactory.createPPrBasePStyle();
        ppr59.setPStyle(pprbasepstyle35);
        pprbasepstyle35.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind12 = wmlObjectFactory.createPPrBaseInd();
        ppr59.setInd(pprbaseind12);
        pprbaseind12.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p60 = wmlObjectFactory.createP();
        body.getContent().add(p60);
        // Create object for pPr
        PPr ppr60 = wmlObjectFactory.createPPr();
        p60.setPPr(ppr60);
        // Create object for rPr
        ParaRPr pararpr60 = wmlObjectFactory.createParaRPr();
        ppr60.setRPr(pararpr60);
        // Create object for color
        Color color139 = wmlObjectFactory.createColor();
        pararpr60.setColor(color139);
        color139.setVal("365F91");
        color139.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color139.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle36 = wmlObjectFactory.createPPrBasePStyle();
        ppr60.setPStyle(pprbasepstyle36);
        pprbasepstyle36.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind13 = wmlObjectFactory.createPPrBaseInd();
        ppr60.setInd(pprbaseind13);
        pprbaseind13.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p61 = wmlObjectFactory.createP();
        body.getContent().add(p61);
        // Create object for pPr
        PPr ppr61 = wmlObjectFactory.createPPr();
        p61.setPPr(ppr61);
        // Create object for rPr
        ParaRPr pararpr61 = wmlObjectFactory.createParaRPr();
        ppr61.setRPr(pararpr61);
        // Create object for color
        Color color140 = wmlObjectFactory.createColor();
        pararpr61.setColor(color140);
        color140.setVal("365F91");
        color140.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color140.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle37 = wmlObjectFactory.createPPrBasePStyle();
        ppr61.setPStyle(pprbasepstyle37);
        pprbasepstyle37.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind14 = wmlObjectFactory.createPPrBaseInd();
        ppr61.setInd(pprbaseind14);
        pprbaseind14.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p62 = wmlObjectFactory.createP();
        body.getContent().add(p62);
        // Create object for pPr
        PPr ppr62 = wmlObjectFactory.createPPr();
        p62.setPPr(ppr62);
        // Create object for rPr
        ParaRPr pararpr62 = wmlObjectFactory.createParaRPr();
        ppr62.setRPr(pararpr62);
        // Create object for color
        Color color141 = wmlObjectFactory.createColor();
        pararpr62.setColor(color141);
        color141.setVal("365F91");
        color141.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color141.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle38 = wmlObjectFactory.createPPrBasePStyle();
        ppr62.setPStyle(pprbasepstyle38);
        pprbasepstyle38.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind15 = wmlObjectFactory.createPPrBaseInd();
        ppr62.setInd(pprbaseind15);
        pprbaseind15.setRight(BigInteger.valueOf(-1238));
        // Create object for p
        P p63 = wmlObjectFactory.createP();
        body.getContent().add(p63);
        // Create object for pPr
        PPr ppr63 = wmlObjectFactory.createPPr();
        p63.setPPr(ppr63);
        // Create object for rPr
        ParaRPr pararpr63 = wmlObjectFactory.createParaRPr();
        ppr63.setRPr(pararpr63);
        // Create object for color
        Color color142 = wmlObjectFactory.createColor();
        pararpr63.setColor(color142);
        color142.setVal("365F91");
        color142.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
        color142.setThemeShade("BF");
        // Create object for pStyle
        PPrBase.PStyle pprbasepstyle39 = wmlObjectFactory.createPPrBasePStyle();
        ppr63.setPStyle(pprbasepstyle39);
        pprbasepstyle39.setVal("BodyText");
        // Create object for ind
        PPrBase.Ind pprbaseind16 = wmlObjectFactory.createPPrBaseInd();
        ppr63.setInd(pprbaseind16);
        pprbaseind16.setRight(BigInteger.valueOf(-1238));
        // Create object for sectPr
        SectPr sectpr = wmlObjectFactory.createSectPr();
        body.setSectPr(sectpr);
        // Create object for pgSz
        // CTDocGrid grid = wmlObjectFactory.createCTDocGrid();
        // grid.setLinePitch(BigInteger.valueOf(360));
        // sectpr.setDocGrid(grid);
        SectPr.PgSz sectprpgsz = wmlObjectFactory.createSectPrPgSz();
        sectpr.setPgSz(sectprpgsz);
        sectprpgsz.setW(BigInteger.valueOf(12240));
        sectprpgsz.setH(BigInteger.valueOf(15840));
        // Create object for pgMar
        SectPr.PgMar sectprpgmar = wmlObjectFactory.createSectPrPgMar();
        sectpr.setPgMar(sectprpgmar);
        sectprpgmar.setTop(BigInteger.valueOf(1440));
        sectprpgmar.setRight(BigInteger.valueOf(1797));
        sectprpgmar.setBottom(BigInteger.valueOf(1440));
        sectprpgmar.setLeft(BigInteger.valueOf(2892));
        sectprpgmar.setHeader(BigInteger.valueOf(720));
        sectprpgmar.setFooter(BigInteger.valueOf(720));
        sectprpgmar.setGutter(BigInteger.valueOf(0));
        // Create object for cols
        CTColumns columns = wmlObjectFactory.createCTColumns();
        sectpr.setCols(columns);
        columns.setSpace(BigInteger.valueOf(720));
        document.setIgnorable("w14 w15 w16se w16cid wp14");

        return document;
    }

    /**
     * Helper method to created a numbered list in the word doc iteratively
     * @param items A List of type String that denotes the items (in order) that are to be listed
     * @param baseTc the Tc that will be used as the parent for the child items that will be added.
     */
    private static void addNumberedList(List<String> items, Tc baseTc){
        org.docx4j.wml.ObjectFactory wmlObjectFactory = new org.docx4j.wml.ObjectFactory();
        for(String item : items){
            P p = wmlObjectFactory.createP();
            R r = wmlObjectFactory.createR();
            Text text = wmlObjectFactory.createText();
            JAXBElement<org.docx4j.wml.Text> textWrapped = wmlObjectFactory.createRT(text);
            RPr rpr = wmlObjectFactory.createRPr();
            RFonts rfonts = wmlObjectFactory.createRFonts();
            BooleanDefaultTrue booleandefaulttrue = wmlObjectFactory.createBooleanDefaultTrue();
            rpr.setBCs(booleandefaulttrue);
            rpr.setRFonts(rfonts);
            rfonts.setAscii("Arial");
            rfonts.setHAnsi("Arial");
            rfonts.setCs("Arial");
            Color color = wmlObjectFactory.createColor();
            rpr.setColor(color);
            color.setVal("365F91");
            color.setThemeColor(org.docx4j.wml.STThemeColor.ACCENT_1);
            color.setThemeShade("BF");
            
            

            text.setValue(item);
            r.getContent().add(textWrapped);
            r.setRPr(rpr);
            p.getContent().add(r);
            baseTc.getContent().add(p);

            PPr ppr = wmlObjectFactory.createPPr();
            
            ParaRPr pararpr = wmlObjectFactory.createParaRPr();
            p.setPPr(ppr);
            ppr.setRPr(pararpr);
            pararpr.setColor(color);
            pararpr.setB(wmlObjectFactory.createBooleanDefaultTrue());
            PPrBase.PStyle pprbasepstyle = wmlObjectFactory.createPPrBasePStyle();
            pprbasepstyle.setVal("ListParagraph");

            PPrBase.NumPr pprbasenumpr = wmlObjectFactory.createPPrBaseNumPr();
            ppr.setNumPr(pprbasenumpr);
            PPrBase.NumPr.Ilvl pprbasenumprilvl = wmlObjectFactory.createPPrBaseNumPrIlvl();
            pprbasenumpr.setIlvl(pprbasenumprilvl);
            pprbasenumprilvl.setVal(BigInteger.valueOf(0));

            PPrBase.NumPr.NumId pprbasenumprnumid = wmlObjectFactory.createPPrBaseNumPrNumId();
            pprbasenumpr.setNumId(pprbasenumprnumid);
            pprbasenumprnumid.setVal(BigInteger.valueOf(43));
        }
    }


    /**
     * Private method to help add blue-coloured text to a P
     * @param baseP - The P object you'd like to add text to
     * @param textToAdd - The string of text you'd like to add to the P
     */
    private static void addToP(P baseP, String textToAdd){
        org.docx4j.wml.ObjectFactory wmlObjectFactory = new org.docx4j.wml.ObjectFactory();
        R r6 = wmlObjectFactory.createR();
        Text text6 = wmlObjectFactory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped6 = wmlObjectFactory.createRT(text6);
        r6.getContent().add(textWrapped6);
        if (textToAdd == null){
            textToAdd = " ";
        }else {
            //I want a starting space between the previous run and this text
            textToAdd = " " + textToAdd;
        }

        text6.setValue(textToAdd);
        // Create object for rPr which stands for "Run Property", defines the properties for the run
        RPr rpr6 = wmlObjectFactory.createRPr();
        r6.setRPr(rpr6);

        Color color7 = wmlObjectFactory.createColor();
        rpr6.setColor(color7);
        color7.setVal("1F497D");
        color7.setThemeColor(org.docx4j.wml.STThemeColor.TEXT_2);
        // Create object for pPr which stands for Paragraph Property, defines the properties for the paragraph
        PPr ppr3 = wmlObjectFactory.createPPr();
        // baseP.setPPr(ppr3);
        // Create object for rPr
        ParaRPr pararpr3 = wmlObjectFactory.createParaRPr();
        ppr3.setRPr(pararpr3);
        
        RFonts rfonts8 = wmlObjectFactory.createRFonts();
        pararpr3.setRFonts(rfonts8);
        rpr6.setRFonts(rfonts8);
        rfonts8.setAscii("Arial");
        rfonts8.setHAnsi("Arial");
        rfonts8.setCs("Arial");
        baseP.getContent().add(r6);
    }
}

