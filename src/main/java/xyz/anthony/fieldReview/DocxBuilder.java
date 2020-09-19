package xyz.anthony.fieldReview;

import java.io.File;
import java.math.BigInteger;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.DocDefaults;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.FooterReference;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.R;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.DocDefaults.PPrDefault;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@NoArgsConstructor
public class DocxBuilder  {
	private WordprocessingMLPackage wordMLPackage;
	private ObjectFactory objectFactory = new ObjectFactory();
	private Logger logger = LoggerFactory.getLogger(this.getClass());
    
    @Getter
    @Setter
    private FieldReviewData reviewData = new FieldReviewData();
    
    public DocxBuilder(FieldReviewData reviewdData){
        this.reviewData = reviewdData;
    }

	public File buildDoc(String docName, List<File> images){
		
		File exportFile = new File(docName);
		
		try{
			logger.info("Creating Word package");
			wordMLPackage = WordprocessingMLPackage.createPackage();
			logger.info("Error here?");
            HeaderPart hdrPart = new HeaderPart();
            FooterPart ftrPart = new FooterPart();
            MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
            mdp.setContents(BodyWithTableBuilder.createIt(reviewData));
			DocDefaults docDefault = objectFactory.createDocDefaults();
			PPrDefault pprDefault = objectFactory.createDocDefaultsPPrDefault();
			PPr ppr = objectFactory.createPPr();
			PPrBase.Spacing pprBaseSpacing = objectFactory.createPPrBaseSpacing();
			pprBaseSpacing.setAfter(BigInteger.valueOf(0));
			ppr.setSpacing(pprBaseSpacing);
			pprDefault.setPPr(ppr);
			docDefault.setPPrDefault(pprDefault);
			
			List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
            hdrPart.setContents(HeaderBuilder.createIt());
            ftrPart.setContents(FooterBuilder.createIt());
            Relationship ftrRel = mdp.addTargetPart(ftrPart);
			Relationship hdrRel = mdp.addTargetPart(hdrPart);
			for(SectionWrapper section : sections){
				logger.info("Section is: " + section.getSectPr().toString());
			}
            createHeaderReference(wordMLPackage, hdrRel);
            createFooterReference(wordMLPackage, ftrRel);
            addFiles(images, wordMLPackage,mdp);
            mdp.getStyleDefinitionsPart().getContents().setDocDefaults(docDefault);
            
			wordMLPackage.save(exportFile);
			
		}catch (Exception e){
			logger.error(e.getMessage());
			
		}

		return exportFile;
		
	}
    
    private static P addImageToParagraph(Inline inline) {
	    ObjectFactory factory = new ObjectFactory();
	    P p = factory.createP();
	    R r = factory.createR();
	    p.getContent().add(r);
	    Drawing drawing = factory.createDrawing();
	    r.getContent().add(drawing);
	    drawing.getAnchorOrInline().add(inline);
	    return p;
	}

    private void addFiles(List<File> fileList, WordprocessingMLPackage wordPackage, MainDocumentPart mainDocumentPart){
        for(File file : fileList){
            try{

                byte[] fileContent = Files.readAllBytes(file.toPath());
                BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordPackage, fileContent);
                Inline inline = imagePart.createImageInline("Baeldung Image (filename hint)", "Alt Text", 1, 2, false);
                P Imageparagraph = addImageToParagraph(inline);
                mainDocumentPart.getContent().add(Imageparagraph);
            }catch(Exception e){
                logger.error(e.getMessage(),e.fillInStackTrace());
            }
        }
    }

	public void buildDoc(){
		try{
			logger.info("Creating Word package");
			wordMLPackage = WordprocessingMLPackage.createPackage();
			logger.info("Error here?");
            HeaderPart hdrPart = new HeaderPart();
            FooterPart ftrPart = new FooterPart();
            MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
            reviewData.addToNumberedList("listItem");
            reviewData.addToNumberedList("ListItem 2");
            mdp.setContents(BodyWithTableBuilder.createIt(reviewData));
			DocDefaults docDefault = objectFactory.createDocDefaults();
			PPrDefault pprDefault = objectFactory.createDocDefaultsPPrDefault();
			PPr ppr = objectFactory.createPPr();
			PPrBase.Spacing pprBaseSpacing = objectFactory.createPPrBaseSpacing();
			pprBaseSpacing.setAfter(BigInteger.valueOf(0));
			ppr.setSpacing(pprBaseSpacing);
			pprDefault.setPPr(ppr);
			docDefault.setPPrDefault(pprDefault);
			
			// List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
            hdrPart.setContents(HeaderBuilder.createIt());
            ftrPart.setContents(FooterBuilder.createIt());
            Relationship ftrRel = mdp.addTargetPart(ftrPart);
			Relationship hdrRel = mdp.addTargetPart(hdrPart);
            createHeaderReference(wordMLPackage, hdrRel);
            createFooterReference(wordMLPackage, ftrRel);
			mdp.getStyleDefinitionsPart().getContents().setDocDefaults(docDefault);
			File exportFile = new File("welcome.docx");
			wordMLPackage.save(exportFile);
			
		}catch (Exception e){
			logger.error(e.getMessage());
			
		}
	}

	//This is taken from the "Create Header and Footer" sample
	private void createHeaderReference(WordprocessingMLPackage mlPackage, Relationship relationship){
		
		List<SectionWrapper> sections = mlPackage.getDocumentModel().getSections();
		   
		SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
		// There is always a section wrapper, but it might not contain a sectPr
		if (sectPr==null) {
			logger.info("Section pointer is null apparently");
			sectPr = objectFactory.createSectPr();
			mlPackage.getMainDocumentPart().addObject(sectPr);
			sections.get(sections.size() - 1).setSectPr(sectPr);
		}
		logger.info("Section pointer is: " + sectPr.toString());
		
		HeaderReference headerReference = objectFactory.createHeaderReference();
		logger.info("Header relationship ID: " + relationship.getId());
		headerReference.setId(relationship.getId());
		headerReference.setType(HdrFtrRef.DEFAULT);
		sectPr.getEGHdrFtrReferences().add(headerReference);// add header or
		sectPr.getEGHdrFtrReferences().add(headerReference);
		sectPr.getEGHdrFtrReferences().remove(0);
		logger.info("Header references: "  + sectPr.getEGHdrFtrReferences().size());
		// footer references
		
    }
    
    private void createFooterReference(WordprocessingMLPackage mlPackage, Relationship relationship){
		
		List<SectionWrapper> sections = mlPackage.getDocumentModel().getSections();
		   
		SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
		// There is always a section wrapper, but it might not contain a sectPr
		if (sectPr==null) {
			logger.info("Section pointer is null apparently");
			sectPr = objectFactory.createSectPr();
			mlPackage.getMainDocumentPart().addObject(sectPr);
			sections.get(sections.size() - 1).setSectPr(sectPr);
		}
		logger.info("Section pointer is: " + sectPr.toString());
		
		FooterReference footerReference = objectFactory.createFooterReference();
		logger.info("Footer relationship ID: " + relationship.getId());
		footerReference.setId(relationship.getId());
		footerReference.setType(HdrFtrRef.DEFAULT);
		sectPr.getEGHdrFtrReferences().add(footerReference);// add header or
		// sectPr.getEGHdrFtrReferences().add(footerReference);
		// sectPr.getEGHdrFtrReferences().remove(mlPackage.getDocumentModel().getSections().size()-1);
		logger.info("Footer references: "  + sectPr.getEGHdrFtrReferences().size());
		// footer references
		
	}

}