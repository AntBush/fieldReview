package xyz.anthony.fieldReview;

import java.io.File;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.DocDefaults;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
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
    private List<String> numberedList = new ArrayList<String>();
    private List<String> dateList = new ArrayList<String>();
    
    @Getter
    @Setter
    private String fileNumber;
    @Getter
    @Setter
    private String date;
    @Getter
    @Setter
    private String location;
    @Getter
    @Setter
    private String referenceDwgs;
    @Getter
    @Setter
    private String projectName;
    @Getter
    @Setter
    private String builder;
    @Getter
    @Setter
    private String weatherCondition;
    @Getter
    @Setter
    private String inspectionCategory;

	public void doShit(){
		try{
			logger.info("Creating Word package");
			wordMLPackage = WordprocessingMLPackage.createPackage();
			logger.info("Error here?");
			HeaderPart hdrPart = new HeaderPart();
			MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
            numberedList.add("Precast one");
            numberedList.add("Item two");
            numberedList.add("Item Three");
			mdp.setContents(BodyWithTableBuilder.createIt(numberedList));
			// StyleDefinitionsPart asdf = objectFactory.style
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
			Relationship rel = mdp.addTargetPart(hdrPart);
			for(SectionWrapper section : sections){
				logger.info("Section is: " + section.getSectPr().toString());
			}
			createHeaderReference(wordMLPackage, rel);
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

}