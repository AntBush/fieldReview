package xyz.anthony.fieldReview;

import java.io.File;
import java.util.List;

import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Text;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import lombok.NoArgsConstructor;

@NoArgsConstructor
public class DocxBuilder  {
	private WordprocessingMLPackage wordMLPackage;
	private ObjectFactory objectFactory = new ObjectFactory();
	private Logger logger = LoggerFactory.getLogger(this.getClass());


	public void doShit(){
		try{
			logger.info("Creating Word package");
			wordMLPackage = WordprocessingMLPackage.createPackage();
			logger.info("Error here?");
			MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
			HeaderPart hdrPart = new HeaderPart();
			
			hdrPart.setContents(HeaderBuilder.createIt());
			Relationship rel = mdp.addTargetPart(hdrPart);
			createHeaderReference(wordMLPackage, rel);
			File exportFile = new File("welcome.docx");
			wordMLPackage.save(exportFile);
			
		}catch (Exception e){
			logger.error(e.getMessage());
			
		}
	}

	private void createHeaderReference(WordprocessingMLPackage mlPackage, Relationship relationship){
		
		List<SectionWrapper> sections = mlPackage.getDocumentModel().getSections();
		   
		SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
		// There is always a section wrapper, but it might not contain a sectPr
		if (sectPr==null ) {
			sectPr = objectFactory.createSectPr();
			mlPackage.getMainDocumentPart().addObject(sectPr);
			sections.get(sections.size() - 1).setSectPr(sectPr);
		}

		HeaderReference headerReference = objectFactory.createHeaderReference();
		headerReference.setId(relationship.getId());
		headerReference.setType(HdrFtrRef.DEFAULT);
		sectPr.getEGHdrFtrReferences().add(headerReference);// add header or
		// footer references
	}

}