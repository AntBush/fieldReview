package xyz.anthony.fieldReview;

import java.io.File;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
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
			Hdr header = objectFactory.createHdr();
			Text text = objectFactory.createText();

			text.setValue("asdf");
			header.getContent().add(text);
			hdrPart.setContents(header);
			mdp.addTargetPart(new HeaderPart());
			File exportFile = new File("welcome.docx");
			wordMLPackage.save(exportFile);
			
		}catch (Exception e){
			logger.error(e.getMessage());
			
		}
	}

}