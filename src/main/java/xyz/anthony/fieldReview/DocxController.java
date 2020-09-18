package xyz.anthony.fieldReview;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.commons.io.IOUtils;
import org.slf4j.LoggerFactory;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;


@RestController
public class DocxController {
    private final org.slf4j.Logger logger = LoggerFactory.getLogger(this.getClass());

    @PostMapping(value="/docBuild", produces = "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    public byte[] createFile(@RequestBody FieldReviewData reviewData, @RequestPart("files") MultipartFile []files) {
        //TODO: process POST request
        DocxBuilder docBuilder = new DocxBuilder(reviewData);

        File docFile = docBuilder.buildDoc(reviewData.getFileNumber() + ".docx");

    
        byte[] byteArray = null;
        try(InputStream in = new FileInputStream(docFile)){
            
            byteArray = IOUtils.toByteArray(in);
        }catch(FileNotFoundException e){
            logger.error(e.getMessage(), e.getCause());   
        }catch(IOException e){
            logger.error(e.getMessage());
            logger.error(e.getMessage(), e.getCause());
        }
        


        return byteArray;
    }
    

}
