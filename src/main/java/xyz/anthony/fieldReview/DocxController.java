package xyz.anthony.fieldReview;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;



@RestController
public class DocxController {
    @Autowired
    private FieldReviewDataRepository fieldReviewDataRepository;
    private final org.slf4j.Logger logger = LoggerFactory.getLogger(this.getClass());


    @PostMapping(value="/docBuild", produces = "application/vnd.openxmlformats-officedocument.wordprocessingml.document", consumes = "multipart/form-data")
    public byte[] createFile(@ModelAttribute FieldReviewData reviewData, @RequestPart("files") MultipartFile []files) {
        //TODO: process POST request
        DocxBuilder docBuilder = new DocxBuilder(reviewData);
        logger.info("Files length: " + files.length);
        logger.info("Review Data: " + reviewData);
        File docFile = docBuilder.buildDoc(reviewData.getFileNumber() + ".docx", converMultiPartFiletoFile(files));

        if(!fieldReviewDataRepository.existsByProjectAddress(reviewData.getProjectAddress())){    
            fieldReviewDataRepository.save(reviewData);
        }
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
    
    private List<File> converMultiPartFiletoFile(MultipartFile []multiFiles){
        List<File> fileList = new ArrayList<File>();

        for (int i = 0; i < multiFiles.length; i++){
            File file = new File(multiFiles[i].getOriginalFilename());
            try{
                logger.info("INFO ON FILES RIGHT NOW");
                logger.info("Multi-file is empty: " + multiFiles[i].isEmpty());
                logger.info("Multi-file size: " + multiFiles[i].getSize());
                logger.info("Multi-file Content type" + multiFiles[i].getContentType());

                // file.createNewFile();
                logger.info("" + file.exists());
                multiFiles[i].transferTo(file.toPath());
            }catch(IOException e){
                logger.error(e.getMessage(), e.fillInStackTrace());
            }
            fileList.add(file);
        }

        return fileList;
    }

    @GetMapping("/profiles")
    public Iterable<FieldReviewData> getMethodName() {
        return fieldReviewDataRepository.findAll();
    }
    

}
