package xyz.anthony.fieldReview;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;



import org.springframework.web.bind.annotation.RequestParam;


@RestController
public class TestEndpoint {
    private Logger logger = LoggerFactory.getLogger(this.getClass());

    @GetMapping(value="/")
    public String getMethodName() {
        
        logger.info("Instantiating doc class");
        DocxBuilder doc = new DocxBuilder();

        doc.doShit();
        
        return "Test";
    }
    
}
