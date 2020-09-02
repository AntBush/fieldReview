package xyz.anthony.fieldReview;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import lombok.NoArgsConstructor;

@NoArgsConstructor
public class DocxBuilder {
    private WordprocessingMLPackage wordMLPackage;
    

    public void doShit(){
        System.out.println("");
        try{
            wordMLPackage = WordprocessingMLPackage.createPackage();
            

        }catch (Exception e){

        }
    }

}