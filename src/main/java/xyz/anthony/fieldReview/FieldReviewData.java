package xyz.anthony.fieldReview;

import java.util.ArrayList;
import java.util.List;

import org.springframework.web.multipart.MultipartFile;

import lombok.Data;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;

@Data
@Entity
public class FieldReviewData {
    
    private List<String> datesVisited = new ArrayList<String>();
    private List<String> inspectionNotes = new ArrayList<String>();

    @Id
    @GeneratedValue(strategy=GenerationType.AUTO)
    private Integer id;

    private String fileNumber;
    private String date;
    private String location;
    private String referenceDwgs;
    private String projectName;
    private String builder;
    private String weatherCondition;
    private String inspectionCategory;
    private String reportNumber;
    private String commonElementNumber;
    private String projectAddress;
    // private MultipartFile files[];

    public void addToNumberedList(String listItem){
        inspectionNotes.add(listItem);
    }

    public void addToDateList(String dateItem){
        datesVisited.add(dateItem);
    }
}
