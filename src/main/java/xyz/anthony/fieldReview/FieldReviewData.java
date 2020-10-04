package xyz.anthony.fieldReview;

import java.util.ArrayList;

import lombok.Data;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Transient;
import javax.persistence.UniqueConstraint;

@Data
@Entity
public class FieldReviewData {
    
    @Transient
    private ArrayList<String> datesVisited = new ArrayList<String>();
    @Transient
    private ArrayList<String> inspectionNotes = new ArrayList<String>();

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
    @Column(unique = true)
    private String projectAddress;
    // private MultipartFile files[];

    public void addToNumberedList(String listItem){
        inspectionNotes.add(listItem);
    }

    public void addToDateList(String dateItem){
        datesVisited.add(dateItem);
    }
}
