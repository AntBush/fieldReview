package xyz.anthony.fieldReview;

import java.util.ArrayList;
import java.util.List;

import lombok.Data;

@Data
public class FieldReviewData {
    private List<String> dateList = new ArrayList<String>();
    private List<String> numberedList = new ArrayList<String>();

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

    public void addToNumberedList(String listItem){
        numberedList.add(listItem);
    }

    public void addToDateList(String dateItem){
        dateList.add(dateItem);
    }
}
