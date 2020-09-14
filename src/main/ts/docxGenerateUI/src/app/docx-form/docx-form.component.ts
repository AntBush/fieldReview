import { Component, OnInit } from '@angular/core';
import { FormControl, FormGroup, FormBuilder, FormArray, Validators} from "@angular/forms";

@Component({
  selector: 'app-docx-form',
  templateUrl: './docx-form.component.html',
  styleUrls: ['./docx-form.component.css']
})
export class DocxFormComponent implements OnInit {
    constructor(private fb: FormBuilder) { }
    docxForm = this.fb.group({
        fieldReviewReports:[''],
        commonElementNo:[''],
        fileNo: [''],
        date: [''],
        projectAddress:[''],
        location:[''],
        referenceDwgs:[''],
        projectName:[''],
        builder:[''],
        weatherCondition:[''],
        inspectionCategory:[''],
        datesVisited: this.fb.array([
            this.fb.control('')
        ]),
        inspectionNotes: this.fb.array([
            this.fb.control('')
        ])
    });
   

  ngOnInit(): void {
  }

  logIt(){
      console.log(this.datesVisited.value);
      console.log(this.docxForm.value);
  }

  get datesVisited() {
      return this.docxForm.get('datesVisited') as FormArray;
  }

  addDateVisited() {
      console.log("Why no work");
      this.datesVisited.push(this.fb.control(''));
  }

  get inspectionNotes(){
      return this.docxForm.get('inspectionNotes') as FormArray;
  }

  addInspectionNotes(){
      this.inspectionNotes.push(this.fb.control(''));
  }


}
