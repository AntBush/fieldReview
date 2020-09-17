import { Component, OnInit } from '@angular/core';
import { FormControl, FormGroup, FormBuilder, FormArray, Validators} from "@angular/forms";
import { HttpClient } from "@angular/common/http";
import { map } from "rxjs/operators";

@Component({
  selector: 'app-docx-form',
  templateUrl: './docx-form.component.html',
  styleUrls: ['./docx-form.component.css']
})
export class DocxFormComponent implements OnInit {
    constructor(private fb: FormBuilder, private http: HttpClient) { }
    docxForm = this.fb.group({
        reportNumber:[''],
        commonElementNumber:[''],
        fileNumber: [''],
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
        ]),
        files: [''],
    });
   

  ngOnInit(): void {
  }

  sendHttpRequest(){
   this.http.post("/docBuild",this.docxForm.value,{responseType: 'arraybuffer'}).subscribe(response => {
    this.downLoadFile(response, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
   })  
  }

  logIt(){
    //   console.log(this.datesVisited.value);
      console.log(this.docxForm.value);
  }

  private downLoadFile(data: any, type: string){
    let blob = new Blob([data], { type: type});
    let url = window.URL.createObjectURL(blob);
    let pwa = window.open(url);
    if (!pwa || pwa.closed || typeof pwa.closed == 'undefined') {
        alert( 'Please disable your Pop-up blocker and try again.');
    }
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
