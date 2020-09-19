import { Component, OnInit } from '@angular/core';
import { FormControl, FormGroup, FormBuilder, FormArray, Validators} from "@angular/forms";
import { HttpClient, HttpHeaders } from "@angular/common/http";
import { map } from "rxjs/operators";

@Component({
  selector: 'app-docx-form',
  templateUrl: './docx-form.component.html',
  styleUrls: ['./docx-form.component.css']
})
export class DocxFormComponent implements OnInit {
    constructor(private fb: FormBuilder, private http: HttpClient) { }
    formData = new FormData;
    filesToUpload: FileList;
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
    });
   

  ngOnInit(): void {
  }

  sendHttpRequest(){
    for (const key in this.docxForm.value) {
        if (Object.prototype.hasOwnProperty.call(this.docxForm.value, key)) {
                    const element = this.docxForm.get(key).value;
                    console.log(key);
                    console.log(element);
                    this.formData.append(key,element);
                    // newObject[key] = element;
        }
    }
    let httpHeader = new HttpHeaders({"Content-Type" : "multipart/form-data"});
    // httpHeader.set()
    for(let i = 0; i < this.filesToUpload.length; i++){
        this.formData.append("files",this.filesToUpload[i]);
    }
    
    this.http.post("/docBuild",this.formData,{responseType: 'arraybuffer'}).subscribe(response => {
    this.downLoadFile(response, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    this.formData = new FormData();
   })  
  }
  handleFileInput(files: FileList) {
    this.filesToUpload = files;
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
