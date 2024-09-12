import { Component, ElementRef, ViewChild } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import * as XLSX from 'xlsx';  // For reading Excel
import { Document, Packer, PageBreak, Paragraph, TextRun } from 'docx';  // For generating Word
import { saveAs } from 'file-saver';  // For downloading files

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  title = 'ExcelToWordConverter';
  @ViewChild('fileInput') fileInput!: ElementRef;

  // Function to handle file input (triggered when user uploads an Excel file)
  onFileChange(event: any) {
    const file = event.target.files[0];

    if (file) {
      const reader = new FileReader();
      reader.onload = (e: any) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Parse Excel data to JSON format
        const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Call the function to generate Word files
        this.convertExcelToWord(excelData);
      };
      reader.readAsArrayBuffer(file);
    }
  }

  // Function to convert parsed Excel data to Word files
  async convertExcelToWord(data: any[]) {
    const headings = data[0];  // First row as headings
    const caseStudies = data.slice(1);  // Remaining rows as case studies

    // Array to hold all paragraphs
    const allParagraphs: Paragraph[] = [];

    console.log('cases', caseStudies);

    caseStudies.forEach((caseStudy, index) => {
      // Create paragraphs for the case study
      const paragraphs = [
        new Paragraph({
          children: [new TextRun({ text: `Case Study #${index + 1}`, bold: true, size: 28 })],
        }),
        ...headings.map((heading: any, i: number) => {
          const text = caseStudy[i] !== undefined && caseStudy[i] !== null ? String(caseStudy[i]) : 'N/A';  // Handle empty fields and convert all values to strings
          console.log('heading', heading);
          console.log('text', text);
          return new Paragraph({
            children: [
              new TextRun({ text: `${heading}: `, bold: true, size: 24 }),
              new TextRun({ text, size: 24 }),
              new TextRun({ text: '', break: 1 })
            ],
          });
        }),
      ];

      // Add paragraphs to the allParagraphs array
      allParagraphs.push(...paragraphs);

      // Add a page break after each case study except the last one
      if (index < caseStudies.length - 1) {
        allParagraphs.push(new Paragraph({ children: [new PageBreak()] }));
      }
    });

    // Create a new Word document with all the paragraphs
    const doc = new Document({
      sections: [{
        properties: {},
        children: allParagraphs,
      }],
    });

    // Generate the Word file as a blob
    Packer.toBlob(doc).then(blob => {
      // Use file-saver to download the generated Word file
      saveAs(blob, `CaseStudies.docx`);
      this.fileInput.nativeElement.value = '';
    });
  }

}
