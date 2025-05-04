import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { Modal } from 'bootstrap'; // ✅ CORRECT way

@Component({
  selector: 'app-sql-script',
  templateUrl: './sql-script.component.html'
})
export class SqlScriptComponent implements OnInit {
  allScripts: any[] = [];
  filteredScripts: any[] = [];
  newScript = { Key: '', Script: '', Description: '' };
  isEditMode = false;
  editingKey = '';
  searchKey = '';  // ✅ Ensure this is defined

  ngOnInit(): void {
    this.loadExcel();
  }

  loadExcel(): void {
    fetch('assets/scripts.xlsx')
      .then(res => res.arrayBuffer())
      .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        this.allScripts = XLSX.utils.sheet_to_json(sheet);
        this.filteredScripts = [...this.allScripts];
      });
  }

  filter(): void {
    const keyword = this.searchKey.toLowerCase();
    this.filteredScripts = this.allScripts.filter(row =>
      row.Key?.toLowerCase().includes(keyword) ||
      row.Script?.toLowerCase().includes(keyword) ||
      row.Description?.toLowerCase().includes(keyword)
    );
  }

  openAddModal(): void {
    this.isEditMode = false;
    this.newScript = { Key: '', Script: '', Description: '' };
    const modalElement = document.getElementById('scriptModal');
    const modal = new Modal(modalElement!);
    modal.show();
  }

  editScript(script: any): void {
    this.isEditMode = true;
    this.newScript = { ...script };
    this.editingKey = script.Key;
    const modalElement = document.getElementById('scriptModal');
    const modal = new Modal(modalElement!);
    modal.show();
  }

  submitScript(): void {
    if (this.isEditMode) {
      const index = this.allScripts.findIndex(s => s.Key === this.editingKey);
      if (index !== -1) {
        this.allScripts[index] = { ...this.newScript };
      }
    } else {
      this.allScripts.push({ ...this.newScript });
    }

    this.filter();
    this.closeModal();
  }

  deleteScript(key: string): void {
    if (confirm('Are you sure you want to delete this script?')) {
      this.allScripts = this.allScripts.filter(s => s.Key !== key);
      this.filter();
    }
  }

  exportToExcel(): void {
    const worksheet = XLSX.utils.json_to_sheet(this.allScripts);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Scripts');
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const file = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(file, 'updated_scripts.xlsx');
  }

  private closeModal(): void {
    const modalElement = document.getElementById('scriptModal');
    const modalInstance = Modal.getInstance(modalElement!); // ✅ Correct usage
    modalInstance?.hide();
  }
}
