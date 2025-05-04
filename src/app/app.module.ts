import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { HttpClientModule } from '@angular/common/http';
import { FormsModule } from '@angular/forms';

import { AppComponent } from './app.component';
//import { SqlScriptComponent } from './sql-script/sql-script.component';
import { SqlScriptComponent } from './sql-script/sql-script.component' //'./sql-script/sql-script.component';

@NgModule({
  declarations: [AppComponent, SqlScriptComponent],
  imports: [BrowserModule, HttpClientModule, FormsModule],
  bootstrap: [AppComponent],
})
export class AppModule {}
