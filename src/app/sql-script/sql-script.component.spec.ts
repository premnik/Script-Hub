import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SqlScriptComponent } from './sql-script.component';

describe('SqlScriptComponent', () => {
  let component: SqlScriptComponent;
  let fixture: ComponentFixture<SqlScriptComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ SqlScriptComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SqlScriptComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
