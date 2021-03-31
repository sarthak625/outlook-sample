import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { AddinRoutingModule } from './addin-routing.module';
import { SettingsComponent } from './settings/settings.component';


@NgModule({
  declarations: [SettingsComponent],
  imports: [
    CommonModule,
    AddinRoutingModule
  ]
})
export class AddinModule { }
