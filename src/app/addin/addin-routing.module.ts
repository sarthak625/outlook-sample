import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';
import { SettingsComponent } from './settings/settings.component';

const routes: Routes = [
  {
    path: '',
    component: SettingsComponent,
    data: { showHeader: false, showFooter: false },
    pathMatch: 'full',
  },
];

@NgModule({
  imports: [RouterModule.forChild(routes)],
  exports: [RouterModule]
})
export class AddinRoutingModule { }
