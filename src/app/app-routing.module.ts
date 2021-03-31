import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';

const routes: Routes = [
  {
    path: 'outlookHello',
    loadChildren: () =>
      import('./addin/addin.module').then(
        (m) => m.AddinModule
      ),
    data: { showHeader: false, showFooter: false }
  },
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
