import { Component, OnInit } from '@angular/core';

declare const Office;
@Component({
  selector: 'app-settings',
  templateUrl: './settings.component.html',
  styleUrls: ['./settings.component.scss']
})
export class SettingsComponent implements OnInit {
  OUTLOOK_URL = 'https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js';
  loaded = false;
  constructor() { }

  loadScript(url): HTMLScriptElement {
    const node = document.createElement('script');
    node.src = url;
    node.type = 'text/javascript';
    node.async = false;
    node.charset = 'utf-8';
    document.getElementsByTagName('head')[0].appendChild(node);
    return node;
  }

  isScriptLoaded(url) {
    if (document.querySelector(`script[src="${url}"]`)) {
      return true;
    }
    return false;
  }

  ngOnInit(): void {
    this.loadScript(this.OUTLOOK_URL).onload = () => {
      this.initializeOutlook();
    };
  }

  initializeOutlook() {
    // If the office context is not present, then call onReady which initializes the office context
    Office.onReady().then(() => {
      this.onOfficeReady();
    });
  }

  onOfficeReady() {
    console.log('Office is ready. Further logic after office has loaded');
    this.loaded = true;
  }
}
