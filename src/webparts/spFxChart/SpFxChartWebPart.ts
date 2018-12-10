import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  // PropertyPaneCheckbox,

} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxChartWebPart.module.scss';
import * as strings from 'SpFxChartWebPartStrings';
 import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

import MockHttpClient from './MockHttpClient';
 import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
// //npm install chart.js --save
 import Chart from 'chart.js';

export interface ISpFxChartWebPartProps {
  description: string;
  testprop: string;
  company: string;
  isActive: boolean;
  like: boolean;
  multiLineDesc: string;
  chartType: string;
}
export interface IProjectItem {
  Id: string;
  Title: string;
  TeamSize: number;
}
export interface IProjects {
  value: IProjectItem[];
}

export default class SpFxChartWebPart extends BaseClientSideWebPart<ISpFxChartWebPartProps> {

  public render(): void {
    debugger;

    this.domElement.innerHTML = `
      <div class="${ styles.spFxChart}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <p style="white-space:pre-wrap">${escape(this.properties.testprop)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button}">
                <span class="${ styles.label}">Learn more</span>
              </a>
              
            </div>
          </div>
          <div class=${styles.multi}>${escape(this.properties.multiLineDesc)}</div>
      <div id="myProjects" class="ms-Grid"></div>  
      <canvas id="myChart"></canvas>
        </div>
      </div>`;


     this.renderProjectsAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('testprop', {
                  label: "TEst Property",
                  multiline: true

                }),
                PropertyPaneTextField("multiLineDesc", {
                  label: "Testing new property",
                  multiline: true,
                }),
                PropertyPaneCheckbox('isActive', {
                  text: "Is it Active?",

                }),
                PropertyPaneDropdown('company', {

                  label: "Best Company",
                  options: [
                    { key: 1, text: 'Google' },
                    { key: 2, text: 'Microsoft' },
                    { key: 3, text: 'Amazon' }
                  ]
                }),
                PropertyPaneToggle('like', {
                  label: 'Do you like it?',
                  onText: "Yes, I do",
                  offText: "No, I don't"
                }),
                PropertyPaneDropdown('chartType', {
                  label: 'Type of chart',
                  options: [{ key: '1', text: 'pie' },
                  { key: '2', text: 'doughnut' }]
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  private _getMockProjects(): Promise<IProjects> {

    return MockHttpClient.get()
      .then((mockData) => {
        let allProjects: IProjects = { value: mockData };
        return allProjects;
      });
  }
  private renderProjectsSimple(items: IProjectItem[]) {
    let html: string = '';
    html += `
          <div class="ms-Grid-row">
              <div class="ms-Grid-col ms-lg4">ID</div>
              <div class="ms-Grid-col ms-lg4">Project Name</div>
              <div class="ms-Grid-col ms-lg4">Team Size</div>
          </div>
          `;

    items.forEach((item) => {
      html += `
          <div class="ms-Grid-row">    
              <div class="ms-Grid-col ms-lg4">${item.Id}</div>
              <div class="ms-Grid-col ms-lg4">${item.Title}</div>
              <div class="ms-Grid-col ms-lg4">${item.TeamSize}</div>
          </div>
    `;
    });
    // let projects: Element = this.domElement.querySelector('#myProjects');
    // projects.innerHTML = html;
    let projectsDiv: Element = this.domElement.querySelector('#myProjects');
    projectsDiv.innerHTML = html;
  }
  private renderProjectsAsync(): void {

    if (Environment.type == EnvironmentType.Local) {
      this._getMockProjects()
        .then((data) => {
          this.renderProjects(data.value);
          //this.renderChartStatic();
          this.renderChart(this.properties.chartType,data.value);
        });
    }
    else if (Environment.type == EnvironmentType.SharePoint) {
      this._getProjectsFromSP().then((data) => {
        this.renderProjects(data.value);
        this.renderChart(this.properties.chartType, data.value);
      });
    }

  }



  private renderProjects(items: IProjectItem[]) {
    let html: string = '';
    let altStyle: string = "";

    html += `
  <div class="ms-Grid-row ${styles.gridHeader}">
              <div class="ms-Grid-col ms-hiddenMdDown ms-lg4">ID</div>
              <div class="ms-Grid-col ms-sm12 ms-md6 ms-lg4">Project Name</div>
              <div class="ms-Grid-col ms-hiddenSm ms-md6 ms-lg4">Team Size</div>
    </div>
  `;

    items.forEach((item) => {
      altStyle = altStyle === "" ? styles.altRow : "";
      html += `
    <div class="ms-Grid-row ${altStyle}">
              <div class="ms-Grid-col ms-hiddenMdDown ms-lg4">${item.Id}</div>
              <div class="ms-Grid-col ms-sm12 ms-md6 ms-lg4">${item.Title}</div>
              <div class="ms-Grid-col ms-hiddenSm ms-md6 ms-lg4">${item.TeamSize}</div>
    </div>
    `;
    });
    let projectsDiv: Element = this.domElement.querySelector('#myProjects');
    projectsDiv.innerHTML = html;
  }
  private _getProjectsFromSP(): Promise<IProjects> {
    
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + '/_api/web/lists/GetByTitle(\'Projects\')/items', SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      });
  }
  private renderChartStatic(): void {
    var ctx = document.getElementById("myChart");
    var myChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: ["Red", "Blue", "Yellow", "Green", "Purple", "Orange"],
        datasets: [{
          label: '# of Votes',
          data: [12, 19, 3, 5, 2, 3],
          backgroundColor: [
            'rgba(255, 99, 132, 0.2)',
            'rgba(54, 162, 235, 0.2)',
            'rgba(255, 206, 86, 0.2)',
            'rgba(75, 192, 192, 0.2)',
            'rgba(153, 102, 255, 0.2)',
            'rgba(255, 159, 64, 0.2)'
          ],
          borderColor: [
            'rgba(255,99,132,1)',
            'rgba(54, 162, 235, 1)',
            'rgba(255, 206, 86, 1)',
            'rgba(75, 192, 192, 1)',
            'rgba(153, 102, 255, 1)',
            'rgba(255, 159, 64, 1)'
          ],
          borderWidth: 1
        }]
      },
      options: {
        scales: {
          yAxes: [{
            ticks: {
              beginAtZero: true
            }
          }]
        }
      }
    });
  }
  private renderChart(chartTypeValue: string, projectItems: IProjectItem[]): void {
    let chartRenderType: string;
    switch (chartTypeValue) {
      case '1':
        {
          chartRenderType = 'pie';
          break;
        }
      case '2': {
        chartRenderType = 'doughnut';
        break;
      }
      default: {
        chartRenderType = 'pie';
        break;
      }
    }
    //var ctx = document.getElementById("myChart");
    let ctx: Element = this.domElement.querySelector('#myChart');
    let myChart = new Chart(ctx, {
      type: chartRenderType,
      data: {
        labels: projectItems.map((item) => {
          return item.Title;
        }),
        datasets: [{
          label: 'Team Size',
          data: projectItems.map((item) => {
            return item.TeamSize;
          }),
          backgroundColor: [
            'rgb(115, 115, 151)',
            'rgb(196, 90, 107)',
            'rgb(232, 240, 124)',
            'rgb(110, 194, 138)',
            'rgba(153, 102, 255, 0.8)',
            'rgba(255, 159, 64, 0.8)'
          ]
        }]
      },

    });
  }
}
