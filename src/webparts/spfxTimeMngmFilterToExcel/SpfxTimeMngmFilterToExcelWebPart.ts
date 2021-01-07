import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxTimeMngmFilterToExcelWebPart.module.scss';
import * as strings from 'SpfxTimeMngmFilterToExcelWebPartStrings';


import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';


export interface ISpfxTimeMngmFilterToExcelWebPartProps {
  description: string;
}

export default class SpfxTimeMngmFilterToExcelWebPart extends BaseClientSideWebPart<ISpfxTimeMngmFilterToExcelWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spfxTimeMngmFilterToExcel }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">

              <div class="${ styles.flex } ${ styles.headLine }">
                <div>
                  <label>חודש</label>
                  <select id="month">
                    <option>1</option>
                    <option>2</option>
                    <option>3</option>
                    <option>4</option>
                    <option>5</option>
                    <option>6</option>
                    <option>7</option>
                    <option>8</option>
                    <option>9</option>
                    <option>10</option>
                    <option>11</option>
                    <option>12</option>
                  </select>
                </div>

                <div>
                  <label>שנה</label>
                  <select id="year">
                    <option>2020</option>
                    <option>2021</option>
                    <option>2022</option>
                    <option>2023</option>
                  </select>
                </div>
                           
                <div>
                  <label>נושא</label>
                  <select id="subject">
                  </select>
                </div>
                
              </div>


              <div id="content" class="${ styles.content } ${ styles.flexCol }">טוען ...</div>
          </div>
        </div>
      </div>`;

    let yyyy = new Date().getFullYear()
    let mm = new Date().getMonth()+1
    if (mm == 1) {
      mm = 12
      yyyy--
    }
    this.domElement.querySelector('#year')['value'] = yyyy.toString()
    this.domElement.querySelector('#month')['value'] = mm.toString()


    //TimeMngAppWorkSubject
    this.getListItems('TimeMngApp-WorkSubject', arr => {
      let h = `<option>בחר</option>`
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i];
        h += `<option>${item.Title}</option>`
      }
      this.domElement.querySelector('#subject').innerHTML = h

      this._setEventsHandlers()
    }, '$select=Title');
  }

  public setSearch(){
    let year = parseInt(this.domElement.querySelector('#year')['value'])
    let month = parseInt(this.domElement.querySelector('#month')['value'])
    let sub = this.domElement.querySelector('#subject')['value']

    let strMM = month.toString()
    if (strMM.length == 1) {
      strMM = `0${strMM}`
    } 

    let start = `datetime'${year}-${strMM}-01T00:00:00Z'`

    let strMMend = (month+1).toString()
    if (strMMend.length == 1) {
      strMMend = `0${strMMend}`
    } 
    let end = `datetime'${year}-${strMMend}-01T00:00:00Z'`
    if (month == 12) {
      end = `datetime'${year+1}-01-01T00:00:00Z'`
    }

    //$filter=Start_x0020_Date le datetime'2016-03-26T09:59:32Z'
    //https://kiserachamin.sharepoint.com/sites/apps-data-center/_api/web/lists/GetByTitle('TimeMngApp-Hours')/Items?$top=1000&
    //$filter=Title eq 'דניאל חן' 
    // and Created ge datetime'2020-12-01T00:00:00Z'
    // and Created le datetime'2021-01-01T00:00:00Z'
    let q = `$filter=Title eq '${sub}' and Created ge ${start} and Created le ${end}`


    this.getListItems('TimeMngApp-Hours', items=>{
      console.log('after search', items);
      let header = `<div class="${ styles.item } ${ styles.header }">
                <div class="${ styles.start }"><label>תאריך התחלה</label></div>
                <div class="${ styles.end }"><label>תאריך סיום</label></div>
                <div class="${ styles.total }"><label>זמן עבודה</label></div>
              </div>`
      let h = header
      let totalTime = 0.0;
      for (let i = 0; i < items.length; i++) {
        const x = items[i];
        x._Created = this.dateFormat(new Date(x.Created))
        x._EndTime = this.dateFormat(new Date(x.EndTime))
        h += `<div class="${ styles.item }">
                <div class="${ styles.start }"><span>${x._Created}</span></div>
                <div class="${ styles.end }"><span>${x._EndTime}</span></div>
                <div class="${ styles.total }"><span>${x.TotalTime}</span></div>
              </div>`
        totalTime += parseFloat(x.TotalMinutes)
        if (i == 10) {
          h += header
        }
      }

      console.log('totalTime', totalTime);
      let hrs = (totalTime/60).toFixed(2)
      h = `<div class="${ styles.item } ${ styles.sumTotal }"><label>סכה"כ שעות</label> &nbsp;&nbsp; <span>${hrs}</span></div>` + h
      this.domElement.querySelector('#content').innerHTML = h

    }, q)
  }

  private dateFormat(d){
    let MM = d.getMinutes()
    if (MM < 10) {
      MM = `0${MM}`
    }
    return `${d.getDate()}/${d.getMonth()+1}/${d.getYear()-100}, ${d.getHours()}:${MM}`
    
  }

  private _setEventsHandlers(): void { 
    console.log('_setEventsHandlers');
    const webPart: SpfxTimeMngmFilterToExcelWebPart = this; 

    this.domElement.querySelectorAll('select')
      //.forEach(elem => elem.addEventListener('click', this.setSearch.bind(this)));
      .forEach(elem => elem.addEventListener('change', this.setSearch.bind(this)));
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
                })
              ]
            }
          ]
        }
      ]
    };
  }




  public getListItems(listname:string, callback:Function, query?:string): void {
    console.log('getListItems asking list items for', listname);
    query = query ? query : '';

    this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl +
      `/_api/web/lists/GetByTitle('${listname}')/Items?$top=1000&${query}`, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
              response.json().then((data)=> {
                  console.log('list items for', listname, query, data);
                  callback(data.value)
              });
          });
    }

}



