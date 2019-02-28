import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'PrintPageApplicationCustomizerStrings';

/*Ext Libs */
var jQuery = require('jQuery');
require('jquery-fab');
require('../../../node_modules/jquery-fab/jquery-fab.css'); 
require('p-loading');
require('../../../node_modules/p-loading/dist/css/p-loading.min.css'); 
import * as html2canvas from 'html2canvas';
import * as jsPDF from 'jspdf';
import * as floatingActionButton from 'materialize-css';

const LOG_SOURCE: string = 'PrintPageExtension';

export interface IPrintPageApplicationCustomizerProperties {
}

export default class PrintPageApplicationCustomizer
  extends BaseApplicationCustomizer<IPrintPageApplicationCustomizerProperties> {

  private _fabButton : PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Caption} button`);
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css');

    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    let that = this;

    // Handling the bottom placeholder
    if (!this._fabButton) {
      this._fabButton =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose });

      if (!this._fabButton) {
        console.error('The expected placeholder (Bottom) was not found.');
        return;
      }

      if (this._fabButton.domElement) {
        this._fabButton.domElement.innerHTML = `<div  >
          <div id='wrapper' style='bottom: 30px'></div>
            </div>
       `;

        $(document).ready(function(){
          var links = [
            {
              "bgcolor":"#2980b9",
              "icon":"<i class='fa fa-plus'></i>",
            },
            {
              "bgcolor":"#f1c40f",
              "color":"fffff",
              "icon":"<i class='fa fa-file-pdf-o'></i>",
              "id": "printPDF"
            },
            {
              "bgcolor":"#FF7F50",
              "color":"fffff",
              "icon":"<i class='fa fa-picture-o'></i>",
              "id": "printImage"
            }
            ];
              var options = {
              rotate: true
            };
          ($('#wrapper')as any).jqueryFab(links, options);

          $("#printPDF").click(function() {
            that.printPage(true);
          });

          $("#printImage").click(function() {
            that.printPage(false);
          });
        });
      }
    }
  }

  private printPage(isPdf: boolean): void{
    //For Pages
    var c:any = document.getElementById('spPageCanvasContent');
    if(!c){
      //For documents library
      c = document.getElementsByClassName("Files-rightPaneInteractionContainer")[0];
    }

    var canvasShiftImg = function(img, shiftAmt, scale, pageHeight, pageWidth){
      var c = document.createElement('canvas'),
        ctx = c.getContext('2d'),
        shifter = Number(shiftAmt || 0),
        scaledImgHeight = img.height * scale,
        scaledImgWidth = img.width * scale;
        
      ctx.canvas.height = pageHeight;
      ctx.canvas.width = pageWidth;
      ctx.drawImage(img, 0, shifter, scaledImgWidth, scaledImgHeight)
      
      return c;
    };

    var canvasToImg = function(canvas, loaded, error){
      var dataURL = canvas.toDataURL('image/png'),
        img = new Image();
      if(!isPdf){
        var win = window.open();
        win.document.write('<iframe src="' + dataURL  + '" frameborder="0" style="border:0; top:0px; left:0px; bottom:0px; right:0px; width:100%; height:100%;" allowfullscreen></iframe>');
      }
      else{
        img.onload = loaded;
        img.onerror = error;
        img.src = dataURL;
      }
    };
    
    var imageToPdf = function(){
      var pageOrientation = this.width >= this.height ? 'landscape' : 'portrait';

      var img = this,
        pdf = new jsPDF({
          orientation: pageOrientation,
          unit: 'px',
          format: [img.width, img.height]
        }),
        pdfInternals = pdf.internal,
        pdfPageSize = pdfInternals.pageSize,
        pdfScaleFactor = pdfInternals.scaleFactor,
        pdfPageWidth = pdfPageSize.width ,
        pdfPageHeight = pdfPageSize.height ,
        pdfPageWidthPx = pdfPageWidth * pdfScaleFactor,
        pdfPageHeightPx = pdfPageHeight * pdfScaleFactor,
        
        imgScaleFactor = Math.min(pdfPageWidthPx / img.width, 1),
        imgScaledHeight = img.height * imgScaleFactor,
        
        shiftAmt = 0,
        done = false;

        var newCanvas = canvasShiftImg(img, shiftAmt, imgScaleFactor, pdfPageHeightPx, pdfPageWidthPx);
        pdf.addImage(newCanvas, 'png', 0, 0, pdfPageWidth, 0, null, 'SLOW');
      
      pdf.save('printThisPage.pdf');
    };

    var imageLoadError = function(){
      alert('there was an image load error :(');
    };

    var scrollRegion = $("div[class^='scrollRegion']");
    var hasVerticalScrollbar = scrollRegion.prop("scrollHeight") > scrollRegion.prop("clientHeight");
    if(hasVerticalScrollbar){
      scrollRegion.animate({ scrollTop: scrollRegion.prop("scrollHeight")}, 2000);
    }

    ($(".SPPageChrome") as any).ploading({
      action: 'show'
    });

    setTimeout(
      function() 
      {
        html2canvas(c,
          {
            useCORS: true,
          }
          ).
          then((canvas) => {
              canvasToImg(canvas, imageToPdf, imageLoadError);
        });
        ($(".SPPageChrome") as any).ploading({
          action: 'hide'
        });
      }, 2000);
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom bottom placeholders.');
  }
}
