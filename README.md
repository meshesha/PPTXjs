PPTXjs
==========
[![MIT License][license-image]][license-url]

[license-image]: http://img.shields.io/badge/license-MIT-blue.svg?style=flat
[license-url]: LICENSE
 
### jQuery plugin for convertation pptx to html using pure javascript.
### Demo: https://meshesha.github.io/pptxjs/

# Version: 
### 1.10.2
# environment
### browsers:
- IE > 10
- Edge
- FireFox
- chrome
### Support:
----
* Text
  * Font size
  * Font family
  * Font style: blod, italic, underline, stok
  * Color
  * hyperlink
  * bullets (include numeric)
* Text block (convert to Div)
  * Align (Horizontal and Vertical)
  * Background color (single color)
  * Border (borderColor, borderWidth, borderType, strokeDasharray)
* Shapes (support most of shapes)
  * Background color (single color, gradient colors)
  * Background image
  * Rotations
  * Align
  * Border
* Custom shape
* Media
  * Picture (jpg/jpeg,png,gif,svg)
  * Video (html5 video player: mp4,ogg,WebM)
    * IE:MP4.
    * Chrome:MP4,	WebM,Ogg.
    * Firefox:MP4,WebM,Ogg.
  * Audio (html5 audio player:mp3,ogg,Wav)
    * IE:mp3.
    * Chrome:mp3,Wav,Ogg.
    * Firefox:mp3,Wav,Ogg  
* Graph
  * Bar chart
  * Line chart
  * Pie chart
  * Scatter chart
* SmartArt diagrams
* Custom table
* Theme table
* Theme
  * Background color
  * Background image
* Equations and formulas
  * display Equations and formulas as image
* and more ...

###  usage:
----
 include necessary css files:
 ```
<link rel="stylesheet" href="./css/pptxjs.css">
<link rel="stylesheet" href="./css/nv.d3.min.css"> <!-- for charts graphs -->
```
 include necessary js files:
 ```
<script type="text/javascript" src="./js/jquery-1.11.3.min.js"></script>
<script type="text/javascript" src="./js/jszip.min.js"></script> <!-- v2.. , NOT v.3.. -->
<script type="text/javascript" src="./js/filereader.js"></script> <!--https://github.com/meshesha/filereader.js -->
<script type="text/javascript" src="./js/d3.min.js"></script> <!-- for charts graphs -->
<script type="text/javascript" src="./js/nv.d3.min.js"></script> <!-- for charts graphs -->
<script type="text/javascript" src="./js/pptxjs.js"></script>
<script type="text/javascript" src="./js/divs2slides.js"></script> <!-- for slide show -->
 ```
 html body :
 ```
 ...
   <div id="your_div_id_result"></div>
   optional:
   <input id="upload_pptx_fiile" type="file" />
 ...
 ```
 add javascript:
 ```
<script type="text/javascript">
 $("#your_div_id_result").pptxToHtml({
   pptxFileUrl: "path/to/yore_pptx_file.pptx", 
   fileInputId: "upload_pptx_fiile",
   slidesScale: "", //Change Slides scale by percent
   slideMode: false,
   keyBoardShortCut: false,
   mediaProcess: true, /** true,false: if true then process video and audio files */
   jsZipV2: "./js/jszip.min.js", /*flase or 'path/to/jsZip.V2.js' */
   slideModeConfig: {  //on slide mode (slideMode: true)
     first: 1,
     nav: false, /** true,false : show or not nav buttons*/
     navTxtColor: "white", /** color */
     showPlayPauseBtn: false,/** true,false */
     keyBoardShortCut: false, /** true,false */
     showSlideNum: false, /** true,false */
     showTotalSlideNum: false, /** true,false */
     autoSlide: false, /** false or seconds (the pause time between slides) , F8 to active(keyBoardShortCut: true) */
     randomAutoSlide: false, /** true,false ,autoSlide:true */ 
     loop: false,  /** true,false */
     background: "black", /** false or color*/
     transition: "default", /** transition type: "slid","fade","default","random" , to show transition efects :transitionTime > 0.5 */
     transitionTime: 1 /** transition time in seconds */           
   }
 });
</script>
 ``` 
# Changelog

* V.1.10.2
  * new divs2slides v.1.3.1
  * fixed some issues
* V.1.10.0
  * added the ability to load jsZip v.2  in case jsZip v.3 is loaded for another use.
  *  (note: using this method will reload the page)
  *  and fixed some errors issue.
* V.1.9.3
  * support Equations and formulas as Image
  * Added an ability to scale Slides in percent
  * and fixed background color issue.
# License
- Copyright © 2017 Meshesha
- MIT
