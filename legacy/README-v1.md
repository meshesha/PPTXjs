PPTXjs
==========
[![MIT License][license-image]][license-url]

[license-image]: http://img.shields.io/badge/license-MIT-blue.svg?style=flat
[license-url]: LICENSE
 
### jQuery plugin for convertation pptx to html using pure javascript.
### Demo: https://pptx.js.org/pages/demos.html

# environment
### browsers:
- IE11
- Edge
- FireFox
- chrome
### Support:
----
* Text
  * Font size
  * Font family
  * Font style: bold, italic, underline, stoke
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
    * YouTube (v1.11.0)
    * vimeo (v1.11.0)
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
* Tables
  * Custom table
  * Theme table
* Theme
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
<script type="text/javascript" src="./js/dingbat.js"></script> <!--for bullets -->
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
   themeProcess: true, /*true (default) , false, "colorsAndImageOnly"*/
   incSlide:{height: 2,width:2 }, /*increase height or/and width by 2 px*/
   slideType: "divs2slidesjs", /*'divs2slidesjs' (default) , 'revealjs'(https://revealjs.com)
   slideModeConfig: {  //divs2slidesjs - on slide mode (slideMode: true)
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
   },
			revealjsConfig: { /* for 'revealjs' settings (https://revealjs.com) */
				transition: 'zoom',
				// backgroundTransition: 'zoom', 
				// autoSlide: 5000,
				// loop: true
				slideNumber: true
			}
 });
</script>
 ``` 
# Changelog
* v1.21.1
  * fixed issues:
    - [#16](https://github.com/meshesha/PPTXjs/issues/16)

* v1.21.00
  * add theme (background) support
  * improved tables
  * improved bullets (add ./js/dingbat.js)
  * fixed issues:
    - [#5](https://github.com/meshesha/PPTXjs/issues/5)
    - [#7](https://github.com/meshesha/PPTXjs/issues/7)
    - [#8](https://github.com/meshesha/PPTXjs/issues/8)
    - [#9](https://github.com/meshesha/PPTXjs/issues/9)
    - [#10](https://github.com/meshesha/PPTXjs/issues/10)
    - [#11](https://github.com/meshesha/PPTXjs/issues/11)
    - [#13](https://github.com/meshesha/PPTXjs/issues/13)
    - [#15](https://github.com/meshesha/PPTXjs/issues/15)
  * more documentation coming soon ...

* v1.11.0
  * Support for embedding video from a link (tested youtube and vimeo links)
  * support 'revealjs'(https://revealjs.com) (It is not recommended to add a theme because it distorts some of the elements like tables )
  * I think i fix issue [officetohtml/issues/7](https://github.com/meshesha/officetohtml/issues/7) (not tested) 
  * Change loading view 
  * Fix center slides in fullscreen mode - (https://github.com/meshesha/divs2slides v1.3.3)
  * Support emf and wmf files - microsoft files, supported only in Internet Explorer (test in IE11)

* V.1.10.4
  * fixed security issue : [#3](https://github.com/meshesha/PPTXjs/issues/3)
  
* V.1.10.3
  * new divs2slides (v.1.3.2)
  * fixed div width issue
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
