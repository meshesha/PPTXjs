/**
 * pptxjs.js
 * Ver. : 1.0.4
 * Author: meshesha , https://github.com/meshesha
 * LICENSE: MIT
 * url:https://github.com/meshesha/PPTXjs
 */

(function ( $ ) {
    if (FileReader.prototype.readAsBinaryString === undefined) {
        FileReader.prototype.readAsBinaryString = function (fileData) {
            var binary = "";
            var pt = this;
            var reader = new FileReader();
            reader.onload = function (e) {
                var bytes = new Uint8Array(reader.result);
                var length = bytes.byteLength;
                for (var i = 0; i < length; i++) {
                    binary += String.fromCharCode(bytes[i]);
                }
                //pt.result  - readonly so assign content to another property
                pt.content = binary;
                pt.onload(); // thanks to @Denis comment
            }
            reader.readAsArrayBuffer(fileData);
        }
    }    
    $.fn.pptxToHtml = function( options ) {
 		//var worker;
        var $result = $(this);
        var divId = $result.attr("id");
        
        var isDone = false;

        var MsgQueue = new Array();

        var themeContent = null;

        var slideLayoutClrOvride = "";

        var chartID = 0;
        var _order = 1;
        var titleFontSize = 42;
        var bodyFontSize = 20;
        var otherFontSize = 16;
        var isSlideMode = false;
        var styleTable = {};        
        // This is the easiest way to have default options.
        var settings = $.extend({
            // These are the defaults.
            pptxFileUrl: "",
            fileInputId: "",
            slideMode: false, /** true,false*/
            keyBoardShortCut: false,  /** true,false ,condition: slideMode: true*/
            slideModeConfig: {
                first: 1,
                nav: true, /** true,false : show or not nav buttons*/
                navTxtColor: "black", /** color */
                navNextTxt:"&#8250;",
                navPrevTxt: "&#8249;",
                keyBoardShortCut: true, /** true,false ,condition: */
                showSlideNum: true, /** true,false */
                showTotalSlideNum: true, /** true,false */
                autoSlide:true, /** false or seconds , F8 to active ,keyBoardShortCut: true */
                randomAutoSlide: false, /** true,false ,autoSlide:true */ 
                loop: false,  /** true,false */
                background: false, /** false or color*/
                transition: "default", /** transition type: "slid","fade","default","random" , to show transition efects :transitionTime > 0.5 */
                transitionTime: 1 /** transition time between slides in seconds */               
            }
        }, options );
        //
        $("#"+divId).prepend(
            $("<span></span>").attr({
                "class":"slides-loadnig-msg",
                "style":"display:block; color:blue; font-size:20px; width:50%; margin:0 auto;"
            }).html("Loading...")
        );
        if(settings.slideMode){
            //check if divs2slides.js was included, include if not
            if(!jQuery().divs2slides) {
                // the plugin is not loaded => load it:
                jQuery.getScript('./js/divs2slides.js', function() {
                    // the plugin is now loaded => use divs2slides.min.js
                });
            }
        }
        if(settings.keyBoardShortCut && settings.slideMode){
            $(document).bind("keydown",function(event){
                event.preventDefault();
                var key = event.keyCode;
                if(key==116 && !isSlideMode){ //F5
                    console.log(key)
                    isSlideMode = true;
                    $("#"+divId+" .slide").hide();
                    //setTimeout(function() {
                    if(isDone){
                        var slideConf = settings.slideModeConfig;
                        $(".slides-loadnig-msg").remove()
                        $("#"+divId).divs2slides({
                            first: slideConf.first,
                            nav: slideConf.nav,
                            showPlayPauseBtn: settings.showPlayPauseBtn,
                            navTxtColor: slideConf.navTxtColor,
                            navNextTxt: slideConf.navNextTxt,
                            navPrevTxt: slideConf.navPrevTxt,
                            keyBoardShortCut: slideConf.keyBoardShortCut,
                            showSlideNum: slideConf.showSlideNum,
                            showTotalSlideNum: slideConf.showTotalSlideNum,
                            autoSlide: slideConf.autoSlide,
                            randomAutoSlide: slideConf.randomAutoSlide,
                            loop: slideConf.loop,
                            background : slideConf.background,
                            transition: slideConf.transition, 
                            transitionTime: slideConf.transitionTime 
                        });
                    }
                    //}, 1500);
                }
            });
        }
 		var loadFile=function(url,callback){
			JSZipUtils.getBinaryContent(url,callback);
        }
        //if(settings.fileInputId ==""){
            loadFile(settings.pptxFileUrl,function(err,content){
                var blob  = new Blob([content]);
                var reader = new FileReader();
                reader.onload = function(aEvent) {
                    if (!aEvent) { //for soport readAsBinaryString in IE11
                        convertToHtml(btoa(reader.content));
                    }else{
                        convertToHtml(btoa(aEvent.target.result));
                    }
                };
                reader.readAsBinaryString(blob); 	
            });
        //}else{
        if(settings.fileInputId !=""){
            $("#"+settings.fileInputId).on("change", function(evt) {
                $result.html("");
                var file = evt.target.files[0];
               // var fileName = file.name;
                var fileType = file.type;
                if(fileType=="application/vnd.openxmlformats-officedocument.presentationml.presentation"){
                    var reader = new FileReader();
                    reader.onload = (function(theFile) {
                        return function(e) {
                            if (!e) { //for soport readAsBinaryString in IE11
                                convertToHtml(btoa(reader.content));
                            }else{
                                convertToHtml(btoa(e.target.result));
                            }
                        }
                    })(file);
                    reader.readAsBinaryString(file);
                }else{
                    alert("This is not pptx file");
                }
            });
        }
        function convertToHtml(file) {
             //'use strict';
            var zip = new JSZip(), s;
            if (typeof file === 'string') { // Load
                zip = zip.load(file, { base64: true });  //zip.load(file, { base64: true });
                var rslt_ary = processPPTX(zip);
                //s = readXmlFile(zip, 'ppt/tableStyles.xml');
                for(var i=0;i<rslt_ary.length;i++){
                    switch(rslt_ary[i]["type"]){
                        case "slide":
                            $result.append(rslt_ary[i]["data"]);
                            break;
                        case "pptx-thumb":
                            //$("#pptx-thumb").attr("src", "data:image/jpeg;base64," +rslt_ary[i]["data"]);
                            break;
                        case "slideSize":
                                //var slideWidth = rslt_ary[i]["data"].width;
                                //var slideHeight = rslt_ary[i]["data"].height;
                            break;
                        case "globalCSS":
                            $result.append("<style>" +rslt_ary[i]["data"] + "</style>");
                            break;
                        case "ExecutionTime":
                            //isDone = true;
                            // $result.prepend("<div id='presentation_toolbar'></div>");
                            processMsgQueue(MsgQueue);
                            setNumericBullets($(".block"));
                            setNumericBullets($("table td"));
                            if(settings.slideMode && !isSlideMode){
                                isSlideMode = true;
                                $("#"+divId+" .slide").hide();
                                setTimeout(function() {
                                    var slideConf = settings.slideModeConfig;
                                    $(".slides-loadnig-msg").remove();
                                    $("#"+divId).divs2slides({
                                        first: slideConf.first,
                                        nav: slideConf.nav,
                                        showPlayPauseBtn: settings.showPlayPauseBtn,
                                        navTxtColor: slideConf.navTxtColor,
                                        navNextTxt: slideConf.navNextTxt,
                                        navPrevTxt: slideConf.navPrevTxt,
                                        keyBoardShortCut: slideConf.keyBoardShortCut,
                                        showSlideNum: slideConf.showSlideNum,
                                        showTotalSlideNum: slideConf.showTotalSlideNum,
                                        autoSlide: slideConf.autoSlide,
                                        randomAutoSlide: slideConf.randomAutoSlide,
                                        loop: slideConf.loop,
                                        background : slideConf.background,
                                        transition: slideConf.transition, 
                                        transitionTime: slideConf.transitionTime 
                                    });   
                                }, 1500);
                            }else if(!settings.slideMode){
                                $(".slides-loadnig-msg").remove();
                            }
                            break;
                        default:                        
                    }
                }
            }
        }
        function processPPTX(zip) {
            var post_ary = [];
            var dateBefore = new Date();
            
            //var zip = new JSZip(data);
            
            if (zip.file("docProps/thumbnail.jpeg") !== null) {
                var pptxThumbImg = base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer());
               post_ary.push({
                    "type": "pptx-thumb",
                    "data": pptxThumbImg
                });
            }
            
            var filesInfo = getContentTypes(zip);
            var slideSize = getSlideSize(zip);
            themeContent = loadTheme(zip);

            tableStyles = readXmlFile(zip, "ppt/tableStyles.xml");

            post_ary.push({
                "type": "slideSize",
                "data": slideSize
            });
            
            var numOfSlides = filesInfo["slides"].length;
            for (var i=0; i<numOfSlides; i++) {
                var filename = filesInfo["slides"][i];
                var slideHtml = processSingleSlide(zip, filename, i, slideSize);
                post_ary.push({
                    "type": "slide",
                    "data": slideHtml
                });
               post_ary.push({
                    "type": "progress-update",
                    "data": (i + 1) * 100 / numOfSlides
                });
            }

           post_ary.push({
                "type": "globalCSS",
                "data": genGlobalCSS()
            });
            
            var dateAfter = new Date();
            post_ary.push({
                "type": "ExecutionTime",
                "data": dateAfter - dateBefore
            });
            return post_ary;
        }

        function readXmlFile(zip, filename) {
            return tXml(zip.file(filename).asText());
        }
        function getContentTypes(zip) {
            var ContentTypesJson = readXmlFile(zip, "[Content_Types].xml");
            var subObj = ContentTypesJson["Types"]["Override"];
            var slidesLocArray = [];
            var slideLayoutsLocArray = [];
            for (var i=0; i<subObj.length; i++) {
                switch (subObj[i]["attrs"]["ContentType"]) {
                    case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                        slidesLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                        break;
                    case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                        slideLayoutsLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                        break;
                    default:
                }
            }
            return {
                "slides": slidesLocArray,
                "slideLayouts": slideLayoutsLocArray
            };
        }

        function getSlideSize(zip) {
            // Pixel = EMUs * Resolution / 914400;  (Resolution = 96)
            var content = readXmlFile(zip, "ppt/presentation.xml");
            var sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"]
            return {
                "width": parseInt(sldSzAttrs["cx"]) * 96 / 914400,
                "height": parseInt(sldSzAttrs["cy"]) * 96 / 914400
            };
        }

        function loadTheme(zip) {
            var preResContent = readXmlFile(zip, "ppt/_rels/presentation.xml.rels");
            var relationshipArray = preResContent["Relationships"]["Relationship"];
            var themeURI = undefined;
            if (relationshipArray.constructor === Array) {
                for (var i=0; i<relationshipArray.length; i++) {
                    if (relationshipArray[i]["attrs"]["Type"] === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                        themeURI = relationshipArray[i]["attrs"]["Target"];
                        break;
                    }
                }
            } else if (relationshipArray["attrs"]["Type"] === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                themeURI = relationshipArray["attrs"]["Target"];
            }
            
            if (themeURI === undefined) {
                throw Error("Can't open theme file.");
            }
            
            return readXmlFile(zip, "ppt/" + themeURI);
        }

        function processSingleSlide(zip, sldFileName, index, slideSize) {
            /*
            self.postMessage({
                "type": "INFO",
                "data": "Processing slide" + (index + 1)
            });
            */
            // =====< Step 1 >=====
            // Read relationship filename of the slide (Get slideLayoutXX.xml)
            // @sldFileName: ppt/slides/slide1.xml
            // @resName: ppt/slides/_rels/slide1.xml.rels
            var resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
            var resContent = readXmlFile(zip, resName);
            var RelationshipArray = resContent["Relationships"]["Relationship"];
            var layoutFilename = "";
            var slideResObj = {};
            if (RelationshipArray.constructor === Array) {
                for (var i=0; i<RelationshipArray.length; i++) {
                    switch (RelationshipArray[i]["attrs"]["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                            layoutFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                            break;
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
                        default:
                            slideResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                                "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                            };
                    }
                }
            } else {
                layoutFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
            }
            //console.log(slideResObj);
            // Open slideLayoutXX.xml
            var slideLayoutContent = readXmlFile(zip, layoutFilename);
            var slideLayoutTables = indexNodes(slideLayoutContent);
            var sldLayoutClrOvr = slideLayoutContent["p:sldLayout"]["p:clrMapOvr"]["a:overrideClrMapping"];

            //console.log(slideLayoutClrOvride);
            if(sldLayoutClrOvr !== undefined){
                slideLayoutClrOvride = sldLayoutClrOvr["attrs"];
            }
            // =====< Step 2 >=====
            // Read slide master filename of the slidelayout (Get slideMasterXX.xml)
            // @resName: ppt/slideLayouts/slideLayout1.xml
            // @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
            var slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
            var slideLayoutResContent = readXmlFile(zip, slideLayoutResFilename);
            RelationshipArray = slideLayoutResContent["Relationships"]["Relationship"];
            var masterFilename = "";
            var layoutResObj = {};
            if (RelationshipArray.constructor === Array) {
                for (var i=0; i<RelationshipArray.length; i++) {
                    switch (RelationshipArray[i]["attrs"]["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
                            masterFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                            break;
                        default:
                            layoutResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                                "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                            };
                    }
                }
            } else {
                masterFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
            }
            // Open slideMasterXX.xml
            var slideMasterContent = readXmlFile(zip, masterFilename);
            var slideMasterTextStyles = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
            var slideMasterTables = indexNodes(slideMasterContent);
            
            /////////////////Amir/////////////
            //Open slideMasterXX.xml.rels
            var slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
            var slideMasterResContent = readXmlFile(zip, slideMasterResFilename);
            RelationshipArray = slideMasterResContent["Relationships"]["Relationship"];
            var themeFilename = "";
            var masterResObj = {};
            if (RelationshipArray.constructor === Array) {
                for (var i=0; i<RelationshipArray.length; i++) {
                    switch (RelationshipArray[i]["attrs"]["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                            themeFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                            break;
                        default:
                            masterResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                                "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                            };
                    }
                }
            } else {
                themeFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
            }
            //console.log(themeFilename)
            //Load Theme file
            if(themeFilename !== undefined){
                themeContent =  readXmlFile(zip, themeFilename);
            }
            // =====< Step 3 >=====
            var slideContent = readXmlFile(zip, sldFileName);
            var nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];
            var warpObj = {
                "zip": zip,
                "slideLayoutTables": slideLayoutTables,
                "slideMasterTables": slideMasterTables,
                "slideResObj": slideResObj,
                "slideMasterTextStyles": slideMasterTextStyles,
                "layoutResObj":layoutResObj,
                "masterResObj":masterResObj
            };
            
            var bgColor = getSlideBackgroundFill(slideContent, slideLayoutContent, slideMasterContent,warpObj);
            
            var result = "<div class='slide' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
            //result += "<div>"+getBackgroundShapes(slideContent, slideLayoutContent, slideMasterContent,warpObj) + "</div>" - TODO
            for (var nodeKey in nodes) {
                if (nodes[nodeKey].constructor === Array) {
                    for (var i=0; i<nodes[nodeKey].length; i++) {
                        result += processNodesInSlide(nodeKey, nodes[nodeKey][i], warpObj);
                    }
                } else {
                    result += processNodesInSlide(nodeKey, nodes[nodeKey], warpObj);
                }
            }
            
            return result + "</div>";
        }

        function indexNodes(content) {
            
            var keys = Object.keys(content);
            var spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];
            
            var idTable = {};
            var idxTable = {};
            var typeTable = {};
            
            for (var key in spTreeNode) {

                if (key == "p:nvGrpSpPr" || key == "p:grpSpPr") {
                    continue;
                }
                
                var targetNode = spTreeNode[key];
                
                if (targetNode.constructor === Array) {
                    for (var i=0; i<targetNode.length; i++) {
                        var nvSpPrNode = targetNode[i]["p:nvSpPr"];
                        var id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                        var idx = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                        var type = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);
                        
                        if (id !== undefined) {
                            idTable[id] = targetNode[i];
                        }
                        if (idx !== undefined) {
                            idxTable[idx] = targetNode[i];
                        }
                        if (type !== undefined) {
                            typeTable[type] = targetNode[i];
                        }
                    }
                } else {
                    var nvSpPrNode = targetNode["p:nvSpPr"];
                    var id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                    var idx = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                    var type = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);
                    
                    if (id !== undefined) {
                        idTable[id] = targetNode;
                    }
                    if (idx !== undefined) {
                        idxTable[idx] = targetNode;
                    }
                    if (type !== undefined) {
                        typeTable[type] = targetNode;
                    }
                }
                
            }
            
            return {"idTable": idTable, "idxTable": idxTable, "typeTable": typeTable};
        }

        function processNodesInSlide(nodeKey, nodeValue, warpObj) {
            
            var result = "";
            
            switch (nodeKey) {
                case "p:sp":    // Shape, Text
                    result = processSpNode(nodeValue, warpObj);
                    break;
                case "p:cxnSp":    // Shape, Text (with connection)
                    result = processCxnSpNode(nodeValue, warpObj);
                    break;
                case "p:pic":    // Picture
                    result = processPicNode(nodeValue, warpObj);
                    break;
                case "p:graphicFrame":    // Chart, Diagram, Table
                    result = processGraphicFrameNode(nodeValue, warpObj);
                    break;
                case "p:grpSp":    
                    result = processGroupSpNode(nodeValue, warpObj);
                    break;
                default:
            }
            
            return result;
            
        }

        function processGroupSpNode(node, warpObj) {
            
            var factor = 96 / 914400;
            
            var xfrmNode = node["p:grpSpPr"]["a:xfrm"];
            var x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * factor;
            var y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * factor;
            var chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * factor;
            var chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * factor;
            var cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * factor;
            var cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * factor;
            var chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * factor;
            var chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * factor;
            
            var order = node["attrs"]["order"];
            
            var result = "<div class='block group' style='z-index: " + order + "; top: " + (y - chy) + "px; left: " + (x - chx) + "px; width: " + (cx - chcx) + "px; height: " + (cy - chcy) + "px;'>";
            
            // Procsee all child nodes
            for (var nodeKey in node) {
                if (node[nodeKey].constructor === Array) {
                    for (var i=0; i<node[nodeKey].length; i++) {
                        result += processNodesInSlide(nodeKey, node[nodeKey][i], warpObj);
                    }
                } else {
                    result += processNodesInSlide(nodeKey, node[nodeKey], warpObj);
                }
            }
            
            result += "</div>";
            
            return result;
        }

        function processSpNode(node, warpObj) {
            
            /*
            *  958    <xsd:complexType name="CT_GvmlShape">
            *  959   <xsd:sequence>
            *  960     <xsd:element name="nvSpPr" type="CT_GvmlShapeNonVisual"     minOccurs="1" maxOccurs="1"/>
            *  961     <xsd:element name="spPr"   type="CT_ShapeProperties"        minOccurs="1" maxOccurs="1"/>
            *  962     <xsd:element name="txSp"   type="CT_GvmlTextShape"          minOccurs="0" maxOccurs="1"/>
            *  963     <xsd:element name="style"  type="CT_ShapeStyle"             minOccurs="0" maxOccurs="1"/>
            *  964     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
            *  965   </xsd:sequence>
            *  966 </xsd:complexType>
            */
            
            var id = getTextByPathList(node, ["p:nvSpPr","p:cNvPr","attrs","id"]);
            var name = getTextByPathList(node, ["p:nvSpPr","p:cNvPr","attrs","name"]);
            var idx = (getTextByPathList(node, ["p:nvSpPr","p:nvPr","p:ph"]) === undefined) ? undefined : getTextByPathList(node, ["p:nvSpPr","p:nvPr","p:ph","attrs","idx"]);
            var type = (getTextByPathList(node, ["p:nvSpPr","p:nvPr","p:ph"]) === undefined) ? undefined : getTextByPathList(node, ["p:nvSpPr","p:nvPr","p:ph","attrs","type"]);
            var order = getTextByPathList(node, ["attrs","order"]);
            
            var slideLayoutSpNode = undefined;
            var slideMasterSpNode = undefined;
            
            if (type !== undefined) {
                if (idx !== undefined) {
                    slideLayoutSpNode = warpObj["slideLayoutTables"]["typeTable"][type];
                    slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
                } else {
                    slideLayoutSpNode = warpObj["slideLayoutTables"]["typeTable"][type];
                    slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
                }
            } else {
                if (idx !== undefined) {
                    slideLayoutSpNode = warpObj["slideLayoutTables"]["idxTable"][idx];
                    slideMasterSpNode = warpObj["slideMasterTables"]["idxTable"][idx];
                } else {
                    // Nothing
                }
            }
            
            if (type === undefined) {
                type = getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (type === undefined) {
                    type = getTextByPathList(slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                }
            }
            
            return genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj);
        }

        function processCxnSpNode(node, warpObj) {

            var id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
            var name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
            //var idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
            //var type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
            //<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
            var order = node["attrs"]["order"];
            
            return genShape(node, undefined, undefined, id, name, undefined, undefined, order, warpObj);
        }

        function genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj) {
            
            var xfrmList = ["p:spPr", "a:xfrm"];
            var slideXfrmNode = getTextByPathList(node, xfrmList);
            var slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList);
            var slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList);
            
            var result = "";
            var shpId = getTextByPathList(node, ["attrs","order"]);
            //console.log("shpId: ",shpId)
            var shapType = getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);

            //custGeom - Amir
            var custShapType = getTextByPathList(node, ["p:spPr", "a:custGeom"]);
            
            var isFlipV = false;
            if ( getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1" || getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1") {
                isFlipV = true;
            }
            /////////////////////////Amir////////////////////////
            //rotate
            var rotate = angleToDegrees(getTextByPathList(slideXfrmNode, ["attrs", "rot"]));
            //console.log("rotate: "+rotate);
            var txtRotate;
            var txtXframeNode = getTextByPathList(node, ["p:txXfrm"]);
            if (txtXframeNode !== undefined){
                var txtXframeRot =  getTextByPathList(txtXframeNode,["attrs","rot"]);
                if (txtXframeRot !== undefined){
                    txtRotate = angleToDegrees(txtXframeRot)+90;
                }else{
                    txtRotate = rotate;
                }
            }
            //////////////////////////////////////////////////
            if (shapType !== undefined || custShapType !== undefined) {
                var off = getTextByPathList(slideXfrmNode, ["a:off", "attrs"]);
                var x = parseInt(off["x"]) * 96 / 914400;
                var y = parseInt(off["y"]) * 96 / 914400;
                
                var ext = getTextByPathList(slideXfrmNode, ["a:ext", "attrs"]);
                var w = parseInt(ext["cx"]) * 96 / 914400;
                var h = parseInt(ext["cy"]) * 96 / 914400;
                
                result += "<svg class='drawing' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                        "' style='" + 
                            getPosition(slideXfrmNode, undefined, undefined) + 
                            getSize(slideXfrmNode, undefined, undefined) +
                            " z-index: " + order + ";" +
                            "transform: rotate(" +rotate+ "deg);"+
                            "'>";
                result += '<defs>'
                // Fill Color
                var fillColor = getShapeFill(node, true,warpObj);
                var grndFillFlg = false;
                var imgFillFlg = false;
                var clrFillType = getFillType(getTextByPathList(node, ["p:spPr"]));
                /////////////////////////////////////////                    
                if(clrFillType == "GRADIENT_FILL"){
                    grndFillFlg = true;
                    var color_arry = fillColor.color;
                    var angl = fillColor.rot;
                    var svgGrdnt = getSvgGradient(w,h,angl,color_arry,shpId);
                    //fill="url(#linGrd)"
                    result +=  svgGrdnt ;
                }else if(clrFillType == "PIC_FILL"){
                    imgFillFlg = true;
                    var svgBgImg = getSvgImagePattern(fillColor,shpId);
                    //fill="url(#imgPtrn)"
                    //console.log(svgBgImg)
                    result +=  svgBgImg ;
                }        
                // Border Color
                var border = getBorder(node, true);
                
                var headEndNodeAttrs = getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
                var tailEndNodeAttrs = getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
                // type: none, triangle, stealth, diamond, oval, arrow
                
                if ( (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) || 
                    (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) ) {
                    var triangleMarker = "<marker id='markerTriangle_"+shpId+"' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border.color + "' fill='" + border.color + 
                                    "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
                    result += triangleMarker;
                }
                result += '</defs>'
            }
            if (shapType !== undefined && custShapType === undefined) {
                
                switch (shapType) {
                    case "accentBorderCallout1":
                    case "accentBorderCallout2":
                    case "accentBorderCallout3":
                    case "accentCallout1":
                    case "accentCallout2":
                    case "accentCallout3":
                    case "actionButtonBackPrevious":
                    case "actionButtonBeginning":
                    case "actionButtonBlank":
                    case "actionButtonDocument":
                    case "actionButtonEnd":
                    case "actionButtonForwardNext":
                    case "actionButtonHelp":
                    case "actionButtonHome":
                    case "actionButtonInformation":
                    case "actionButtonMovie":
                    case "actionButtonReturn":
                    case "actionButtonSound":
                    case "arc":
                    case "bevel":
                    case "blockArc":
                    case "borderCallout1":
                    case "borderCallout2":
                    case "borderCallout3":
                    case "bracePair":
                    case "bracketPair":
                    case "callout1":
                    case "callout2":
                    case "callout3":
                    case "can":
                    case "chartPlus":
                    case "chartStar":
                    case "chartX":
                    case "chevron":
                    case "chord":
                    case "cloud":
                    case "cloudCallout":
                    case "corner":
                    case "cornerTabs":
                    case "cube":
                    case "diagStripe":
                    case "donut":
                    case "doubleWave":
                    case "downArrowCallout":
                    case "ellipseRibbon":
                    case "ellipseRibbon2":
                    case "flowChartAlternateProcess":
                    case "flowChartCollate":
                    case "flowChartConnector":
                    case "flowChartDecision":
                    case "flowChartDelay":
                    case "flowChartDisplay":
                    case "flowChartDocument":
                    case "flowChartExtract":
                    case "flowChartInputOutput":
                    case "flowChartInternalStorage":
                    case "flowChartMagneticDisk":
                    case "flowChartMagneticDrum":
                    case "flowChartMagneticTape":
                    case "flowChartManualInput":
                    case "flowChartManualOperation":
                    case "flowChartMerge":
                    case "flowChartMultidocument":
                    case "flowChartOfflineStorage":
                    case "flowChartOffpageConnector":
                    case "flowChartOnlineStorage":
                    case "flowChartOr":
                    case "flowChartPredefinedProcess":
                    case "flowChartPreparation":
                    case "flowChartProcess":
                    case "flowChartPunchedCard":
                    case "flowChartPunchedTape":
                    case "flowChartSort":
                    case "flowChartSummingJunction":
                    case "flowChartTerminator":
                    case "folderCorner":
                    case "frame":
                    case "funnel":
                    case "halfFrame":
                    case "heart":
                    case "homePlate":
                    case "horizontalScroll":
                    case "irregularSeal1":
                    case "irregularSeal2":
                    case "leftArrowCallout":
                    case "leftBrace":
                    case "leftBracket":
                    case "leftRightArrowCallout":
                    case "leftRightRibbon":
                    case "irregularSeal1":
                    case "lightningBolt":
                    case "lineInv":
                    case "mathDivide":
                    case "mathEqual":
                    case "mathMinus":
                    case "mathMultiply":
                    case "mathNotEqual":
                    case "mathPlus":
                    case "moon":
                    case "nonIsoscelesTrapezoid":
                    case "noSmoking":
                    case "plaque":
                    case "plaqueTabs":
                    case "quadArrowCallout":
                    case "rect":
                    case "ribbon":
                    case "ribbon2":
                    case "rightArrowCallout":
                    case "rightBrace":
                    case "rightBracket":
                    case "round1Rect":
                    case "round2DiagRect":
                    case "round2SameRect":
                    case "smileyFace":
                    case "snip1Rect":
                    case "snip2DiagRect":
                    case "snip2SameRect":
                    case "snipRoundRect":
                    case "squareTabs":
                    case "sun":
                    case "teardrop":
                    case "upArrowCallout":
                    case "upDownArrowCallout":
                    case "verticalScroll":
                    case "wave":
                    case "wedgeEllipseCallout":
                    case "wedgeRectCallout":
                    case "wedgeRoundRectCallout":
                    case "rect":
                        result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "ellipse":
                        result += "<ellipse cx='" + (w / 2) + "' cy='" + (h / 2) + "' rx='" + (w / 2) + "' ry='" + (h / 2) + "' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "roundRect":
                        result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' rx='7' ry='7' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bentConnector2": 
                        var d = "";
                        if (isFlipV) {
                            d = "M 0 " + w + " L " + h + " " + w + " L " + h + " 0";
                        } else {
                            d = "M " + w + " 0 L " + w + " " + h + " L 0 " + h;
                        }
                        result += "<path d='" + d + "' stroke='" + border.color + 
                                        "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' fill='none' ";
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_"+shpId+")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_"+shpId+")' ";
                        }
                        result += "/>";
                        break;
                    case "rtTriangle":
                        result += " <polygon points='0 0,0 " + h + ","+w+" "+h+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "triangle":
                        var shapAdjst = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd", "attrs", "fmla"]);
                        var shapAdjst_val = 0.5;
                        if(shapAdjst !== undefined){
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) * 96 / 9144000;
                            //console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nshapAdjst_val: "+shapAdjst_val);
                        }
                        result += " <polygon points='"+(w*shapAdjst_val)+" 0,0 " + h + ","+w+" "+h+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";            
                        break;
                    case "diamond":
                        result += " <polygon points='" + (w/2) + " 0,0 " + (h/2) + "," + (w/2)+" "+h+"," + w + " " + (h/2) +"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "trapezoid":
                        var shapAdjst = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd", "attrs", "fmla"]);
                        var adjst_val = 0.25;
                        var max_adj_const = 0.7407;
                        if(shapAdjst !== undefined){
                            var adjst = parseInt(shapAdjst.substr(4)) * 96 / 9144000;
                            adjst_val = (adjst*0.5)/max_adj_const;
                        // console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nadjst_val: "+adjst_val);
                        }
                        result += " <polygon points='"+(w*adjst_val)+" 0,0 " + h + ","+w+" "+h+","+(1-adjst_val)*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";    
                        break;
                    case "parallelogram":
                        var shapAdjst = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd", "attrs", "fmla"]);
                        var adjst_val = 0.25;
                        var max_adj_const;
                        if(w > h){
                            max_adj_const = w/h;
                        }else{
                            max_adj_const = h/w;
                        }
                        if(shapAdjst !== undefined){
                            var adjst = parseInt(shapAdjst.substr(4)) /100000;
                            adjst_val = adjst/max_adj_const;
                        //console.log("w: "+w+"\nh: "+h+"\nadjst: "+adjst_val+"\nmax_adj_const: "+max_adj_const);
                        }
                        result += " <polygon points='"+adjst_val*w+" 0,0 " + h + ","+(1-adjst_val)*w+" "+h+","+w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";    
                        break;

                        break;
                    case "pentagon":
                        result += " <polygon points='" + (0.5*w) + " 0,0 " + (0.375*h) + "," + (0.15*w)+" "+h+"," + 0.85*w + " " + h + "," + w + " " + 0.375*h + "' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "hexagon":
                        var shapAdjst_ary = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);
                        var shapAdjst = undefined;
                        for(var i=0; i<shapAdjst_ary.length; i++){
                            if( getTextByPathList(shapAdjst_ary[i],["attrs","name"]) =="adj"){
                                shapAdjst = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                            }
                        }
                        var adjst_val = 0.25;
                        var max_adj_const = 0.62211;
                    
                        if(shapAdjst !== undefined){
                            var adjst = parseInt(shapAdjst.substr(4)) * 96 / 9144000;
                            adjst_val = (adjst*0.5)/max_adj_const;
                            //console.log("w: "+w+"\nh: "+h+"\nadjst: "+adjst_val);
                        }
                        result += " <polygon points='"+(w*adjst_val)+" 0,0 " + (h/2) + ","+(w*adjst_val)+" "+h+","+(1-adjst_val)*w+" "+h+","+w+" "+(h/2)+","+(1-adjst_val)*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";    
                        break;
                    case "heptagon":
                        result += " <polygon points='" + (0.5*w) + " 0,"+w/8+" " + h/4 + ",0 "+(5/8)*h+"," + w/4 + " " + h + "," + (3/4)*w + " " +h +","+
                        w+" "+(5/8)*h+","+(7/8)*w+" "+h/4+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "octagon":
                        var shapAdjst = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd", "attrs", "fmla"]);
                        var adj1 = 0.25;
                        if(shapAdjst !== undefined){
                            adj1 = parseInt(shapAdjst.substr(4)) /100000;
                            
                        }
                        var adj2 = (1-adj1);
                        //console.log("adj1: "+adj1+"\nadj2: "+adj2);
                        result += " <polygon points='"+adj1*w+" 0,0 " + adj1*h + ",0 "+ adj2*h+","+adj1*w+" "+h+","+adj2*w+" "+h+","+
                        w+" "+adj2*h+","+w+" "+adj1*h+","+adj2*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";    

                        break;            
                    case "decagon":
                        result += " <polygon points='"+(3/8)*w+" 0,"+w/8+" " + h/8 + ",0 "+ h/2+","+w/8+" "+(7/8)*h+","+(3/8)*w+" "+h+","+
                            (5/8)*w+" "+h+","+(7/8)*w+" "+(7/8)*h+","+w+" "+h/2+","+(7/8)*w+" "+h/8+","+(5/8)*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "dodecagon":
                        result += " <polygon points='"+(3/8)*w+" 0,"+w/8+" " + h/8 + ",0 "+ (3/8)*h+ ",0 "+ (5/8)*h+","+w/8+" "+(7/8)*h+","+(3/8)*w+" "+h+","+
                            (5/8)*w+" "+h+","+(7/8)*w+" "+(7/8)*h+","+w+" "+(5/8)*h+","+w+" "+(3/8)*h+","+(7/8)*w+" "+h/8+","+(5/8)*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star4":
                    case "star5":
                    case "star6":
                    case "star7":
                    case "star8":    
                    case "star10":
                    case "star12":
                    case "star16":
                    case "star24":
                    case "star32":
                        var shapAdjst = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);//[0]["attrs"]["fmla"];
                        var starNum = shapType.substr(4);
                        var shapAdjst1 , adj;
                        switch(starNum){
                            case "4":
                                adj = 30;
                                break;
                            case "5":
                                adj = 40;
                                break;
                            case "6":
                                adj = 60;
                                break;
                            case "7":
                                adj = 70;
                                break;
                            case "8":
                                adj = 77;
                                break;
                            case "10":
                                adj = 86;
                                break;
                            case "12":
                            case "16":
                            case "24":
                            case "32":
                                adj = 75;
                                break;
                        }
                        if(shapAdjst !== undefined){
                            shapAdjst1 = getTextByPathList(shapAdjst, ["attrs", "fmla"]);
                            if(shapAdjst1 === undefined){
                                shapAdjst1 = shapAdjst[0]["attrs"]["fmla"];
                            }
                            if(shapAdjst1 !== undefined){
                                adj = 2*parseInt(shapAdjst1.substr(4)) /1000;
                            }
                        }
                        
                        var points = shapeStar(adj,starNum);
                        result += " <polygon points='"+points+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "pie":
                    case "pieWedge":
                        var shapAdjst = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);
                        var adj1, adj2 ,pieSize, shapAdjst1 , shapAdjst2;
                        if(shapType == "pie"){
                            adj1 = 0;
                            adj2 = 270;
                            pieSize = h;
                        }else{ //pieWedge
                            adj1 = 180;
                            adj2 = 270;
                            pieSize = 2*h;
                        }
                        if(shapAdjst !== undefined){
                            shapAdjst1 = getTextByPathList(shapAdjst, ["attrs", "fmla"]);
                            shapAdjst2 = shapAdjst1;
                            if(shapAdjst1 === undefined){
                                shapAdjst1 = shapAdjst[0]["attrs"]["fmla"];
                                shapAdjst2 = shapAdjst[1]["attrs"]["fmla"];
                            }
                            if(shapAdjst1 !== undefined){
                                adj1 = parseInt(shapAdjst1.substr(4)) /60000;
                            }
                            if(shapAdjst2 !== undefined){
                                adj2 = parseInt(shapAdjst2.substr(4)) /60000;
                            }
                        }
                        var pieVals = shapePie(pieSize,adj1,adj2);
                        //console.log("shapAdjst: ",shapAdjst,"\nadj1: ",adj1,"\nadj2: ",adj2,"\npieVals: ",pieVals);
                        result += "<path   d='"+pieVals[0]+"' transform='"+pieVals[1]+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "gear6":
                    case "gear9":
                        txtRotate = 0;
                        var gearNum = shapType.substr(4) , d;
                        if(gearNum == "6"){
                            d = shapeGear(w,h/3.5,parseInt(gearNum));
                        }else{ //gearNum=="9"
                            d = shapeGear(w,h/3.5,parseInt(gearNum));
                        }
                        result += "<path   d='"+d+"' transform='rotate(20,"+(3/7)*h+","+(3/7)*h+")' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bentConnector3":
                        var shapAdjst = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd", "attrs", "fmla"]);
                        //console.log("isFlipV: "+String(isFlipV)+"\nshapAdjst: "+shapAdjst)
                        var shapAdjst_val = 0.5;
                        if(shapAdjst !== undefined){
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) /100000;
                            //console.log("isFlipV: "+String(isFlipV)+"\nshapAdjst: "+shapAdjst+"\nshapAdjst_val: "+shapAdjst_val);
                            if(isFlipV){
                                result += " <polyline points='"+w+" 0," + ((1-shapAdjst_val)*w) + " 0,"+((1-shapAdjst_val)*w)+" "+h+",0 "+h+"' fill='transparent'" + 
                                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' "; 
                            }else{
                                result += " <polyline points='0 0,"+(shapAdjst_val)*w+" 0," + (shapAdjst_val)*w + " "+h+","+w+" "+h+"' fill='transparent'" + 
                                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                            }
                            if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                                result += "marker-start='url(#markerTriangle_"+shpId+")' ";
                            }
                            if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                                result += "marker-end='url(#markerTriangle_"+shpId+")' ";
                            }
                            result += "/>";                        
                        }                
                        break;
                    case "plus":
                        var shapAdjst = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd", "attrs", "fmla"]);
                        var adj1 = 0.25;
                        if(shapAdjst !== undefined){
                            adj1 = parseInt(shapAdjst.substr(4)) /100000;
                            
                        }
                        var adj2 = (1-adj1);
                        result += " <polygon points='"+adj1*w+" 0,"+adj1*w+" " + adj1*h + ",0 "+adj1*h+",0 "+adj2*h+","+
                                    adj1*w+" "+adj2*h+","+adj1*w+" "+h+","+adj2*w+" "+h+","+adj2*w+" "+adj2*h+","+w+" "+adj2*h+","+
                                    +w+" "+adj1*h+","+adj2*w+" "+adj1*h+","+adj2*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";    
                        //console.log((!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")")) 
                        break;
                    case "line":
                    case "straightConnector1":
                    case "bentConnector4":
                    case "bentConnector5":
                    case "curvedConnector2":
                    case "curvedConnector3":
                    case "curvedConnector4":
                    case "curvedConnector5":
                        if (isFlipV) {
                            result += "<line x1='" + w + "' y1='0' x2='0' y2='" + h + "' stroke='" + border.color + 
                                        "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                        } else {
                            result += "<line x1='0' y1='0' x2='" + w + "' y2='" + h + "' stroke='" + border.color + 
                                        "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                        }
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_"+shpId+")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_"+shpId+")' ";
                        }
                        result += "/>";
                        break;
                    case "rightArrow":
                        var shapAdjst_ary = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);
                        var sAdj1,sAdj1_val = 0.25;//0.5;
                        var sAdj2,sAdj2_val = 0.5;
                        var max_sAdj2_const = w/h;
                        if(shapAdjst_ary !== undefined){
                            for(var i=0; i<shapAdjst_ary.length; i++){
                                var sAdj_name = getTextByPathList(shapAdjst_ary[i],["attrs","name"]);
                                if(sAdj_name =="adj1"){
                                    sAdj1 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    sAdj1_val = 0.5-(parseInt(sAdj1.substr(4)) /200000);
                                }else if(sAdj_name =="adj2"){
                                    sAdj2 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) /100000;
                                    sAdj2_val = 1 - ((sAdj2_val2)/max_sAdj2_const);
                                }
                            }
                        }
                    //console.log("w: "+w+"\nh: "+h+"\nsAdj1: "+sAdj1_val+"\nsAdj2: "+sAdj2_val);
                        
                        result += " <polygon points='"+w+" "+h/2+","+sAdj2_val*w+" 0," +sAdj2_val*w+" "+sAdj1_val*h+",0 "+sAdj1_val*h+
                                    ",0 "+(1-sAdj1_val)*h+","+sAdj2_val*w+" "+(1-sAdj1_val)*h+", "+sAdj2_val*w+" "+h+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";               
                        break;
                    case "leftArrow":
                        var shapAdjst_ary = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);
                        var sAdj1,sAdj1_val = 0.25;//0.5;
                        var sAdj2,sAdj2_val = 0.5;
                        var max_sAdj2_const = w/h;
                        if(shapAdjst_ary !== undefined){
                            for(var i=0; i<shapAdjst_ary.length; i++){
                                var sAdj_name = getTextByPathList(shapAdjst_ary[i],["attrs","name"]);
                                if(sAdj_name =="adj1"){
                                    sAdj1 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    sAdj1_val = 0.5-(parseInt(sAdj1.substr(4)) /200000);
                                }else if(sAdj_name =="adj2"){
                                    sAdj2 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) /100000;
                                    sAdj2_val = (sAdj2_val2)/max_sAdj2_const;
                                }
                            }
                        }
                        //console.log("w: "+w+"\nh: "+h+"\nsAdj1: "+sAdj1_val+"\nsAdj2: "+sAdj2_val);

                        result += " <polygon points='0 "+h/2+","+sAdj2_val*w+" "+h+"," +sAdj2_val*w+" "+(1-sAdj1_val)*h+","+w+" "+(1-sAdj1_val)*h+
                                    ","+w+" "+sAdj1_val*h+","+sAdj2_val*w+" "+sAdj1_val*h+", "+sAdj2_val*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "downArrow":
                        var shapAdjst_ary = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);
                        var sAdj1,sAdj1_val = 0.25;//0.5;
                        var sAdj2,sAdj2_val = 0.5;
                        var max_sAdj2_const = h/w;
                        if(shapAdjst_ary !== undefined){
                            for(var i=0; i<shapAdjst_ary.length; i++){
                                var sAdj_name = getTextByPathList(shapAdjst_ary[i],["attrs","name"]);
                                if(sAdj_name =="adj1"){
                                    sAdj1 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) /200000;
                                }else if(sAdj_name =="adj2"){
                                    sAdj2 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) /100000;
                                    sAdj2_val = (sAdj2_val2)/max_sAdj2_const;
                                }
                            }
                        }
                    // console.log("w: "+w+"\nh: "+h+"\nsAdj1: "+sAdj1_val+"\nsAdj2: "+sAdj2_val);
                        
                        result += " <polygon points='"+(0.5-sAdj1_val)*w+" 0,"+(0.5-sAdj1_val)*w+" "+(1-sAdj2_val)*h+",0 " +(1-sAdj2_val)*h+","+(w/2)+" "+h+
                                    ","+w+" "+(1-sAdj2_val)*h+","+(0.5+sAdj1_val)*w+" "+(1-sAdj2_val)*h+", "+(0.5+sAdj1_val)*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";               
                        break; 
                    case "upArrow":
                        var shapAdjst_ary = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);
                        var sAdj1,sAdj1_val = 0.25;//0.5;
                        var sAdj2,sAdj2_val = 0.5;
                        var max_sAdj2_const = h/w;
                        if(shapAdjst_ary !== undefined){
                            for(var i=0; i<shapAdjst_ary.length; i++){
                                var sAdj_name = getTextByPathList(shapAdjst_ary[i],["attrs","name"]);
                                if(sAdj_name =="adj1"){
                                    sAdj1 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) /200000;
                                }else if(sAdj_name =="adj2"){
                                    sAdj2 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) /100000;
                                    sAdj2_val = (sAdj2_val2)/max_sAdj2_const;
                                }
                            }
                        }
                        result += " <polygon points='"+(w/2)+" 0,0 "+sAdj2_val*h+"," + (0.5-sAdj1_val)*w + " "+sAdj2_val*h+","+(0.5-sAdj1_val)*w+" "+h+
                                    ","+(0.5+sAdj1_val)*w+" "+h+","+(0.5+sAdj1_val)*w+" "+sAdj2_val*h+", "+w+" "+sAdj2_val*h+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";               
                        break;
                    case "leftRightArrow":
                        var shapAdjst_ary = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);
                        var sAdj1,sAdj1_val = 0.25;
                        var sAdj2,sAdj2_val = 0.25;
                        var max_sAdj2_const = w/h;
                        if(shapAdjst_ary !== undefined){
                            for(var i=0; i<shapAdjst_ary.length; i++){
                                var sAdj_name = getTextByPathList(shapAdjst_ary[i],["attrs","name"]);
                                if(sAdj_name =="adj1"){
                                    sAdj1 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    sAdj1_val = 0.5-(parseInt(sAdj1.substr(4)) /200000);
                                }else if(sAdj_name =="adj2"){
                                    sAdj2 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) /100000;
                                    sAdj2_val = (sAdj2_val2)/max_sAdj2_const;
                                }
                            }
                        }
                        //console.log("w: "+w+"\nh: "+h+"\nsAdj1: "+sAdj1_val+"\nsAdj2: "+sAdj2_val);

                        result += " <polygon points='0 "+h/2+","+sAdj2_val*w+" "+h+"," +sAdj2_val*w+" "+(1-sAdj1_val)*h+","+(1-sAdj2_val)*w+" "+(1-sAdj1_val)*h+
                                    ","+(1-sAdj2_val)*w+" "+h+","+w+" "+h/2+", "+(1-sAdj2_val)*w+" 0,"+(1-sAdj2_val)*w+" "+sAdj1_val*h+","+
                                    sAdj2_val*w+" "+sAdj1_val*h+","+sAdj2_val*w+" 0' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "upDownArrow":
                        var shapAdjst_ary = getTextByPathList(node, ["p:spPr", "a:prstGeom","a:avLst","a:gd"]);
                        var sAdj1,sAdj1_val = 0.25;
                        var sAdj2,sAdj2_val = 0.25;
                        var max_sAdj2_const = h/w;
                        if(shapAdjst_ary !== undefined){
                            for(var i=0; i<shapAdjst_ary.length; i++){
                                var sAdj_name = getTextByPathList(shapAdjst_ary[i],["attrs","name"]);
                                if(sAdj_name =="adj1"){
                                    sAdj1 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    sAdj1_val = 0.5-(parseInt(sAdj1.substr(4)) /200000);
                                }else if(sAdj_name =="adj2"){
                                    sAdj2 = getTextByPathList(shapAdjst_ary[i],["attrs","fmla"]);
                                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) /100000;
                                    sAdj2_val = (sAdj2_val2)/max_sAdj2_const;
                                }
                            }
                        }
                        //console.log("w: "+w+"\nh: "+h+"\nsAdj1: "+sAdj1_val+"\nsAdj2: "+sAdj2_val);

                        result += " <polygon points='"+w/2+" 0,0 "+sAdj2_val*h+"," +sAdj1_val*w+" "+sAdj2_val*h+","+sAdj1_val*w+" "+(1-sAdj2_val)*h+
                                    ",0 "+(1-sAdj2_val)*h+","+w/2+" "+h+", "+w+" "+(1-sAdj2_val)*h+","+(1-sAdj1_val)*w+" "+(1-sAdj2_val)*h+","+
                                    (1-sAdj1_val)*w+" "+sAdj2_val*h+","+w+" "+sAdj2_val*h+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bentArrow":
                    case "bentUpArrow":
                    case "stripedRightArrow":
                    case "quadArrow":
                    case "circularArrow":
                    case "swooshArrow":
                    case "leftRightUpArrow":
                    case "leftUpArrow":
                    case "leftCircularArrow":
                    case "notchedRightArrow":
                    case "curvedDownArrow":
                    case "curvedLeftArrow":
                    case "curvedRightArrow":
                    case "curvedUpArrow":
                    case "uturnArrow":
                    case "leftRightCircularArrow":
                        break;
                    case undefined:
                    default:
                        console.warn("Undefine shape type.");
                }
                
                result += "</svg>";
                
                result += "<div class='block content " + getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
                        "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                        "' style='" + 
                            getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                            getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                            " z-index: " + order + ";" +
                            "transform: rotate(" +txtRotate+ "deg);"+
                        "'>";
                
                // TextBody
                if (node["p:txBody"] !== undefined) {
                    result += genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                }
                result += "</div>";
            }else if(custShapType !== undefined){
                //custGeom here - Amir ///////////////////////////////////////////////////////
                //http://officeopenxml.com/drwSp-custGeom.php
                var pathLstNode = getTextByPathList(custShapType, ["a:pathLst"]);
                var pathNode = getTextByPathList(pathLstNode, ["a:path", "attrs"]);
                var maxX = parseInt(pathNode["w"]) * 96 / 914400;
                var maxY = parseInt(pathNode["h"]) * 96 / 914400;
                //console.log("w = "+w+"\nh = "+h+"\nmaxX = "+maxX +"\nmaxY = " + maxY);
                //cheke if it is close shape
                var closeNode = getTextByPathList(pathLstNode, ["a:path","a:close"]);
                var startPoint = getTextByPathList(pathLstNode, ["a:path","a:moveTo","a:pt","attrs"]);
                var spX = parseInt(startPoint["x"]) * 96 / 914400;
                var spY = parseInt(startPoint["y"]) * 96 / 914400;
                var d = "M"+spX+","+spY;
                var pathNodes =  getTextByPathList(pathLstNode, ["a:path"]);
                var lnToNodes = pathNodes["a:lnTo"];
                var cubicBezToNodes = pathNodes["a:cubicBezTo"];
                var sortblAry = [];
                if(lnToNodes !== undefined){
                    Object.keys(lnToNodes).forEach(function(key) {
                        var lnToPtNode = lnToNodes[key]["a:pt"];
                        if(lnToPtNode !== undefined){
                            Object.keys(lnToPtNode).forEach(function(key2) {
                                var ptObj = {};
                                var lnToNoPt = lnToPtNode[key2];
                                var ptX = lnToNoPt["attrs","x"];
                                var ptY = lnToNoPt["attrs","y"];
                                var ptOrdr = lnToNoPt["attrs","order"];
                                ptObj.type = "lnto";
                                ptObj.order = ptOrdr;
                                ptObj.x = ptX;
                                ptObj.y = ptY;
                                sortblAry.push(ptObj);
                                //console.log(key2, lnToNoPt);
                            
                            });
                        }
                    });
                    
                }
                if(cubicBezToNodes !== undefined){
                    Object.keys(cubicBezToNodes).forEach(function(key) {
                        //console.log("cubicBezTo["+key+"]:");
                        var cubicBezToPtNodes = cubicBezToNodes[key]["a:pt"];
                        if(cubicBezToPtNodes !== undefined){
                            Object.keys(cubicBezToPtNodes).forEach(function(key2) {
                                //console.log("cubicBezTo["+key+"]pt["+key2+"]:");
                                var cubBzPts = cubicBezToPtNodes[key2];
                                Object.keys(cubBzPts).forEach(function(key3) {
                                    //console.log(key3, cubBzPts[key3]);
                                    var ptObj = {};
                                    var cubBzPt = cubBzPts[key3];
                                    var ptX = cubBzPt["attrs","x"];
                                    var ptY = cubBzPt["attrs","y"];
                                    var ptOrdr = cubBzPt["attrs","order"];
                                    ptObj.type = "cubicBezTo";
                                    ptObj.order = ptOrdr;
                                    ptObj.x = ptX;
                                    ptObj.y = ptY;
                                    sortblAry.push(ptObj);                            
                                });
                            });
                        }
                    });
                }
                var sortByOrder = sortblAry.slice(0);
                sortByOrder.sort(function(a,b) {
                    return a.order - b.order;
                });
                //console.log(sortByOrder);
                var k = 0;
                while(k<sortByOrder.length){
                    if(sortByOrder[k].type=="lnto"){
                        var Lx = parseInt(sortByOrder[k].x) * 96 / 914400;
                        var Ly = parseInt(sortByOrder[k].y) * 96 / 914400;
                        d += "L" + Lx + "," + Ly;
                        k++;
                    }else{ //"cubicBezTo"
                        var Cx1 = parseInt(sortByOrder[k].x) * 96 / 914400;
                        var Cy1 = parseInt(sortByOrder[k].y) * 96 / 914400;
                        var Cx2 = parseInt(sortByOrder[k+1].x) * 96 / 914400;
                        var Cy2 = parseInt(sortByOrder[k+1].y) * 96 / 914400;
                        var Cx3 = parseInt(sortByOrder[k+2].x) * 96 / 914400;
                        var Cy3 = parseInt(sortByOrder[k+2].y) * 96 / 914400; 

                        d += "C" + Cx1 + "," + Cy1 +" "+ Cx2 + "," + Cy2 + " " + Cx3 + "," + Cy3;
                        k += 3 ; 
                    }
                }
                result += "<path d='" + d + "' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                if(closeNode !== undefined){
                    //console.log("Close shape");
                    result += "/>";
                }else{
                    //console.log("Open shape");
                    //check and add "marker-start" and "marker-end"
                    if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                        result += "marker-start='url(#markerTriangle_"+shpId+")' ";
                    }
                    if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                        result += "marker-end='url(#markerTriangle_"+shpId+")' ";
                    } 
                    result += "/>";
                    
                }
                
                result += "</svg>";
                
                result += "<div class='block content " + getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
                        "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                        "' style='" + 
                            getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                            getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                            " z-index: " + order + ";" +
                            "transform: rotate(" +txtRotate+ "deg);"+
                        "'>";
                
                // TextBody
                if (node["p:txBody"] !== undefined) {
                    result += genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                }
                result += "</div>";

            // result = "";
            } else {
                
                result += "<div class='block content " + getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
                        "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                        "' style='" + 
                            getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                            getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                            getBorder(node, false) +
                            getShapeFill(node, false,warpObj) +
                            " z-index: " + order + ";" +
                            "transform: rotate(" +txtRotate+ "deg);"+
                        "'>";
                
                // TextBody
                if (node["p:txBody"] !== undefined) {
                    result += genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                }
                result += "</div>";
                
            }
            
            return result;
        }
        function shapeStar(adj, starNum){
            var innerRadius = adj; /*1-100*/
            var outerRadius = 100;//star.outerRadius;
            var numPoints = starNum;/*1-100*/
            var center = Math.max(innerRadius, outerRadius);
            var angle  = Math.PI / numPoints;
            var points = [];  
            
            for (var i = 0; i < numPoints * 2; i++) {
              var radius = i & 1 ? innerRadius : outerRadius;  
              points.push(center + radius * Math.sin(i * angle));
              points.push(center - radius * Math.cos(i * angle));
            }
            
            return points;            
        }
        function shapePie(pieSize,adj1,adj2){
            var pieVal = parseInt(adj2);
            var piAngle = parseInt(adj1);
            var	size = parseInt(pieSize),
            radius = (size / 2),
            value = pieVal - piAngle;
            if(value < 0){
                value = 360 + value;
            }
            //console.log("value: ",value)      
            value = Math.min(Math.max(value, 0), 360);
                
            //calculate x,y coordinates of the point on the circle to draw the arc to. 
            var x = Math.cos((2 * Math.PI)/(360/value));
            var y = Math.sin((2 * Math.PI)/(360/value));
            
            //should the arc go the long way round?
            var longArc = (value <= 180) ? 0 : 1;

            //d is a string that describes the path of the slice.
            var d = "M" + radius + "," + radius + " L" + radius + "," + 0 + " A" + radius + "," + radius + " 0 " + longArc + ",1 " + (radius + y*radius) + "," + (radius - x*radius) + " z";	
            var rot = "rotate("+(piAngle-270)+", "+radius+", "+radius+")";

            return [d,rot];
        }
        function shapeGear(w,h,points) {
              var innerRadius = h;//gear.innerRadius;
              var outerRadius = 1.5*innerRadius; 
              var cx = outerRadius;//Math.max(innerRadius, outerRadius),                   // center x
                cy = outerRadius;//Math.max(innerRadius, outerRadius),                    // center y
                notches =  points,//gear.points,                      // num. of notches
                radiusO = outerRadius,                    // outer radius
                radiusI = innerRadius,                    // inner radius
                taperO  = 50,                     // outer taper %
                taperI  = 35,                     // inner taper %
            
                // pre-calculate values for loop
            
                pi2     = 2 * Math.PI,            // cache 2xPI (360deg)
                angle   = pi2 / (notches * 2),    // angle between notches
                taperAI = angle * taperI * 0.005, // inner taper offset (100% = half notch)
                taperAO = angle * taperO * 0.005, // outer taper offset
                a       = angle,                  // iterator (angle)
                toggle  = false;
              // move to starting point
            var d = " M"+(cx + radiusO * Math.cos(taperAO))+" "+ (cy + radiusO * Math.sin(taperAO));
            
            // loop
            for (; a <= pi2+angle; a += angle) {
                // draw inner to outer line
                if (toggle) {
                    d +=  " L"+(cx + radiusI * Math.cos(a - taperAI)) + "," + (cy + radiusI * Math.sin(a - taperAI));
                    d +=  " L"+(cx + radiusO * Math.cos(a + taperAO)) + "," + (cy + radiusO * Math.sin(a + taperAO));
                }else { // draw outer to inner line
                    d +=  " L"+(cx + radiusO * Math.cos(a - taperAO)) + "," +  (cy + radiusO * Math.sin(a - taperAO)); // outer line
                    d +=  " L"+(cx + radiusI * Math.cos(a + taperAI)) + "," +  (cy + radiusI * Math.sin(a + taperAI));// inner line
                               
                }
                // switch level
                toggle = !toggle;
            }
            // close the final line
            d += " ";
            return d;
        }
        /*
        function shapePolygon(sidesNum) {
            var sides  = sidesNum;
            var radius = 100;
            var angle  = 2 * Math.PI / sides;
            var points = []; 
            
            for (var i = 0; i < sides; i++) {
                points.push(radius + radius * Math.sin(i * angle));
                points.push(radius - radius * Math.cos(i * angle));
            }
            
            return points;
        }
        */       
        function processPicNode(node, warpObj) {
            
            
            var order = node["attrs"]["order"];
            
            var rid = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
            var imgName = warpObj["slideResObj"][rid]["target"];
            var imgFileExt = extractFileExtension(imgName).toLowerCase();
            var zip = warpObj["zip"];
            var imgArrayBuffer = zip.file(imgName).asArrayBuffer();
            var mimeType = "";
            var xfrmNode = node["p:spPr"]["a:xfrm"];
            ///////////////////////////////////////Amir//////////////////////////////
            var rotate = angleToDegrees(node["p:spPr"]["a:xfrm"]["attrs"]["rot"]);
            //////////////////////////////////////////////////////////////////////////
            mimeType = getImageMimeType(imgFileExt);
            return "<div class='block content' style='" + getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) +
                    " z-index: " + order + ";" +
                    "transform: rotate(" +rotate+ "deg);"+
                    "'><img src='data:" + mimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/></div>";
        }

        function processGraphicFrameNode(node, warpObj) {
            
            var result = "";
            var graphicTypeUri = getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);
            
            switch (graphicTypeUri) {
                case "http://schemas.openxmlformats.org/drawingml/2006/table":
                    result = genTable(node, warpObj);
                    break;
                case "http://schemas.openxmlformats.org/drawingml/2006/chart":
                    result = genChart(node, warpObj);
                    break;
                case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
                    result = genDiagram(node, warpObj);
                    break;
                default:
            }
            
            return result;
        }

        function processSpPrNode(node, warpObj) {
            
            /*
            * 2241 <xsd:complexType name="CT_ShapeProperties">
            * 2242   <xsd:sequence>
            * 2243     <xsd:element name="xfrm" type="CT_Transform2D"  minOccurs="0" maxOccurs="1"/>
            * 2244     <xsd:group   ref="EG_Geometry"                  minOccurs="0" maxOccurs="1"/>
            * 2245     <xsd:group   ref="EG_FillProperties"            minOccurs="0" maxOccurs="1"/>
            * 2246     <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
            * 2247     <xsd:group   ref="EG_EffectProperties"          minOccurs="0" maxOccurs="1"/>
            * 2248     <xsd:element name="scene3d" type="CT_Scene3D"   minOccurs="0" maxOccurs="1"/>
            * 2249     <xsd:element name="sp3d" type="CT_Shape3D"      minOccurs="0" maxOccurs="1"/>
            * 2250     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
            * 2251   </xsd:sequence>
            * 2252   <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
            * 2253 </xsd:complexType>
            */
            
            // TODO:
        }

        function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {


            var text = "";
            var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
            
            if (textBodyNode === undefined) {
                return text;
            }
            //rtl : <p:txBody>
            //          <a:bodyPr wrap="square" rtlCol="1">
            
            //var rtlStr = "";
            if (textBodyNode["a:p"].constructor === Array) {
                // multi p
                for (var i=0; i<textBodyNode["a:p"].length; i++) {
                    var pNode = textBodyNode["a:p"][i];
                    var rNode = pNode["a:r"];
                    
                    //var isRTL = getTextDirection(pNode, type, slideMasterTextStyles);
                    //rtlStr = "";//"dir='"+isRTL+"'";

                    text += "<div  class='" + getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + "'>";
                    text += genBuChar(pNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);

                    if (rNode === undefined) {
                        // without r
                        text += genSpanElement(pNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                    } else if (rNode.constructor === Array) {
                        // with multi r
                        for (var j=0; j<rNode.length; j++) {
                            text += genSpanElement(rNode[j], spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                            //////////////////Amir////////////
                            if(pNode["a:br"] !== undefined){
                                text += "<br>";
                            }
                            //////////////////////////////////                    
                        }
                    } else {
                        // with one r
                        text += genSpanElement(rNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                    }
                    text += "</div>";
                }
            } else {
                // one p
                var pNode = textBodyNode["a:p"];
                var rNode = pNode["a:r"];

                //var isRTL = getTextDirection(pNode, type, slideMasterTextStyles);
                //rtlStr = "";//"dir='"+isRTL+"'";

                text += "<div class='slide-prgrph " + getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + "'>";
                text += genBuChar(pNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                if (rNode === undefined) {
                    // without r
                    text += genSpanElement(pNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                } else if (rNode.constructor === Array) {
                    // with multi r
                    for (var j=0; j<rNode.length; j++) {
                        text += genSpanElement(rNode[j], spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                        //////////////////Amir////////////
                        if(pNode["a:br"] !== undefined){
                            text += "<br>";
                        }
                        //////////////////////////////////
                    }
                } else {
                    // with one r
                    text += genSpanElement(rNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                }
                text += "</div>";
            }
            
            return text;
        }

        function genBuChar(node, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
            ///////////////////////////////////////Amir///////////////////////////////
            var sldMstrTxtStyles = warpObj["slideMasterTextStyles"];

            var rNode = getTextByPathList(node,["a:r"]);
            if(rNode !== undefined && rNode.constructor === Array){
                rNode = rNode[0];
            }
            var dfltBultColor,dfltBultSize,bultColor,bultSize;
            if (rNode !== undefined) {
                dfltBultColor = getFontColor(rNode, spNode, type, sldMstrTxtStyles);
                dfltBultSize = getFontSize(rNode, slideLayoutSpNode, slideMasterSpNode, type, sldMstrTxtStyles);       
            }else{
                dfltBultColor = getFontColor(node, spNode, type, sldMstrTxtStyles);
                dfltBultSize = getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, sldMstrTxtStyles);         
            }
            //console.log("Bullet Size: " + bultSize);
            
            var bullet = "";
            /////////////////////////////////////////////////////////////////

            
            var pPrNode = node["a:pPr"];
            
            //////////////////cheke if is rtl ///Amir ////////////////////////////////////
            var getRtlVal = getTextByPathList(pPrNode, ["attrs", "rtl"])
            var isRTL = false;
            if(getRtlVal !== undefined && getRtlVal=="1"){
                isRTL = true;
            }
            ////////////////////////////////////////////////////////////
            
            var lvl = parseInt( getTextByPathList(pPrNode, ["attrs", "lvl"]) );
            if (isNaN(lvl)) {
                lvl = 0;
            }
            
            var buChar = getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
            /////////////////////////////////Amir///////////////////////////////////
            var buType = "TYPE_NONE";
            var buNum = getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
            var buPic = getTextByPathList(pPrNode, ["a:buBlip"]);
            if(buChar !== undefined){
                buType = "TYPE_BULLET";
                // console.log("Bullet Chr to code: " + buChar.charCodeAt(0));
            }
            if(buNum !== undefined){
                buType = "TYPE_NUMERIC";
            }
            if(buPic !== undefined){
                buType = "TYPE_BULPIC";
            }

            if(buType != "TYPE_NONE"){
                var buFontAttrs = getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
            }
            //console.log("Bullet Type: " + buType);
            //console.log("NumericTypr: " + buNum);
            //console.log("buChar: " + (buChar === undefined?'':buChar.charCodeAt(0)));
            //get definde bullet COLOR
            var buClrNode = getTextByPathList(pPrNode,["a:buClr"]);
            var defBultColor = "NoNe";
            if(buClrNode !== undefined){
                defBultColor = getSolidFill(buClrNode);
            }else{
            // console.log("buClrNode: " + buClrNode);
            }

            if(defBultColor == "NoNe"){
                bultColor = dfltBultColor;
            }else{
                bultColor = "#" + defBultColor;
            }
            //get definde bullet SIZE
            var buFontSize;
            buFontSize = getTextByPathList(pPrNode, ["a:buSzPts", "attrs","val"]); //pt
            if(buFontSize !== undefined){
                bultSize = parseInt(buFontSize) / 100 +"pt";
            }else{
                buFontSize = getTextByPathList(pPrNode, ["a:buSzPct", "attrs","val"]);
                if(buFontSize !== undefined){
                    var prcnt = parseInt(buFontSize) /100000;
                    //dfltBultSize = XXpt
                    var dfltBultSizeNoPt = dfltBultSize.substr(0,dfltBultSize.length-2);
                    bultSize = prcnt*(parseInt(dfltBultSizeNoPt))+"pt";
                }else{
                    bultSize = dfltBultSize;
                }
            }
            ////////////////////////////////////////////////////////////////////////
            if (buType == "TYPE_BULLET") {
                //var buFontAttrs = getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
                if (buFontAttrs !== undefined) {
                    var marginLeft = parseInt( getTextByPathList(pPrNode, ["attrs", "marL"]) ) * 96 / 914400;
                    var marginRight = parseInt(buFontAttrs["pitchFamily"]);
                    if (isNaN(marginLeft)) {
                        marginLeft = 328600 * 96 / 914400;
                    }
                    if (isNaN(marginRight)) {
                        marginRight = 0;
                    }
                    var typeface = buFontAttrs["typeface"];

                    bullet =  "<span style='font-family: " + typeface + 
                            "; margin-left: " + marginLeft * lvl + "px" +
                            "; margin-right: " + marginRight + "px" +
                            ";color:" + bultColor + 
                            ";font-size:" + bultSize +";";
                    if(isRTL){
                        bullet += " float: right;  direction:rtl"; 
                    }
                    bullet +="'>" + buChar + "</span>";
                } else {
                    marginLeft = 328600 * 96 / 914400 * lvl;
                    
                    bullet = "<span style='margin-left: " + marginLeft + "px;'>" + buChar + "</span>";
                }
            } else if(buType == "TYPE_NUMERIC") { ///////////Amir///////////////////////////////
                if (buFontAttrs !== undefined) {
                    var marginLeft = parseInt( getTextByPathList(pPrNode, ["attrs", "marL"]) ) * 96 / 914400;
                    var marginRight = parseInt(buFontAttrs["pitchFamily"]);

                    if (isNaN(marginLeft)) {
                        marginLeft = 328600 * 96 / 914400;
                    }
                    if (isNaN(marginRight)) {
                        marginRight = 0;
                    }
                    //var typeface = buFontAttrs["typeface"];
                    
                    bullet =  "<span style='margin-left: " + marginLeft * lvl + "px" +
                            "; margin-right: " + marginRight + "px" +
                            ";color:" + bultColor + 
                            ";font-size:" + bultSize +";";
                    if(isRTL){
                        bullet += " float: right; direction:rtl;"; 
                    }else{
                        bullet += " float: left; direction:ltr;";
                    }
                    bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></span>";
                } else {
                    marginLeft = 328600 * 96 / 914400 * lvl;
                    bullet =  "<span style='margin-left: " + marginLeft + "px;";
                    if(isRTL){
                        bullet += " float: right; direction:rtl;"; 
                    }else{
                        bullet += " float: left; direction:ltr;";
                    }
                    bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></span>";
                }
            
            }else if(buType == "TYPE_BULPIC"){ //PIC BULLET
                var marginLeft = parseInt( getTextByPathList(pPrNode, ["attrs", "marL"]) ) * 96 / 914400;
                var marginRight = parseInt( getTextByPathList(pPrNode, ["attrs", "marR"]) ) * 96 / 914400;

                if (isNaN(marginRight)) {
                    marginRight = 0;
                }
                //console.log("marginRight: "+marginRight)
                //buPic
                if (isNaN(marginLeft)) {
                    marginLeft = 328600 * 96 / 914400;
                }else{
                    marginLeft = 0;
                }
                //var buPicId = getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
                var buPicId = getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                var svgPicPath = ""; 
                var buImg;
                if(buPicId !== undefined){
                    //svgPicPath = warpObj["slideResObj"][buPicId]["target"];
                    //buImg = warpObj["zip"].file(svgPicPath).asText();
                    //}else{
                    //buPicId = getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                    var imgPath =  warpObj["slideResObj"][buPicId]["target"];
                    var imgArrayBuffer = warpObj["zip"].file(imgPath).asArrayBuffer();
                    var imgExt = imgPath.split(".").pop();
                    var imgMimeType = getImageMimeType(imgExt);
                    buImg = "<img src='data:" + imgMimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/>"
                    //console.log("imgPath: "+imgPath+"\nimgMimeType: "+imgMimeType)
                }
                if(buPicId === undefined){
                    buImg = "&#8227;";
                }
                bullet =  "<span style='margin-left: " + marginLeft * lvl + "px" +
                            "; margin-right: " + marginRight + "px" +
                            ";width:" + bultSize +";display: inline-block; ";
                if(isRTL){
                    bullet += " float: right;direction:rtl"; 
                }             
                bullet += "'>"+buImg+"  </span>";
                //////////////////////////////////////////////////////////////////////////////////////
            } else {
                bullet =  "<span style='margin-left: " + 328600 * 96 / 914400 * lvl + "px" +
                            "; margin-right: " + 0 + "px;'></span>";
            }
            
            return bullet;
        }

        function  genSpanElement(node, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
            
            var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
            
            var text = node["a:t"];
            if (typeof text !== 'string') {
                text = getTextByPathList(node, ["a:fld", "a:t"]);
                if (typeof text !== 'string') {
                    text = "&nbsp;";
                }
            }
            //console.log("genSpanElement: ",node)
            var styleText = 
                "color:" + getFontColor(node, spNode, type, slideMasterTextStyles) + 
                ";font-size:" + getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + 
                ";font-family:" + getFontType(node, type, slideMasterTextStyles) + 
                ";font-weight:" + getFontBold(node, type, slideMasterTextStyles) + 
                ";font-style:" + getFontItalic(node, type, slideMasterTextStyles) + 
                ";text-decoration:" + getFontDecoration(node, type, slideMasterTextStyles) +
                ";text-align:" + getTextHorizontalAlign(node, type, slideMasterTextStyles) + 
                ";vertical-align:" + getTextVerticalAlign(node, type, slideMasterTextStyles) + 
                ";";
            //////////////////Amir///////////////
            var highlight = getTextByPathList(node, ["a:rPr", "a:highlight"]);
            if(highlight !== undefined){
                styleText += "background-color:#" + getSolidFill(highlight) +";";
                styleText += "Opacity:"+ getColorOpacity(highlight) + ";";
            }
            ///////////////////////////////////////////
            var cssName = "";
            
            if (styleText in styleTable) {
                cssName = styleTable[styleText]["name"];
            } else {
                cssName = "_css_" + (Object.keys(styleTable).length + 1);
                styleTable[styleText] = {
                    "name": cssName,
                    "text": styleText
                };
            }
            
            var linkID = getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
            //get link colors : TODO
            if (linkID !== undefined) {
                var linkURL = warpObj["slideResObj"][linkID]["target"];
                return "<span class='text-block " + cssName + "'><a href='" + linkURL + "' target='_blank'>" + text.replace(/\s/i, "&nbsp;") + "</a></span>";
            } else {
                return "<span class='text-block " + cssName + "'>" + text.replace(/\s/i, "&nbsp;") + "</span>";
            }
            
        }

        function genGlobalCSS() {
            var cssText = "";
            for (var key in styleTable) {
                cssText += "div ." + styleTable[key]["name"] + "{" + styleTable[key]["text"] + "}\n"; //section > div
            }
            return cssText;
        }

        function genTable(node, warpObj) {
            
            var order = node["attrs"]["order"];
            var tableNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
            var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
            /////////////////////////////////////////Amir////////////////////////////////////////////////
            var getTblPr = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl","a:tblPr"]);
            var getColsGrid = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl","a:tblGrid","a:gridCol"]);
            var tblDir = "";
            if(getTblPr !== undefined){
                var isRTL = getTblPr["attrs"]["rtl"];
                tblDir = (isRTL==1?"dir=rtl":"dir=ltr");
            }
            var firstRowAttr =  getTblPr["attrs"]["firstRow"]; //associated element <a:firstRow> in the table styles
            var firstColAttr =  getTblPr["attrs"]["firstCol"]; //associated element <a:firstCol> in the table styles
            var lastRowAttr =  getTblPr["attrs"]["lastRow"]; //associated element <a:lastRow> in the table styles
            var lastColAttr =  getTblPr["attrs"]["lastCol"]; //associated element <a:lastCol> in the table styles
            var bandRowAttr =  getTblPr["attrs"]["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
            var bandColAttr =  getTblPr["attrs"]["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
            //console.log(firstColAttr);
            ////////////////////////////////////////////////////////////////////////////////////////////
            var tableHtml = "<table "+tblDir+" style='border-collapse: collapse;" + getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) + " z-index: " + order + ";'>";
            
            var trNodes = tableNode["a:tr"];
            if (trNodes.constructor === Array) {
                for (var i=0; i<trNodes.length; i++) {
                    //////////////rows Style ////////////Amir
                    var rowHeightParam = trNodes[i]["attrs"]["h"];
                    var rowHeight = 0;
                    var rowsStyl = "";
                    if(rowHeightParam !== undefined){
                        rowHeight = parseInt(rowHeightParam) * 96 / 914400;
                        rowsStyl += "height:"+rowHeight+"px;";
                        //tableHtml += "<tr style='height:"+rowHeight+"px;'>";
                    }
                    
                    //get from Theme (tableStyles.xml) TODO 
                    //get tableStyleId = a:tbl => a:tblPr => a:tableStyleId
                    var thisTblStyle;
                    var tbleStyleId = getTblPr["a:tableStyleId"];
                    if(tbleStyleId !== undefined){
                        //get Style from tableStyles.xml by {var tbleStyleId}
                        //table style object : tableStyles
                        var tbleStylList = tableStyles["a:tblStyleLst"]["a:tblStyle"];
                        
                        for(var k=0;k<tbleStylList.length;k++){
                            if(tbleStylList[k]["attrs"]["styleId"] == tbleStyleId){
                                thisTblStyle = tbleStylList[k];
                            }
                        }
                    }
                        //console.log(thisTblStyle);
                    if(i==0 && firstRowAttr !== undefined){
                        var fillColor="fff";
                        var colorOpacity = 1;
                        if(thisTblStyle["a:firstRow"] !==undefined){
                            var bgFillschemeClr =  getTextByPathList(thisTblStyle, ["a:firstRow","a:tcStyle","a:fill","a:solidFill"]);
                            if(bgFillschemeClr !==undefined){
                                fillColor = getSolidFill(bgFillschemeClr);
                                colorOpacity = getColorOpacity(bgFillschemeClr);
                            }
                            //console.log(thisTblStyle["a:firstRow"])
                            
                            //borders color
                            //borders Width
                            var borderStyl = getTextByPathList(thisTblStyle,["a:firstRow","a:tcStyle","a:tcBdr"]);
                            if(borderStyl !== undefined){
                                var row_borders = getTableBorders(borderStyl);
                                rowsStyl += row_borders;
                            }
                            //console.log(thisTblStyle["a:firstRow"])
                            
                            //Text Style - TODO
                            var rowTxtStyl = getTextByPathList(thisTblStyle,["a:firstRow","a:tcTxStyle"]);
                            if(rowTxtStyl !== undefined){
                                /*
                            var styleText = 
                                "color:" + getFontColor(node, type, slideMasterTextStyles) + 
                                ";font-size:" + getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + 
                                ";font-family:" + getFontType(node, type, slideMasterTextStyles) + 
                                ";font-weight:" + getFontBold(node, type, slideMasterTextStyles) + 
                                ";font-style:" + getFontItalic(node, type, slideMasterTextStyles) + 
                                ";text-decoration:" + getFontDecoration(node, type, slideMasterTextStyles) +
                                ";text-align:" + getTextHorizontalAlign(node, type, slideMasterTextStyles) + 
                                ";vertical-align:" + getTextVerticalAlign(node, type, slideMasterTextStyles) +
                                ";";
                                */
                            }
                            
                        }
                        rowsStyl += " background-color:#" + fillColor +";" + 
                                    " opacity:" + colorOpacity + ";";

                    }else if(i>0 && bandRowAttr!== undefined){
                        var fillColor="fff";
                        var colorOpacity = 1;
                        if((i%2)==0){
                            if(thisTblStyle["a:band2H"] !==undefined){
                                //console.log(thisTblStyle["a:band2H"]);
                                var bgFillschemeClr = getTextByPathList(thisTblStyle,["a:band2H","a:tcStyle","a:fill","a:solidFill"]);
                                if(bgFillschemeClr !==undefined){
                                    fillColor = getSolidFill(bgFillschemeClr);
                                    colorOpacity = getColorOpacity(bgFillschemeClr);
                                }
                                //borders color
                                //borders Width
                                var borderStyl = getTextByPathList(thisTblStyle,["a:band2H","a:tcStyle","a:tcBdr"]);
                                if(borderStyl !== undefined){
                                    var row_borders = getTableBorders(borderStyl);
                                    rowsStyl += row_borders;
                                }
                                //console.log(thisTblStyle["a:band2H"])
                                
                                //Text Style - TODO
                                var rowTxtStyl = getTextByPathList(thisTblStyle,["a:band2H","a:tcTxStyle"]);
                                if(rowTxtStyl !== undefined){
                                    
                                }
                                //console.log(i,thisTblStyle)
                            }/*else{
                                var bgFillschemeClr = thisTblStyle["a:wholeTbl"]["a:tcStyle"]["a:fill"]["a:solidFill"];
                                if(bgFillschemeClr !==undefined){
                                    fillColor = getSolidFill(bgFillschemeClr);
                                    colorOpacity = getColorOpacity(bgFillschemeClr);
                                }
                                //borders color
                                //borders Width
                                var borderStyl = thisTblStyle["a:wholeTbl"]["a:tcStyle"]["a:tcBdr"];
                                if(borderStyl !== undefined){
                                    var row_borders = getTableBorders(borderStyl);
                                    rowsStyl += row_borders;
                                }
                                //console.log(thisTblStyle["a:wholeTbl"])
                                
                                //Text Style - TODO
                                var rowTxtStyl = thisTblStyle["a:wholeTbl"]["a:tcTxStyle"];
                                if(rowTxtStyl !== undefined){
                                    
                                }                        
                            }*/
                        }else{
                            if(thisTblStyle["a:band1H"] !==undefined){
                                var bgFillschemeClr = getTextByPathList(thisTblStyle,["a:band1H","a:tcStyle","a:fill","a:solidFill"]);
                                if(bgFillschemeClr !==undefined){
                                    fillColor = getSolidFill(bgFillschemeClr);
                                    colorOpacity = getColorOpacity(bgFillschemeClr);
                                }
                                //borders color
                                //borders Width
                                var borderStyl = getTextByPathList(thisTblStyle,["a:band1H","a:tcStyle","a:tcBdr"]);
                                if(borderStyl !== undefined){
                                    var row_borders = getTableBorders(borderStyl);
                                    rowsStyl += row_borders;
                                }
                                //console.log(thisTblStyle["a:band1H"])
                                
                                //Text Style - TODO
                                var rowTxtStyl = getTextByPathList(thisTblStyle,["a:band1H","a:tcTxStyle"]);
                                if(rowTxtStyl !== undefined){
                                    
                                }
                            }
                        }
                        rowsStyl += " background-color:#" + fillColor +";" + 
                                    " opacity:" + colorOpacity + ";";
                    }
                    tableHtml += "<tr style='"+rowsStyl+"'>";
                    ////////////////////////////////////////////////
                
                    var tcNodes = trNodes[i]["a:tc"];
                    
                    if (tcNodes.constructor === Array) {
                        for (var j=0; j<tcNodes.length; j++) {
                            var text = genTextBody(tcNodes[j]["a:txBody"], node, undefined, undefined, undefined, warpObj);        
                            var rowSpan = getTextByPathList(tcNodes[j], ["attrs", "rowSpan"]);
                            var colSpan = getTextByPathList(tcNodes[j], ["attrs", "gridSpan"]);
                            var vMerge = getTextByPathList(tcNodes[j], ["attrs", "vMerge"]);
                            var hMerge = getTextByPathList(tcNodes[j], ["attrs", "hMerge"]);
                            //Cells Style : TODO /////////////Amir
                            //console.log(tcNodes[j]);
                            //if(j==0 && ())
                            var colWidthParam = getColsGrid[j]["attrs"]["w"];
                            var colStyl = "";
                            if(colWidthParam !== undefined){
                                var colWidth =  parseInt(colWidthParam) * 96 / 914400;
                                colStyl += "width:" + colWidth +"px;"
                            }
                            var getFill = tcNodes[j]["a:tcPr"]["a:solidFill"];
                            var fillColor = "";
                            var colorOpacity=1;
                            if(getFill !== undefined){
                                //console.log(getFill);
                                fillColor = getSolidFill(getFill);
                                colorOpacity = getColorOpacity(getFill);
                            }else{
                                //get from Theme (tableStyles.xml) TODO 
                                //get tableStyleId = a:tbl => a:tblPr => a:tableStyleId
                                var tbleStyleId = getTblPr["a:tableStyleId"];
                                if(tbleStyleId !== undefined){
                                    //get Style from tableStyles.xml by {var tbleStyleId}
                                    //table style object : tableStyles
                                    var tbleStylList = tableStyles["a:tblStyleLst"]["a:tblStyle"];
                                    
                                    for(var k=0;k<tbleStylList.length;k++){
                                        if(tbleStylList[k]["attrs"]["styleId"] == tbleStyleId){
                                            //console.log(tbleStylList[k]);
                                        }
                                    }
                                }
                                //console.log(tbleStyleId);
                            }
                            if(fillColor != ""){
                                colStyl += " background-color:#" + fillColor +";";
                                colStyl += " opacity" + colorOpacity +";";
                            }
                            //console.log(fillColor);
                            ////////////////////////////////////
                        
                            
                            if (rowSpan !== undefined) {
                                tableHtml += "<td rowspan='" + parseInt(rowSpan) + "' style='"+colStyl+"'>" + text + "</td>";
                            } else if (colSpan !== undefined) {
                                tableHtml += "<td colspan='" + parseInt(colSpan) + "' style='"+colStyl+"'>" + text + "</td>";
                            } else if (vMerge === undefined && hMerge === undefined) {
                                tableHtml += "<td style='"+colStyl+"'>" + text + "</td>";
                            }
                        }
                    } else {
                        var text = genTextBody(tcNodes["a:txBody"]);
                        //Cells Style : TODO /////////////Amir
                        var colWidthParam = getColsGrid[0]["attrs"]["w"];
                        var colStyl = "";
                        if(colWidthParam !== undefined){
                            var colWidth =  parseInt(colWidthParam) * 96 / 914400;
                            colStyl += "width:" + colWidth +"px;"
                        }
                        var getFill = tcNodes["a:tcPr"]["a:solidFill"];
                        var fillColor = "";
                        var colorOpacity = 1;
                        if(getFill !== undefined){
                            //console.log(getFill);   
                            fillColor = getSolidFill(getFill);
                            colorOpacity = getColorOpacity(getFill);
                        }else{
                            //get from Theme TODO
                        }
                        if(fillColor != ""){
                            colStyl += " background-color:#" + fillColor +";"
                            colStyl += " opacity" + colorOpacity +";";
                        }                
                        ////////////////////////////////////
                        tableHtml += "<td style='"+colStyl+"'>" + text + "</td>";
                    }
                    tableHtml += "</tr>";
                }
            } else {
                //////////////row height ////////////Amir
                var rowHeightParam = trNodes["attrs"]["h"];
                var rowHeight = 0;
                if(rowHeightParam !== undefined){
                    rowHeight = parseInt(rowHeightParam) * 96 / 914400;
                    tableHtml += "<tr style='height:"+rowHeight+"px;'>";
                }else{
                    tableHtml += "<tr>";
                }
                ////////////////////////////////////////////////
                var tcNodes = trNodes["a:tc"];
                if (tcNodes.constructor === Array) {
                    for (var j=0; j<tcNodes.length; j++) {
                        var text = genTextBody(tcNodes[j]["a:txBody"]);
                        //Cells Style : TODO /////////////Amir
                        var colWidthParam = getColsGrid[j]["attrs"]["w"];
                        var colStyl = "";
                        if(colWidthParam !== undefined){
                            var colWidth =  parseInt(colWidthParam) * 96 / 914400;
                            colStyl += "width:" + colWidth +"px;"
                        }
                        var getFill = tcNodes[j]["a:tcPr"]["a:solidFill"];
                        var fillColor = "";
                        var colorOpacity = 1;
                        if(getFill !== undefined){ 
                            fillColor = getSolidFill(getFill);
                            colorOpacity = getColorOpacity(getFill);
                        }else{
                            //get from Theme TODO
                            //get tableStyleId
                            // a:tbl => a:tblPr => a:tableStyleId
                        }
                        if(fillColor != ""){
                            colStyl += " background-color:#" + fillColor +";"
                            colStyl += " opacity" + colorOpacity +";";
                        }                
                        ////////////////////////////////////
                    tableHtml += "<td style='"+colStyl+"'>" + text + "</td>";
                    }
                } else {
                    var text = genTextBody(tcNodes["a:txBody"]);
                        //Cells Style : TODO /////////////Amir
                        var colWidthParam = getColsGrid[0]["attrs"]["w"];
                        var colStyl = "";
                        if(colWidthParam !== undefined){
                            var colWidth =  parseInt(colWidthParam) * 96 / 914400;
                            colStyl += "width:" + colWidth +"px;"
                        }
                        var getFill = tcNodes[j]["a:tcPr"]["a:solidFill"];
                        var fillColor = "";
                        var colorOpacity = 1;
                        if(getFill !== undefined){
                            //console.log(getFill);
                            fillColor = getSolidFill(getFill);
                            colorOpacity = getColorOpacity(getFill);
                        }else{
                            //get from Theme TODO
                        }
                        if(fillColor != ""){
                            colStyl += " background-color:#" + fillColor +";"
                            colStyl += " opacity" + colorOpacity +";";
                        }                
                        ////////////////////////////////////
                    tableHtml += "<td style='"+colStyl+"'>" + text + "</td>";
                }
                tableHtml += "</tr>";
            }
            
            return tableHtml;
        }

        function genChart(node, warpObj) {
            
            var order = node["attrs"]["order"];
            var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
            var result = "<div id='chart" + chartID + "' class='block content' style='" + 
                            getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) + 
                            " z-index: " + order + ";'></div>";
            
            var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
            var refName = warpObj["slideResObj"][rid]["target"];
            var content = readXmlFile(warpObj["zip"], refName);
            var plotArea = getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);
            
            var chartData = null;
            for (var key in plotArea) {
                switch (key) {
                    case "c:lineChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + chartID,
                                "chartType": "lineChart",
                                "chartData": extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:barChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + chartID,
                                "chartType": "barChart",
                                "chartData": extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:pieChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + chartID,
                                "chartType": "pieChart",
                                "chartData": extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:pie3DChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + chartID,
                                "chartType": "pie3DChart",
                                "chartData": extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:areaChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + chartID,
                                "chartType": "areaChart",
                                "chartData": extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:scatterChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + chartID,
                                "chartType": "scatterChart",
                                "chartData": extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:catAx":
                        break;
                    case "c:valAx":
                        break;
                    default:
                }
            }
            
            if (chartData !== null) {
                MsgQueue.push(chartData);
            }
            
            chartID++;
            return result;
        }

        function genDiagram(node, warpObj) {
            //console.log(warpObj)
            //readXmlFile(zip, sldFileName)
            /**files define the diagram:
             * 1-colors#.xml,
             * 2-data#.xml, 
             * 3-layout#.xml,
             * 4-quickStyle#.xml.
             * 5-drawing#.xml, which Microsoft added as an extension for persisting diagram layout information.
             */
            ///get colors#.xml, data#.xml , layout#.xml , quickStyle#.xml
            var order = node["attrs"]["order"];
            var zip = warpObj["zip"];
            var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
            var dgmRelIds = getTextByPathList(node, ["a:graphic","a:graphicData","dgm:relIds","attrs"]);
             //console.log(dgmRelIds)
            var dgmClrFileId = dgmRelIds["r:cs"];
            var dgmDataFileId = dgmRelIds["r:dm"];
            var dgmLayoutFileId = dgmRelIds["r:lo"];
            var dgmQuickStyleFileId = dgmRelIds["r:qs"];
            var dgmClrFileName = warpObj["slideResObj"][dgmClrFileId].target,
                dgmDataFileName = warpObj["slideResObj"][dgmDataFileId].target,
                dgmLayoutFileName = warpObj["slideResObj"][dgmLayoutFileId].target;
                dgmQuickStyleFileName = warpObj["slideResObj"][dgmQuickStyleFileId].target;
            //console.log(dgmClrFileName,"\n",dgmDataFileName,"\n",dgmLayoutFileName,"\n",dgmQuickStyleFileName);
            var dgmClr = readXmlFile(zip, dgmClrFileName);
            var dgmData = readXmlFile(zip, dgmDataFileName);
            var dgmLayout = readXmlFile(zip, dgmLayoutFileName);
            var dgmQuickStyle = readXmlFile(zip, dgmQuickStyleFileName);
            //console.log(dgmClr,dgmData,dgmLayout,dgmQuickStyle)
             ///get drawing#.xml
             var dgmDrwFileName = "";
             var dataModelExt = getTextByPathList(dgmData, ["dgm:dataModel","dgm:extLst","a:ext","dsp:dataModelExt","attrs"]);
            if(dataModelExt !== undefined){
                var dgmDrwFileId = dataModelExt["relId"];
                dgmDrwFileName =  warpObj["slideResObj"][dgmDrwFileId]["target"];
            }
            //console.log("dgmDrwFileName: ",dgmDrwFileName);
            var dgmDrwFile = "";
            if(dgmDrwFileName != ""){
                dgmDrwFile =  readXmlFile(zip, dgmDrwFileName);
            }
            //console.log("dgmDrwFile: ",dgmDrwFile);
            //processSpNode(node, warpObj)
            var dgmDrwSpArray = getTextByPathList(dgmDrwFile,["dsp:drawing","dsp:spTree","dsp:sp"]);
            var rslt="";
            if(dgmDrwSpArray !== undefined){
                var dgmDrwSpArrayLen = dgmDrwSpArray.length;
                for(var i=0;i<dgmDrwSpArrayLen;i++){
                    var dspSp = dgmDrwSpArray[i];
                    var dspSpObjToStr = JSON.stringify(dspSp);
                    var pSpStr = dspSpObjToStr.replace(/dsp:/g,"p:");
                    var pSpStrToObj = JSON.parse(pSpStr);
                    //console.log("pSpStrToObj["+i+"]: ",pSpStrToObj);
                    rslt += processSpNode(pSpStrToObj, warpObj)
                    //console.log("rslt["+i+"]: ",rslt);
                }
                // dgmDrwFile: "dsp:"-> "p:"
            }

            return "<div class='block content' style='" + 
                        getPosition(xfrmNode, undefined, undefined) + 
                        getSize(xfrmNode, undefined, undefined) + 
                    "'>"+rslt+"</div>";
        }

        function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
            
            var off = undefined;
            var x = -1, y = -1;
            
            if (slideSpNode !== undefined) {
                off = slideSpNode["a:off"]["attrs"];
            } else if (slideLayoutSpNode !== undefined) {
                off = slideLayoutSpNode["a:off"]["attrs"];
            } else if (slideMasterSpNode !== undefined) {
                off = slideMasterSpNode["a:off"]["attrs"];
            }
            
            if (off === undefined) {
                return "";
            } else {
                x = parseInt(off["x"]) * 96 / 914400;
                y = parseInt(off["y"]) * 96 / 914400;
                return (isNaN(x) || isNaN(y)) ? "" : "top:" + y + "px; left:" + x + "px;";
            }
            
        }

        function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
            
            var ext = undefined;
            var w = -1, h = -1;
            
            if (slideSpNode !== undefined) {
                ext = slideSpNode["a:ext"]["attrs"];
            } else if (slideLayoutSpNode !== undefined) {
                ext = slideLayoutSpNode["a:ext"]["attrs"];
            } else if (slideMasterSpNode !== undefined) {
                ext = slideMasterSpNode["a:ext"]["attrs"];
            }
            
            if (ext === undefined) {
                return "";
            } else {
                w = parseInt(ext["cx"]) * 96 / 914400;
                h = parseInt(ext["cy"]) * 96 / 914400;
                return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
            }    
            
        }

        function getHorizontalAlign(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {

            var algn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            if (algn === undefined) {
                algn = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                if (algn === undefined) {
                    algn = getTextByPathList(slideMasterSpNode, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                    if (algn === undefined) {
                        switch (type) {
                            case "title":
                            case "subTitle":
                            case "ctrTitle":
                                algn = getTextByPathList(slideMasterTextStyles, ["p:titleStyle", "a:lvl1pPr", "attrs", "alng"]);
                                break;
                            default:
                                algn = getTextByPathList(slideMasterTextStyles, ["p:otherStyle", "a:lvl1pPr", "attrs", "alng"]);
                        }
                    }
                }
            }
            // TODO:
            if (algn === undefined) {
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    return "h-mid";
                } else if (type == "sldNum") {
                    return "h-right";
                }
            }
            return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
        }

        function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
            
            // 上中下對齊: X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
            var anchor = getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
            if (anchor === undefined) {
                anchor = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                if (anchor === undefined) {
                    anchor = getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                }
            }
            
            return anchor === "ctr" ? "v-mid" : anchor === "b" ?  "v-down" : "v-up";
        }

        function getFontType(node, type, slideMasterTextStyles) {
            var typeface = getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);
            
            if (typeface === undefined) {
                var fontSchemeNode = getTextByPathList(themeContent, ["a:theme", "a:themeElements", "a:fontScheme"]);
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    typeface = getTextByPathList(fontSchemeNode, ["a:majorFont", "a:latin", "attrs", "typeface"]);
                } else if (type == "body") {
                    typeface = getTextByPathList(fontSchemeNode, ["a:minorFont", "a:latin", "attrs", "typeface"]);
                } else {
                    typeface = getTextByPathList(fontSchemeNode, ["a:minorFont", "a:latin", "attrs", "typeface"]);
                }
            }
            
            return (typeface === undefined) ? "inherit" : typeface;
        }

        function getFontColor(node, spNode, type, slideMasterTextStyles) {
            var solidFillNode = getTextByPathStr(node, "a:rPr a:solidFill");
            var color;
            if(solidFillNode === undefined){
                var sPstyle = getTextByPathList(spNode, ["p:style","a:fontRef"]);
                if(sPstyle !== undefined){
                    color = getSolidFill(sPstyle);
                }
            }else{
                color =   getSolidFill(solidFillNode);
            }
            
            //console.log(themeContent)
            //var schemeClr = getTextByPathList(buClrNode ,["a:schemeClr", "attrs","val"]);
            return (color === undefined || color === "FFF") ? "#000" : "#" + color;
        }
        function getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
            var fontSize = undefined;
            if (node["a:rPr"] !== undefined) {
                fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
            }
            
            if ((isNaN(fontSize) || fontSize === undefined)) {
                var sz = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:lstStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
            }
            
            if (isNaN(fontSize) || fontSize === undefined) {
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    var sz = getTextByPathList(slideMasterTextStyles, ["p:titleStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
                } else if (type == "body") {
                    var sz = getTextByPathList(slideMasterTextStyles, ["p:bodyStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
                } else if (type == "dt" || type == "sldNum") {
                    var sz = "1200";
                } else if (type === undefined) {
                    var sz = getTextByPathList(slideMasterTextStyles, ["p:otherStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
                }
                fontSize = parseInt(sz) / 100;
            }
            
            var baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
            if (baseline !== undefined && !isNaN(fontSize)) {
                fontSize -= 10;
            }
            
            return isNaN(fontSize) ? "inherit" : (fontSize + "pt");
        }

        function getFontBold(node, type, slideMasterTextStyles) {
            return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "initial";
        }

        function getFontItalic(node, type, slideMasterTextStyles) {
            return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "normal";
        }

        function getFontDecoration(node, type, slideMasterTextStyles) {
            ///////////////////////////////Amir///////////////////////////////
            if(node["a:rPr"] !== undefined){
                var underLine = node["a:rPr"]["attrs"]["u"] !== undefined? node["a:rPr"]["attrs"]["u"]:"none";
                var strikethrough = node["a:rPr"]["attrs"]["strike"] !== undefined?  node["a:rPr"]["attrs"]["strike"]:'noStrike';
                //console.log("strikethrough: "+strikethrough);
                
                if(underLine != "none" && strikethrough == "noStrike"){
                    return "underline";
                }else if(underLine == "none" && strikethrough != "noStrike"){
                    return "line-through";
                }else if(underLine != "none" && strikethrough != "noStrike"){
                    return "underline line-through";
                }else{
                    return "initial";
                }
            }else{
                return "initial";
            }
            /////////////////////////////////////////////////////////////////
            //return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "initial";
        }
        ////////////////////////////////////Amir/////////////////////////////////////
        function getTextHorizontalAlign(node, type, slideMasterTextStyles){
            var getAlgn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            var align = "initial";
            if(getAlgn !== undefined){
                switch(getAlgn){
                    case "l":
                        align = "left";
                        break;
                    case "r":
                        align = "right";
                        break;
                    case "ctr":
                        align = "center";
                        break;
                    case "just":
                        align = "justify";
                        break;
                    case "dist":
                        align = "justify";
                        break;
                    default:
                        align = "initial";
                }
            }
            return align;
        }
        /////////////////////////////////////////////////////////////////////
        function getTextVerticalAlign(node, type, slideMasterTextStyles) {
            var baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
            return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
        }
        ///////////////////////////////////Amir/////////////////////////////
        function getTextDirection(node, type, slideMasterTextStyles){
            //get lvl
        var pprLvl = getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
            var pprLvlNum = pprLvl===undefined?1:Number(pprLvl)+1;
            var lvlNode = "a:lvl"+pprLvlNum+"pPr";
            var pprAlgn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            var isDir = getTextByPathList(slideMasterTextStyles, ["p:bodyStyle",lvlNode, "attrs", "rtl"]);
            //var tmp = getTextByPathList(node, ["a:r", "a:t"]);
            var dir = "";
            if (isDir !== undefined){
                if(isDir=="1" && (pprAlgn ===undefined || pprAlgn =="r")){
                    dir = "rtl";
                }else{ //isDir =="0"
                    dir = "ltr";
                }
            }
            //console.log(tmp,isDir,pprAlgn,dir)
            return dir;
        }
        function getTableBorders(node){
            var borderStyle = "";
            if(node["a:bottom"] !== undefined){
                var obj = {
                    "p:spPr":{
                        "a:ln":node["a:bottom"]["a:ln"]
                    }
                }
                var borders = getBorder(obj, false);
                borderStyle += borders.replace("border","border-bottom");
            }
            if(node["a:top"] !== undefined){
                var obj = {
                    "p:spPr":{
                        "a:ln":node["a:top"]["a:ln"]
                    }
                }
                var borders = getBorder(obj, false);
                borderStyle += borders.replace("border","border-top");
            }
            if(node["a:right"] !== undefined){
                var obj = {
                    "p:spPr":{
                        "a:ln":node["a:right"]["a:ln"]
                    }
                }
                var borders = getBorder(obj, false);
                borderStyle += borders.replace("border","border-right");
            }
            if(node["a:left"] !== undefined){
                var obj = {
                    "p:spPr":{
                        "a:ln":node["a:left"]["a:ln"]
                    }
                }
                var borders = getBorder(obj, false);
                borderStyle += borders.replace("border","border-left");
            }

            return borderStyle;
        }
        //////////////////////////////////////////////////////////////////
        function getBorder(node, isSvgMode) {
            
            var cssText = "border: ";
            
            // 1. presentationML
            var lineNode = node["p:spPr"]["a:ln"];
            
            // Border width: 1pt = 12700, default = 0.75pt
            var borderWidth = parseInt(getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
            if (isNaN(borderWidth) || borderWidth < 1) {
                cssText += "1pt ";
            } else {
                cssText += borderWidth + "pt ";
            }
            // Border type
            var borderType = getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
            var strokeDasharray = "0";
            switch (borderType) {
                case "solid":
                    cssText += "solid";
                    strokeDasharray = "0";
                    break;
                case "dash":
                    cssText += "dashed";
                    strokeDasharray = "5";
                    break;
                case "dashDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 5, 1, 5";
                    break;
                case "dot":
                    cssText += "dotted";
                    strokeDasharray = "1, 5";
                    break;
                case "lgDash":
                    cssText += "dashed";
                    strokeDasharray = "10, 5";
                    break;
                case "lgDashDotDot":
                    cssText += "dashed";
                    strokeDasharray = "10, 5, 1, 5, 1, 5";
                    break;
                case "sysDash":
                    cssText += "dashed";
                    strokeDasharray = "5, 2";
                    break;
                case "sysDashDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 2, 1, 5";
                    break;
                case "sysDashDotDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 2, 1, 5, 1, 5";
                    break;
                case "sysDot":
                    cssText += "dotted";
                    strokeDasharray = "2, 5";
                    break;
                case undefined:
                    //console.log(borderType);
                default:
                    cssText += "solid";
                    strokeDasharray = "0";
            }    
            // Border color
            var borderColor = getTextByPathList(lineNode, ["a:solidFill", "a:srgbClr", "attrs", "val"]);
            if (borderColor === undefined) {
                var schemeClrNode = getTextByPathList(lineNode, ["a:solidFill", "a:schemeClr"]);
                if(schemeClrNode !== undefined){
                    var schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);    
                    var borderColor = getSchemeColorFromTheme(schemeClr,undefined);
                }
            }
            
            // 2. drawingML namespace
            if (borderColor === undefined) {
                var schemeClrNode = getTextByPathList(node, ["p:style", "a:lnRef", "a:schemeClr"]);
                if(schemeClrNode !== undefined){
                    var schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);    
                    var borderColor = getSchemeColorFromTheme(schemeClr,undefined);
                }
                
                if (borderColor !== undefined) {
                    var shade = getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
                    if (shade !== undefined) {
                        shade = parseInt(shade) / 100000;
                        var color = new colz.Color("#" + borderColor);
                        color.setLum(color.hsl.l * shade);
                        borderColor = color.hex.replace("#", "");
                    }
                }
                
            }
            
            if (borderColor === undefined) {
                if (isSvgMode) {
                    borderColor = "none";
                } else {
                    borderColor = "#000";
                }
            } else {
                borderColor = "#" + borderColor;
                
            }
            cssText += " " + borderColor + " ";
            

            
            if (isSvgMode) {
                return {"color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray};
            } else {
                return cssText + ";";
            }
        }

        function getSlideBackgroundFill(slideContent, slideLayoutContent, slideMasterContent,warpObj) {
            //console.log(slideContent)
            //getFillType(node)
            var bgPr = getTextByPathList(slideContent, ["p:sld", "p:cSld","p:bg","p:bgPr"]);
            var bgRef = getTextByPathList(slideContent, ["p:sld", "p:cSld","p:bg","p:bgRef"]);
            var bgcolor;
            
            if(bgPr !== undefined){
                //bgcolor = "background-color: blue;";
                var bgFillTyp =  getFillType(bgPr);

                if(bgFillTyp == "SOLID_FILL"){
                    var sldFill = bgPr["a:solidFill"];
                    var bgColor = getSolidFill(sldFill);
                    var sldTint =  getColorOpacity(sldFill);
                    bgcolor =  "background: rgba("+ hexToRgbNew(bgColor)+","+ sldTint+");";
                    
                }else if(bgFillTyp == "GRADIENT_FILL"){
                    bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent);
                }else if(bgFillTyp == "PIC_FILL"){
                    bgcolor = getBgPicFill(bgPr, "slideBg", warpObj);

                }
            //console.log(slideContent,slideMasterContent,color_ary,tint_ary,rot,bgcolor)
            }else if(bgRef !== undefined){
                //console.log("slideContent",bgRef)
                var phClr;
                if (bgRef["a:srgbClr"] !== undefined) {
                    phClr = getTextByPathList(bgRef,["a:srgbClr","attrs", "val"]); //#...
                }else if(bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                    var schemeClr = getTextByPathList(bgRef,["a:schemeClr","attrs", "val"]);
                    phClr = getSchemeColorFromTheme("a:" + schemeClr,slideMasterContent); //#...
                    //console.log("schemeClr",schemeClr,"phClr=",phClr)
                }
                var idx = Number(bgRef["attrs"]["idx"]);
            

                if(idx == 0 || idx==1000){
                    //no background
                }else if(idx > 0 && idx < 1000){
                    //fillStyleLst in themeContent
                    //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                        //bgcolor = "background: red;";
                }else if(idx > 1000 ){
                    //bgFillStyleLst  in themeContent
                    //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                    var trueIdx = idx - 1000;
                    var bgFillLst = themeContent["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    var sortblAry = [];
                    Object.keys(bgFillLst).forEach(function(key) {
                        var bgFillLstTyp = bgFillLst[key];
                        if(key != "attrs"){
                            if(bgFillLstTyp.constructor === Array ){
                                for(var i=0;i<bgFillLstTyp.length;i++){
                                    var obj = {};
                                    obj[key] = bgFillLstTyp[i];
                                    obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                    sortblAry.push(obj)
                                }
                            }else{
                                var obj = {};
                                obj[key] = bgFillLstTyp;
                                obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                sortblAry.push(obj)
                            }
                        }
                    });
                    var sortByOrder = sortblAry.slice(0);
                    sortByOrder.sort(function(a,b) {
                        return a.idex - b.idex;
                    });
                    var bgFillLstIdx = sortByOrder[trueIdx-1];
                    var bgFillTyp =  getFillType(bgFillLstIdx);
                    if(bgFillTyp == "SOLID_FILL"){
                        var sldFill = bgFillLstIdx["a:solidFill"];
                        //var sldBgColor = getSolidFill(sldFill);
                        var sldTint =  getColorOpacity(sldFill);
                            bgcolor =  "background: rgba("+ hexToRgbNew(phClr)+","+ sldTint+");";
                            //console.log("slideMasterContent - sldFill",sldFill)
                    }else if(bgFillTyp == "GRADIENT_FILL"){
                        bgcolor = getBgGradientFill(bgPr, phClr, slideMasterContent);
                    }
                }
                
            }else{
                bgPr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld","p:bg","p:bgPr"]);
                bgRef = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld","p:bg","p:bgRef"]);
                //console.log("slideLayoutContent",bgPr,bgRef)
                if(bgPr !== undefined){
                    var bgFillTyp =  getFillType(bgPr);
                    if(bgFillTyp == "SOLID_FILL"){
                        var sldFill = bgPr["a:solidFill"];
                        var bgColor = getSolidFill(sldFill);
                        var sldTint =  getColorOpacity(sldFill);
                        bgcolor =  "background: rgba("+ hexToRgbNew(bgColor)+","+ sldTint+");";
                    }else if(bgFillTyp == "GRADIENT_FILL"){
                        bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent);
                    }else if(bgFillTyp == "PIC_FILL"){
                        bgcolor = getBgPicFill(bgPr, "layoutBg", warpObj);

                    }
                    //console.log("slideLayoutContent",bgcolor)
                }else if(bgRef !== undefined){
                    bgcolor = "background: red;";
                }else{
                    bgPr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld","p:bg","p:bgPr"]);
                    bgRef = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld","p:bg","p:bgRef"]);
                    
                    //console.log("bgRef",bgRef["a:schemeClr"]["attrs"]["val"])
                    if(bgPr !== undefined){
                        var bgFillTyp =  getFillType(bgPr);
                        if(bgFillTyp == "SOLID_FILL"){
                            var sldFill = bgPr["a:solidFill"];
                            var bgColor = getSolidFill(sldFill);
                            var sldTint =  getColorOpacity(sldFill);
                            bgcolor =  "background: rgba("+ hexToRgbNew(bgColor)+","+ sldTint+");";
                        }else if(bgFillTyp == "GRADIENT_FILL"){
                            bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent);
                        }else if(bgFillTyp == "PIC_FILL"){
                            bgcolor = getBgPicFill(bgPr ,"masterBg" , warpObj);
                        }
                    }else if(bgRef !== undefined){
                        //var obj={
                        //    "a:solidFill": bgRef
                        //}
                        //var phClr = getSolidFill(bgRef);
                        var phClr;
                        if (bgRef["a:srgbClr"] !== undefined) {
                            phClr = getTextByPathList(bgRef,["a:srgbClr","attrs", "val"]); //#...
                        }else if(bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                            var schemeClr = getTextByPathList(bgRef,["a:schemeClr","attrs", "val"]);
                            
                            phClr = getSchemeColorFromTheme("a:" + schemeClr,slideMasterContent); //#...
                            //console.log("phClr",phClr)
                        }
                        var idx = Number(bgRef["attrs"]["idx"]);
                        //console.log("phClr=",phClr,"idx=",idx)

                        if(idx == 0 || idx==1000){
                            //no background
                        }else if(idx > 0 && idx < 1000){
                            //fillStyleLst in themeContent
                            //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                            //bgcolor = "background: red;";
                        }else if(idx > 1000 ){
                            //bgFillStyleLst  in themeContent
                            //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                            var trueIdx = idx - 1000;
                            var bgFillLst = themeContent["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                            var sortblAry = [];
                            Object.keys(bgFillLst).forEach(function(key) {
                                //console.log("cubicBezTo["+key+"]:");
                                var bgFillLstTyp = bgFillLst[key];
                                if(key != "attrs"){
                                    if(bgFillLstTyp.constructor === Array ){
                                        for(var i=0;i<bgFillLstTyp.length;i++){
                                            var obj = {};
                                            obj[key] = bgFillLstTyp[i];
                                            obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                            sortblAry.push(obj)
                                        }
                                    }else{
                                        var obj = {};
                                        obj[key] = bgFillLstTyp;
                                        obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                        sortblAry.push(obj)
                                    }
                                }
                            });
                            var sortByOrder = sortblAry.slice(0);
                            sortByOrder.sort(function(a,b) {
                                return a.idex - b.idex;
                            });
                            var bgFillLstIdx = sortByOrder[trueIdx-1];
                            var bgFillTyp =  getFillType(bgFillLstIdx);
                            //console.log("bgFillLstIdx",bgFillLstIdx);
                            if(bgFillTyp == "SOLID_FILL"){
                                var sldFill = bgFillLstIdx["a:solidFill"];
                                var sldTint =  getColorOpacity(sldFill);
                                bgcolor =  "background: rgba("+ hexToRgbNew(phClr)+","+ sldTint+");";
                            }else if(bgFillTyp == "GRADIENT_FILL"){
                                bgcolor = getBgGradientFill(bgPr, phClr, slideMasterContent);
                            }else{
                                console.log(bgFillTyp)
                            }
                        }
                    }
                }
            }
            
            //console.log("bgcolor: ",bgcolor)   
            return bgcolor;
        }
        function getBgGradientFill(bgPr, phClr, slideMasterContent){
            var bgcolor;
            var grdFill = bgPr["a:gradFill"];
            var gsLst = grdFill["a:gsLst"]["a:gs"]; 
            var startColorNode , endColorNode;
            var color_ary = [];
            var tint_ary = [];
            for(var i=0;i<gsLst.length;i++){
                var lo_tint;
                var lo_color = "";
                if (gsLst[i]["a:srgbClr"] !== undefined) {
                    if(phClr === undefined){
                        lo_color = getTextByPathList(gsLst[i],["a:srgbClr","attrs", "val"]); //#...
                    }
                    lo_tint = getTextByPathList(gsLst[i],["a:srgbClr","a:tint","attrs","val"]);
                }else if(gsLst[i]["a:schemeClr"] !== undefined) { //a:schemeClr
                    if(phClr === undefined){
                        var schemeClr = getTextByPathList(gsLst[i],["a:schemeClr","attrs", "val"]);
                        lo_color = getSchemeColorFromTheme("a:" + schemeClr,slideMasterContent); //#...
                    }
                    lo_tint = getTextByPathList(gsLst[i],["a:schemeClr","a:tint","attrs","val"]);
                    //console.log("schemeClr",schemeClr,slideMasterContent)
                }
                //console.log("lo_color",lo_color)
                color_ary[i] =  lo_color;
                tint_ary[i] = (lo_tint !==undefined)?parseInt(lo_tint) / 100000:1;
            } 
            //get rot
            var lin = grdFill["a:lin"];
            var rot = 90;
            if(lin !== undefined){
                rot = angleToDegrees(lin["attrs"]["ang"]) + 90;
            } 
            bgcolor =  "background: linear-gradient("+rot+"deg,";
            for(var i=0;i<gsLst.length;i++){
                if(i==gsLst.length-1){
                    if(phClr === undefined){
                        bgcolor += "rgba("+ hexToRgbNew(color_ary[i])+","+ tint_ary[i]+")"+");";
                    }else{
                        bgcolor += "rgba("+ hexToRgbNew(phClr)+","+ tint_ary[i]+")"+");";
                    }
                }else{
                    if(phClr === undefined){
                        bgcolor += "rgba("+ hexToRgbNew(color_ary[i])+","+ tint_ary[i]+")"+", ";
                    }else{
                        bgcolor += "rgba("+ hexToRgbNew(phClr)+","+ tint_ary[i]+")"+", ";
                    }
                }
                    
            }   
            return bgcolor;
        }
        function getBgPicFill(bgPr, sorce, warpObj){
            var bgcolor;
            var picFillBase64 = getPicFill(sorce, bgPr["a:blipFill"], warpObj);
            var ordr = bgPr["attrs"]["order"];
            //a:srcRect
            //a:stretch => a:fillRect =>attrs (l:-17000, r:-17000)
            bgcolor = "background-image: url(" + picFillBase64 + ");  z-index: " + ordr + ";";
            return bgcolor;
        }
        function hexToRgbNew(hex) {
        var arrBuff = new ArrayBuffer(4);
        var vw = new DataView(arrBuff);
        vw.setUint32(0,parseInt(hex, 16),false);
        var arrByte = new Uint8Array(arrBuff);

        return arrByte[1] + "," + arrByte[2] + "," + arrByte[3];
        }
        function getShapeFill(node, isSvgMode, warpObj) {
            
            // 1. presentationML
            // p:spPr [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
            // From slide
            //Fill Type:
            //console.log("ShapeFill: ", node)
            var fillType = getFillType(getTextByPathList(node, ["p:spPr"]));
            var fillColor;
            if (fillType == "NO_FILL") {
                return isSvgMode ? "none" : "background-color: initial;";
            }else if(fillType == "SOLID_FILL"){
                var shpFill = node["p:spPr"]["a:solidFill"];
                fillColor = getSolidFill(shpFill);
            }else if(fillType == "GRADIENT_FILL"){
                var shpFill = node["p:spPr"]["a:gradFill"];
                // fillColor = getSolidFill(shpFill);
                fillColor = getGradientFill(shpFill);
                //console.log("shpFill",shpFill,grndColor.color)
            }else if(fillType == "PATTERN_FILL"){
                var shpFill = node["p:spPr"]["a:pattFill"];
                fillColor = getPatternFill(shpFill);
            }else if(fillType == "PIC_FILL"){
                var shpFill = node["p:spPr"]["a:blipFill"];
                fillColor = getPicFill("slideBg",shpFill, warpObj);
            }
            

            // 2. drawingML namespace
            if (fillColor === undefined) {
                var clrName = getTextByPathList(node, ["p:style", "a:fillRef"]);
                fillColor = getSolidFill(clrName);
            }

            if (fillColor !== undefined) {
                
                if(fillType == "GRADIENT_FILL"){

                    if (isSvgMode) {
                        // console.log("GRADIENT_FILL color", fillColor.color[0])
                        return fillColor;
                    } else {
                        var colorAry = fillColor.color;
                        var rot = fillColor.rot;
                        
                        var bgcolor =  "background: linear-gradient("+rot+"deg,";
                        for(var i=0;i<colorAry.length;i++){
                            if(i==colorAry.length-1){
                                bgcolor += colorAry[i]+");";
                            }else{
                                bgcolor += colorAry[i]+", ";
                            }
                                
                        } 
                        return bgcolor;
                    }
                }else if(fillType == "PIC_FILL"){
                    if (isSvgMode) {
                        return fillColor;
                    } else {

                        return "background-image:url(" + fillColor + ");";
                    }            
                }else{
                    if (isSvgMode) {
                        var color = new colz.Color(fillColor);
                        fillColor =  color.rgb.toString();
                        
                        return fillColor;
                    } else {
                        //console.log(node,"fillColor: ",fillColor,"fillType: ",fillType,"isSvgMode: ",isSvgMode)
                        return "background-color: #" + fillColor + ";";
                    }
                }
            } else {
                if (isSvgMode) {
                    return "none";
                } else {
                    return "background-color: initial;";
                }
                
            }
            
        }
        ///////////////////////Amir//////////////////////////////
        function getFillType(node){
            //Need to test/////////////////////////////////////////////
            //SOLID_FILL
            //PIC_FILL
            //GRADIENT_FILL
            //PATTERN_FILL
            //NO_FILL
            var fillType = "";
            if (node["a:noFill"] !== undefined) {
                fillType = "NO_FILL";
            }
            if (node["a:solidFill"] !== undefined) {
                fillType = "SOLID_FILL";
            }
            if (node["a:gradFill"] !== undefined) {
                fillType = "GRADIENT_FILL";
            }
            if (node["a:pattFill"] !== undefined) {
                fillType = "PATTERN_FILL";
            }
            if (node["a:blipFill"] !== undefined) {
                fillType = "PIC_FILL";
            }

            return fillType;
        }
        function getGradientFill(node){
            var gsLst = node["a:gsLst"]["a:gs"];
            //get start color
            var color_ary = [];
            var tint_ary = [];
            for(var i=0;i<gsLst.length;i++){
                var lo_tint;
                var lo_color = getSolidFill(gsLst[i]);
                if (gsLst[i]["a:srgbClr"] !== undefined) {
                    var lumMod = parseInt(getTextByPathList(node, ["a:srgbClr", "a:lumMod", "attrs", "val"])) / 100000;
                    var lumOff = parseInt(getTextByPathList(node, ["a:srgbClr", "a:lumOff", "attrs", "val"])) / 100000;
                    if (isNaN(lumMod)) {
                        lumMod = 1.0;
                    }
                    if (isNaN(lumOff)) {
                        lumOff = 0;
                    }
                    //console.log([lumMod, lumOff]);
                    lo_color = applyLumModify(lo_color, lumMod, lumOff);
                }else if(gsLst[i]["a:schemeClr"] !== undefined) { //a:schemeClr
                    var lumMod = parseInt(getTextByPathList(gsLst[i], ["a:schemeClr", "a:lumMod", "attrs", "val"])) / 100000;
                    var lumOff = parseInt(getTextByPathList(gsLst[i], ["a:schemeClr", "a:lumOff", "attrs", "val"])) / 100000;
                    if (isNaN(lumMod)) {
                        lumMod = 1.0;
                    }
                    if (isNaN(lumOff)) {
                        lumOff = 0;
                    }
                    //console.log([lumMod, lumOff]);
                    lo_color = applyLumModify(lo_color, lumMod, lumOff);
                }
                //console.log("lo_color",lo_color)
                color_ary[i] =  lo_color;
            } 
            //get rot
            var lin = node["a:lin"];
            var rot = 0;
            if(lin !== undefined){
                rot = angleToDegrees(lin["attrs"]["ang"]) + 90;
            }
            return {
                "color":color_ary,
                "rot": rot
            }
        }
        function getPicFill(type,node,warpObj){
            //Need to test/////////////////////////////////////////////
            //rId
            //TODO - Image Properties - Tile, Stretch, or Display Portion of Image
                //(http://officeopenxml.com/drwPic-tile.php)
            var img;
            var rId = node["a:blip"]["attrs"]["r:embed"];
            var imgPath;
            if(type=="slideBg"){
                imgPath =  getTextByPathList(warpObj,["slideResObj",rId,"target"]);
            }else if(type == "layoutBg"){
                imgPath =  getTextByPathList(warpObj,["layoutResObj",rId,"target"]);
            }else if(type=="masterBg"){
                imgPath =  getTextByPathList(warpObj,["masterResObj",rId,"target"]);
            }
            if(imgPath === undefined){
            return undefined;
            } 
            var imgExt = imgPath.split(".").pop();
            if(imgExt=="xml"){
                return undefined;
            }    
            var imgArrayBuffer = warpObj["zip"].file(imgPath).asArrayBuffer();
            var imgMimeType = getImageMimeType(imgExt);
            img = "data:" + imgMimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer);
            return img;
        }
        function getPatternFill(node){
            //Need to test/////////////////////////////////////////////
            var color = "";
            var bgClr = node["a:bgClr"];
            color = getSolidFill(bgClr);
            return color;
        }

        function getSolidFill(node) {
            
            if (node === undefined) {
                return undefined;
            }
            
            var color = "FFF";
            
            if (node["a:srgbClr"] !== undefined) {
                color = getTextByPathList(node,["a:srgbClr","attrs", "val"]); //#...
            }else if(node["a:schemeClr"] !== undefined) { //a:schemeClr
                var schemeClr = getTextByPathList(node,["a:schemeClr","attrs", "val"]);
                //console.log(schemeClr)
                color = getSchemeColorFromTheme("a:" + schemeClr,undefined); //#...
                
            }else if(node["a:scrgbClr"] !== undefined){
                //<a:scrgbClr r="50%" g="50%" b="50%"/>  //Need to test/////////////////////////////////////////////
                var defBultColorVals = node["a:scrgbClr"]["attrs"];
                var red = (defBultColorVals["r"].indexOf("%") != -1)?defBultColorVals["r"].split("%").shift():defBultColorVals["r"];
                var green = (defBultColorVals["g"].indexOf("%") != -1)?defBultColorVals["g"].split("%").shift():defBultColorVals["g"];
                var blue = (defBultColorVals["b"].indexOf("%") != -1)?defBultColorVals["b"].split("%").shift():defBultColorVals["b"];
                var scrgbClr = red + "," + green + "," + blue;
                color = toHex(255*(Number(red)/100)) + toHex(255*(Number(green)/100)) + toHex(255*(Number(blue)/100));
                    //console.log("scrgbClr: " + scrgbClr);

            }else if(node["a:prstClr"] !== undefined){
                //<a:prstClr val="black"/>  //Need to test/////////////////////////////////////////////
                var prstClr  = node["a:prstClr"]["attrs"]["val"];
                color = getColorName2Hex(prstClr);
                //console.log("prstClr: " + prstClr+" => hexClr: "+color);
            }else if(node["a:hslClr"] !== undefined){
                //<a:hslClr hue="14400000" sat="100%" lum="50%"/>  //Need to test/////////////////////////////////////////////
                    var defBultColorVals = node["a:hslClr"]["attrs"];
                    var hue = Number(defBultColorVals["hue"])/100000;
                    var sat = Number((defBultColorVals["sat"].indexOf("%") != -1)?defBultColorVals["sat"].split("%").shift():defBultColorVals["sat"])/100;
                    var lum = Number((defBultColorVals["lum"].indexOf("%") != -1)?defBultColorVals["lum"].split("%").shift():defBultColorVals["lum"])/100;
                    var hslClr = defBultColorVals["hue"] + "," + defBultColorVals["sat"] + "," + defBultColorVals["lum"];
                    var hsl2rgb = hslToRgb(hue, sat, lum);
                    color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
                    //defBultColor = cnvrtHslColor2Hex(hslClr); //TODO
                    // console.log("hslClr: " + hslClr);
            }else if(node["a:sysClr"] !== undefined){
                //<a:sysClr val="windowText" lastClr="000000"/>  //Need to test/////////////////////////////////////////////
                var sysClr = getTextByPathList(node,["a:sysClr","attrs","lastClr"]);
                if(sysClr !== undefined){
                    color = sysClr;
                }
            }
            return color;
        }
        function toHex(n) {
        var hex = n.toString(16);
        while (hex.length < 2) {hex = "0" + hex; }
        return hex;
        }
        function hslToRgb(hue, sat, light) {
        var t1, t2, r, g, b;
        hue = hue / 60;
        if ( light <= 0.5 ) {
            t2 = light * (sat + 1);
        } else {
            t2 = light + sat - (light * sat);
        }
        t1 = light * 2 - t2;
        r = hueToRgb(t1, t2, hue + 2) * 255;
        g = hueToRgb(t1, t2, hue) * 255;
        b = hueToRgb(t1, t2, hue - 2) * 255;
        return {r : r, g : g, b : b};
        }
        function hueToRgb(t1, t2, hue) {
        if (hue < 0) hue += 6;
        if (hue >= 6) hue -= 6;
        if (hue < 1) return (t2 - t1) * hue + t1;
        else if(hue < 3) return t2;
        else if(hue < 4) return (t2 - t1) * (4 - hue) + t1;
        else return t1;
        }
        function getColorName2Hex(name) {
            var hex;
            var colorName =  ['AliceBlue','AntiqueWhite','Aqua','Aquamarine','Azure','Beige','Bisque','Black','BlanchedAlmond','Blue','BlueViolet','Brown','BurlyWood','CadetBlue','Chartreuse','Chocolate','Coral','CornflowerBlue','Cornsilk','Crimson','Cyan','DarkBlue','DarkCyan','DarkGoldenRod','DarkGray','DarkGrey','DarkGreen','DarkKhaki','DarkMagenta','DarkOliveGreen','DarkOrange','DarkOrchid','DarkRed','DarkSalmon','DarkSeaGreen','DarkSlateBlue','DarkSlateGray','DarkSlateGrey','DarkTurquoise','DarkViolet','DeepPink','DeepSkyBlue','DimGray','DimGrey','DodgerBlue','FireBrick','FloralWhite','ForestGreen','Fuchsia','Gainsboro','GhostWhite','Gold','GoldenRod','Gray','Grey','Green','GreenYellow','HoneyDew','HotPink','IndianRed','Indigo','Ivory','Khaki','Lavender','LavenderBlush','LawnGreen','LemonChiffon','LightBlue','LightCoral','LightCyan','LightGoldenRodYellow','LightGray','LightGrey','LightGreen','LightPink','LightSalmon','LightSeaGreen','LightSkyBlue','LightSlateGray','LightSlateGrey','LightSteelBlue','LightYellow','Lime','LimeGreen','Linen','Magenta','Maroon','MediumAquaMarine','MediumBlue','MediumOrchid','MediumPurple','MediumSeaGreen','MediumSlateBlue','MediumSpringGreen','MediumTurquoise','MediumVioletRed','MidnightBlue','MintCream','MistyRose','Moccasin','NavajoWhite','Navy','OldLace','Olive','OliveDrab','Orange','OrangeRed','Orchid','PaleGoldenRod','PaleGreen','PaleTurquoise','PaleVioletRed','PapayaWhip','PeachPuff','Peru','Pink','Plum','PowderBlue','Purple','RebeccaPurple','Red','RosyBrown','RoyalBlue','SaddleBrown','Salmon','SandyBrown','SeaGreen','SeaShell','Sienna','Silver','SkyBlue','SlateBlue','SlateGray','SlateGrey','Snow','SpringGreen','SteelBlue','Tan','Teal','Thistle','Tomato','Turquoise','Violet','Wheat','White','WhiteSmoke','Yellow','YellowGreen'];
            var colorHex =  ['f0f8ff','faebd7','00ffff','7fffd4','f0ffff','f5f5dc','ffe4c4','000000','ffebcd','0000ff','8a2be2','a52a2a','deb887','5f9ea0','7fff00','d2691e','ff7f50','6495ed','fff8dc','dc143c','00ffff','00008b','008b8b','b8860b','a9a9a9','a9a9a9','006400','bdb76b','8b008b','556b2f','ff8c00','9932cc','8b0000','e9967a','8fbc8f','483d8b','2f4f4f','2f4f4f','00ced1','9400d3','ff1493','00bfff','696969','696969','1e90ff','b22222','fffaf0','228b22','ff00ff','dcdcdc','f8f8ff','ffd700','daa520','808080','808080','008000','adff2f','f0fff0','ff69b4','cd5c5c','4b0082','fffff0','f0e68c','e6e6fa','fff0f5','7cfc00','fffacd','add8e6','f08080','e0ffff','fafad2','d3d3d3','d3d3d3','90ee90','ffb6c1','ffa07a','20b2aa','87cefa','778899','778899','b0c4de','ffffe0','00ff00','32cd32','faf0e6','ff00ff','800000','66cdaa','0000cd','ba55d3','9370db','3cb371','7b68ee','00fa9a','48d1cc','c71585','191970','f5fffa','ffe4e1','ffe4b5','ffdead','000080','fdf5e6','808000','6b8e23','ffa500','ff4500','da70d6','eee8aa','98fb98','afeeee','db7093','ffefd5','ffdab9','cd853f','ffc0cb','dda0dd','b0e0e6','800080','663399','ff0000','bc8f8f','4169e1','8b4513','fa8072','f4a460','2e8b57','fff5ee','a0522d','c0c0c0','87ceeb','6a5acd','708090','708090','fffafa','00ff7f','4682b4','d2b48c','008080','d8bfd8','ff6347','40e0d0','ee82ee','f5deb3','ffffff','f5f5f5','ffff00','9acd32'];
            var findIndx = colorName.indexOf(name);
            if(findIndx != -1){
                hex = colorHex[findIndx];
            }
        return hex;
        }
        function getColorOpacity(solidFill){
            
            if (solidFill === undefined) {
                return undefined;
            }
            var opcity = 1;

            if (solidFill["a:srgbClr"] !== undefined) {
                var tint = getTextByPathList(solidFill,["a:srgbClr","a:tint","attrs", "val"]);
                if(tint !== undefined){
                    opcity =  parseInt(tint) / 100000;
                }
            } else if (solidFill["a:schemeClr"] !== undefined) {
                var tint = getTextByPathList(solidFill,["a:schemeClr","a:tint","attrs", "val"]);
                if(tint !== undefined){
                    opcity =  parseInt(tint) / 100000;
                }
            }else if(solidFill["a:scrgbClr"] !== undefined){
                var tint = getTextByPathList(solidFill,["a:scrgbClr","a:tint","attrs", "val"]);
                if(tint !== undefined){
                    opcity =  parseInt(tint) / 100000;
                }

            }else if(solidFill["a:prstClr"] !== undefined){
                var tint = getTextByPathList(solidFill,["a:prstClr","a:tint","attrs", "val"]);
                if(tint !== undefined){
                    opcity =  parseInt(tint) / 100000;
                }
            }else if(solidFill["a:hslClr"] !== undefined){
                var tint = getTextByPathList(solidFill,["a:hslClr","a:tint","attrs", "val"]);
                if(tint !== undefined){
                    opcity =  parseInt(tint) / 100000;
                }
            }else if(solidFill["a:sysClr"] !== undefined){
                var tint = getTextByPathList(solidFill,["a:sysClr","a:tint","attrs", "val"]);
                if(tint !== undefined){
                    opcity =  parseInt(tint) / 100000;
                }
            }

            return opcity;
        }
        function getSchemeColorFromTheme(schemeClr,sldMasterNode) {
            //<p:clrMap ...> in slide master
            // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1" slideLayoutClrOvride
            
            if(slideLayoutClrOvride == "" || slideLayoutClrOvride === undefined ){
                slideLayoutClrOvride = getTextByPathList(sldMasterNode,["p:sldMaster","p:clrMap","attrs"])
            }
            //console.log(slideLayoutClrOvride);
            var schmClrName =  schemeClr.substr(2);
            switch (schmClrName) {
                case "tx1":
                case "tx2":
                case "bg1":
                case "bg2":
                    schemeClr = "a:"+slideLayoutClrOvride[schmClrName];
                    //console.log(schmClrName+ "=> "+schemeClr);
                    break;
            }
            
            var refNode = getTextByPathList(themeContent, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
            var color = getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
            if (color === undefined) {
                color = getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
            }
            return color;
        }

        function extractChartData(serNode) {
            
            var dataMat = new Array();
            
            if (serNode === undefined) {
                return dataMat;
            }
            
            if (serNode["c:xVal"] !== undefined) {
                var dataRow = new Array();
                eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function(innerNode, index) {
                    dataRow.push(parseFloat(innerNode["c:v"]));
                    return "";
                });
                dataMat.push(dataRow);
                dataRow = new Array();
                eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function(innerNode, index) {
                    dataRow.push(parseFloat(innerNode["c:v"]));
                    return "";
                });
                dataMat.push(dataRow);
            } else {
                eachElement(serNode, function(innerNode, index) {
                    var dataRow = new Array();
                    var colName = getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

                    // Category (string or number)
                    var rowNames = {};
                    if (getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], function(innerNode, index) {
                            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                            return "";
                        });
                    } else if (getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], function(innerNode, index) {
                            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                            return "";
                        });
                    }
                    
                    // Value
                    if (getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], function(innerNode, index) {
                            dataRow.push({x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"])});
                            return "";
                        });
                    }
                    
                    dataMat.push({key: colName, values: dataRow, xlabels: rowNames});
                    return "";
                });
            }
            
            return dataMat;
        }

        // ===== Node functions =====
        /**
         * getTextByPathStr
         * @param {Object} node
         * @param {string} pathStr
         */
        function getTextByPathStr(node, pathStr) {
            return getTextByPathList(node, pathStr.trim().split(/\s+/));
        }

        /**
         * getTextByPathList
         * @param {Object} node
         * @param {string Array} path
         */
        function getTextByPathList(node, path) {

            if (path.constructor !== Array) {
                throw Error("Error of path type! path is not array.");
            }
            
            if (node === undefined) {
                return undefined;
            }
            
            var l = path.length;
            for (var i=0; i<l; i++) {
                node = node[path[i]];
                if (node === undefined) {
                    return undefined;
                }
            }
            
            return node;
        }

        /**
         * eachElement
         * @param {Object} node
         * @param {function} doFunction
         */
        function eachElement(node, doFunction) {
            if (node === undefined) {
                return;
            }
            var result = "";
            if (node.constructor === Array) {
                var l = node.length;
                for (var i=0; i<l; i++) {
                    result += doFunction(node[i], i);
                }
            } else {
                result += doFunction(node, 0);
            }
            return result;
        }

        // ===== Color functions =====
        /**
         * applyShade
         * @param {string} rgbStr
         * @param {number} shadeValue
         */
        function applyShade(rgbStr, shadeValue) {
            var color = new colz.Color(rgbStr);
            color.setLum(color.hsl.l * shadeValue);
            return color.rgb.toString();
        }

        /**
         * applyTint
         * @param {string} rgbStr
         * @param {number} tintValue
         */
        function applyTint(rgbStr, tintValue) {
            var color = new colz.Color(rgbStr);
            color.setLum(color.hsl.l * tintValue + (1 - tintValue));
            return color.rgb.toString();
        }

        /**
         * applyLumModify
         * @param {string} rgbStr
         * @param {number} factor
         * @param {number} offset
         */
        function applyLumModify(rgbStr, factor, offset) {
            var color = new colz.Color(rgbStr);
            //color.setLum(color.hsl.l * factor);
            color.setLum(color.hsl.l * (1 + offset));
            return color.rgb.toString();
        }

        ///////////////////////Amir////////////////
        function angleToDegrees(angle) {
            if (angle == "" || angle==null) {
                return 0;
            }
            return Math.round(angle / 60000);
        }
        function getImageMimeType(imgFileExt){
            var mimeType = "";
            //console.log(imgFileExt)
            switch (imgFileExt.toLowerCase()) {
                case "jpg":
                case "jpeg":
                    mimeType = "image/jpeg";
                    break;
                case "png":
                    mimeType = "image/png";
                    break;
                case "gif":
                    mimeType = "image/gif";
                    break;
                case "emf": // Not native support
                    mimeType = "image/x-emf";
                    break;
                case "wmf": // Not native support
                    mimeType = "image/x-wmf";
                    break;
                case "svg":
                    mimeType = "image/svg+xml";
                    break;
                default:
                    mimeType = "image/*";
            }
            return mimeType;
        }
        function getSvgGradient(w,h,angl,color_arry,shpId){
            var stopsArray = getMiddleStops(color_arry.length-2);
            
            var svgAngle = '',
            svgHeight = h,
            svgWidth = w,
            svg = '',
            xy_ary = SVGangle(angl, svgHeight,svgWidth),
            x1 = xy_ary[0], 
            y1 = xy_ary[1], 
            x2 = xy_ary[2], 
            y2 = xy_ary[3]; 

            var sal = stopsArray.length, 
            sr = sal < 20 ? 100 : 1000; 
            svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"'; 
            svgAngle = '<linearGradient id="linGrd_'+shpId+'"' + svgAngle + '>\n';
            svg += svgAngle;

            for (var i = 0; i < sal; i++) {
                svg += '<stop offset="' + Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr + '" stop-color="' + color_arry[i] + '"';
                svg += '/>\n'
            }

            svg += '</linearGradient>\n' + ''; 
            
            return svg   
        }
        function getMiddleStops(s) {
            var sArry = ['0%', '100%'];
            if (s == 0) { 
                return true 
            }else {
                var i = s;
                while (i--) {
                    var middleStop = 100 - ((100 / (s + 1)) * (i + 1)), // AM: Ex - For 3 middle stops, progression will be 25%, 50%, and 75%, plus 0% and 100% at the ends.
                    middleStopString = middleStop + "%";
                    sArry.splice(-1, 0, middleStopString);
                } // AM: add into stopsArray before 100%
            }
            return sArry
        }
        function SVGangle(deg,svgHeight,svgWidth) {
            var w = parseFloat(svgWidth), 
            h = parseFloat(svgHeight), 
            ang = parseFloat(deg),
            o = 2, 
            n = 2,
            wc = w / 2, 
            hc = h / 2,
            tx1 = 2, 
            ty1 = 2, 
            tx2 = 2, 
            ty2 = 2,
            k = (((ang % 360) + 360) % 360),
            j = (360 - k) * Math.PI / 180,
            i = Math.tan(j),
            l = hc - i * wc;
            
            if (k == 0) {
                tx1 = w, 
                ty1 = hc, 
                tx2 = 0, 
                ty2 = hc 
            }else if (k < 90) {
                n = w, 
                o = 0 
            }else if (k == 90) { 
                tx1 = wc, 
                ty1 = 0, 
                tx2 = wc, 
                ty2 = h 
            }else if (k < 180) { 
                n = 0,
                o = 0 
            }else if (k == 180) { 
                tx1 = 0, 
                ty1 = hc, 
                tx2 = w, 
                ty2 = hc 
            }else if (k < 270) {
                n = 0, 
                o = h 
            }else if (k == 270) { 
                tx1 = wc, 
                ty1 = h, 
                tx2 = wc,
                ty2 = 0 
            }else { 
                n = w, 
                o = h; 
            }
            // AM: I could not quite figure out what m, n, and o are supposed to represent from the original code on visualcsstools.com.
            var m = o + (n / i),
            tx1 = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1,
            ty1 = ty1 == 2 ? i * tx1 + l : ty1,
            tx2 = tx2 == 2 ? w - tx1 : tx2,
            ty2 = ty2 == 2 ? h - ty1 : ty2,
            x1 = Math.round(tx2 / w * 100 * 100) / 100, 
            y1 = Math.round(ty2 / h * 100 * 100) / 100,
            x2 = Math.round(tx1 / w * 100 * 100) / 100, 
            y2 = Math.round(ty1 / h * 100 * 100) / 100;
            return [x1,y1,x2,y2];
        }
        function getSvgImagePattern(fillColor,shpId){
            var ptrn =  '<pattern id="imgPtrn_'+shpId+'"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
            ptrn += '<image  xlink:href="'+fillColor+'" preserveAspectRatio="none" width="1" height="1"></image>';
            ptrn += '</pattern>';
            return ptrn;
        }
        
        function processMsgQueue(queue) {
            for (var i=0; i<queue.length; i++) {
                processSingleMsg(queue[i].data);
            }
        }

        function processSingleMsg(d) {
            
            var chartID = d.chartID;
            var chartType = d.chartType;
            var chartData = d.chartData;

            var data =  [];
            
            var chart = null;
            switch (chartType) {
                case "lineChart":
                    data = chartData;
                    chart = nv.models.lineChart()
                                .useInteractiveGuideline(true);
                    chart.xAxis.tickFormat(function(d) { return chartData[0].xlabels[d] || d; });
                    break;
                case "barChart":
                    data = chartData;
                    chart = nv.models.multiBarChart();
                    chart.xAxis.tickFormat(function(d) { return chartData[0].xlabels[d] || d; });
                    break;
                case "pieChart":
                case "pie3DChart":
                    data = chartData[0].values;
                    chart = nv.models.pieChart();
                    break;
                case "areaChart":
                    data = chartData;
                    chart = nv.models.stackedAreaChart()
                                .clipEdge(true)
                                .useInteractiveGuideline(true);
                    chart.xAxis.tickFormat(function(d) { return chartData[0].xlabels[d] || d; });
                    break;
                case "scatterChart":
                    
                    for (var i=0; i<chartData.length; i++) {
                        var arr = [];
                        for (var j=0; j<chartData[i].length; j++) {
                            arr.push({x: j, y: chartData[i][j]});
                        }
                        data.push({key: 'data' + (i + 1), values: arr});
                    }
                    
                    //data = chartData;
                    chart = nv.models.scatterChart()
                                .showDistX(true)
                                .showDistY(true)
                                .color(d3.scale.category10().range());
                    chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'));
                    chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'));
                    break;
                default:
            }
            
            if (chart !== null) {
                
                d3.select("#" + chartID)
                    .append("svg")
                    .datum(data)
                    .transition().duration(500)
                    .call(chart);
                
                nv.utils.windowResize(chart.update);
                isDone = true;
            }
            
        }

        function setNumericBullets(elem){
            var prgrphs_arry = elem;
            for(var i=0; i< prgrphs_arry.length; i++){
                var buSpan = $(prgrphs_arry[i]).find('.numeric-bullet-style');
                if(buSpan.length > 0){
                    //console.log("DIV-"+i+":");
                    var prevBultTyp = "";
                    var prevBultLvl = "";
                    var buletIndex = 0;
                    var tmpArry = new Array();
                    var tmpArryIndx = 0;
                    var buletTypSrry = new Array();
                    for(var j=0; j< buSpan.length; j++){
                        var bult_typ = $(buSpan[j]).data("bulltname");
                        var bult_lvl = $(buSpan[j]).data("bulltlvl");
                        //console.log(j+" - "+bult_typ+" lvl: "+bult_lvl );
                        if(buletIndex==0){
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                            buletIndex++;
                        }else{
                            if(bult_typ == prevBultTyp && bult_lvl == prevBultLvl){
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                buletIndex++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                            }else if(bult_typ != prevBultTyp && bult_lvl == prevBultLvl){
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                                buletIndex = 1;
                            }else if(bult_typ != prevBultTyp && Number(bult_lvl) > Number(prevBultLvl)){
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                                buletIndex = 1;
                            }else if(bult_typ != prevBultTyp && Number(bult_lvl) < Number(prevBultLvl)){
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx--;
                                buletIndex = tmpArry[tmpArryIndx]+1;
                            }
                        }
                        //console.log(buletTypSrry[tmpArryIndx]+" - "+buletIndex);
                        var numIdx = getNumTypeNum(buletTypSrry[tmpArryIndx],buletIndex);
                        $(buSpan[j]).html(numIdx);
                    }
                }
            }
        }
        function getNumTypeNum(numTyp,num){
            var rtrnNum = "";
            switch(numTyp){
                case "arabicPeriod":
                    rtrnNum = num + ". ";
                    break;
                case "arabicParenR":
                    rtrnNum = num + ") ";
                    break;					
                case "alphaLcParenR":
                    rtrnNum = alphaNumeric(num,"lowerCase") + ") ";
                    break;
                case "alphaLcPeriod":
                    rtrnNum = alphaNumeric(num,"lowerCase") + ". ";
                    break;
                            
                case "alphaUcParenR":
                    rtrnNum = alphaNumeric(num,"upperCase") + ") ";
                    break;
                case "alphaUcPeriod":
                    rtrnNum = alphaNumeric(num,"upperCase") + ". ";
                    break;

                case "romanUcPeriod":
                    rtrnNum = romanize(num) + ". ";
                    break;		
                case "romanLcParenR":
                    rtrnNum = romanize(num) + ") ";
                    break;
                case "hebrew2Minus":
                    rtrnNum = hebrew2Minus.format(num) + "-";
                    break;
                default:
                    rtrnNum = num;
            }
            return rtrnNum;
        }
        function romanize (num) {
            if (!+num)
                return false;
            var digits = String(+num).split(""),
                key = ["","C","CC","CCC","CD","D","DC","DCC","DCCC","CM",
                    "","X","XX","XXX","XL","L","LX","LXX","LXXX","XC",
                    "","I","II","III","IV","V","VI","VII","VIII","IX"],
                roman = "",
                i = 3;
            while (i--)
                roman = (key[+digits.pop() + (i * 10)] || "") + roman;
            return Array(+digits.join("") + 1).join("M") + roman;
        }
        var hebrew2Minus = archaicNumbers([
                        [1000,''],
                        [400,'ת'],
                        [300,'ש'],
                        [200,'ר'],
                        [100,'ק'],
                        [90,'צ'],
                        [80,'פ'],
                        [70,'ע'],
                        [60,'ס'],
                        [50,'נ'],
                        [40,'מ'],
                        [30,'ל'],
                        [20,'כ'],
                        [10,'י'],
                        [9,'ט'],
                        [8,'ח'],
                        [7,'ז'],
                        [6,'ו'],
                        [5,'ה'],
                        [4,'ד'],
                        [3,'ג'],
                        [2,'ב'],
                        [1,'א'],
                        [/יה/, 'ט״ו'],
                        [/יו/, 'ט״ז'],
                        [/([א-ת])([א-ת])$/, '$1״$2'], 
                        [/^([א-ת])$/, "$1׳"] 
        ]); 
        function archaicNumbers(arr){
            var arrParse = arr.slice().sort(function (a,b) {return b[1].length - a[1].length});
            return {
                format: function(n){
                    var ret = '';
                    jQuery.each(arr, function(){
                        var num = this[0];
                        if (parseInt(num) > 0){
                            for (; n >= num; n -= num) ret += this[1];
                        }else{
                            ret = ret.replace(num, this[1]);
                        }
                    });
                    return ret; 
                }
            }
        }
        function alphaNumeric(num,upperLower){
            num = Number(num)-1;
            var aNum = "";
            if(upperLower=="upperCase"){
                aNum = (( (num/26>=1)? String.fromCharCode(num/26+64):'') + String.fromCharCode(num%26+65)).toUpperCase();
            }else if(upperLower=="lowerCase"){
                aNum = (( (num/26>=1)? String.fromCharCode(num/26+64):'') + String.fromCharCode(num%26+65)).toLowerCase();
            }
            return aNum;
        }
        function base64ArrayBuffer(arrayBuffer) {
            var base64    = '';
            var encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
            var bytes         = new Uint8Array(arrayBuffer);
            var byteLength    = bytes.byteLength;
            var byteRemainder = byteLength % 3;
            var mainLength    = byteLength - byteRemainder;
        
            var a, b, c, d;
            var chunk;
        
            for (var i = 0; i < mainLength; i = i + 3) {
                chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
                a = (chunk & 16515072) >> 18;
                b = (chunk & 258048)   >> 12;
                c = (chunk & 4032)     >>  6;
                d = chunk & 63;
                base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
            }
        
            if (byteRemainder == 1) {
                chunk = bytes[mainLength];
                a = (chunk & 252) >> 2;
                b = (chunk & 3)   << 4;
                base64 += encodings[a] + encodings[b] + '==';
            } else if (byteRemainder == 2) {
                chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
                a = (chunk & 64512) >> 10;
                b = (chunk & 1008)  >>  4;
                c = (chunk & 15)    <<  2;
                base64 += encodings[a] + encodings[b] + encodings[c] + '=';
            }
        
            return base64;
        }
        
        function extractFileExtension(filename) {
            return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
        }
        
        function escapeHtml(text) {
            var map = {
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#039;'
            };
            return text.replace(/[&<>"']/g, function(m) { return map[m]; });
        }
        /////////////////////////////////////tXml///////////////////////////
        function tXml(S) {
            
            "use strict";
            var openBracket = "<";
            var openBracketCC = "<".charCodeAt(0);
            var closeBracket = ">";
            var closeBracketCC = ">".charCodeAt(0);
            var minus = "-";
            var minusCC = "-".charCodeAt(0);
            var slash = "/";
            var slashCC = "/".charCodeAt(0);
            var exclamation = '!';
            var exclamationCC = '!'.charCodeAt(0);
            var singleQuote = "'";
            var singleQuoteCC = "'".charCodeAt(0);
            var doubleQuote = '"';
            var doubleQuoteCC = '"'.charCodeAt(0);
            var questionMark = '?';
            var questionMarkCC = '?'.charCodeAt(0);
            
            /**
             *    returns text until the first nonAlphebetic letter
             */
            var nameSpacer = "\r\n\t>/= ";
            
            var pos = 0;
            
            /**
             * Parsing a list of entries
             */
            function parseChildren() {
                var children = [];
                while (S[pos]) {
                    if (S.charCodeAt(pos) == openBracketCC) {
                        if (S.charCodeAt(pos+1) === slashCC) { // </
                            //while (S[pos]!=='>') { pos++; }
                            pos = S.indexOf(closeBracket, pos);
                            return children;
                        } else if (S.charCodeAt(pos+1) === exclamationCC) { // <! or <!--
                            if (S.charCodeAt(pos+2) == minusCC) {
                                // comment support
                                while (!(S.charCodeAt(pos) === closeBracketCC && S.charCodeAt(pos-1) == minusCC && 
                                        S.charCodeAt(pos-2) == minusCC && pos != -1)) {
                                    pos = S.indexOf(closeBracket, pos+1);
                                }
                                if (pos === -1) {
                                    pos = S.length;
                                }
                            } else {
                                // doctype support
                                pos += 2;
                                for (; S.charCodeAt(pos) !== closeBracketCC; pos++) {}
                            }
                            pos++;
                            continue;
                        } else if (S.charCodeAt(pos+1) === questionMarkCC) { // <?
                            // XML header support
                            pos = S.indexOf(closeBracket, pos);
                            pos++;
                            continue;
                        }
                        pos++;
                        var startNamePos = pos;
                        for (; nameSpacer.indexOf(S[pos]) === -1; pos++) {}
                        var node_tagName = S.slice(startNamePos, pos);
        
                        // Parsing attributes
                        var attrFound = false;
                        var node_attributes = {};
                        for (; S.charCodeAt(pos) !== closeBracketCC; pos++) {
                            var c = S.charCodeAt(pos);
                            if ((c > 64 && c < 91) || (c > 96 && c < 123)) {
                                startNamePos = pos;
                                for (; nameSpacer.indexOf(S[pos]) === -1; pos++) {}
                                var name = S.slice(startNamePos, pos);
                                // search beginning of the string
                                var code = S.charCodeAt(pos);
                                while (code !== singleQuoteCC && code !== doubleQuoteCC) {
                                    pos++;
                                    code = S.charCodeAt(pos);
                                }
                                
                                var startChar = S[pos];
                                var startStringPos= ++pos;
                                pos = S.indexOf(startChar, startStringPos);
                                var value = S.slice(startStringPos, pos);
                                if (!attrFound) {
                                    node_attributes = {};
                                    attrFound = true;
                                }
                                node_attributes[name] = value;
                            }
                        }
                        
                        // Optional parsing of children
                        if (S.charCodeAt(pos-1) !== slashCC) {
                            pos++;
                            var node_children = parseChildren();
                        }
                        
                        children.push({
                            "children": node_children,
                            "tagName": node_tagName,
                            "attrs": node_attributes
                        });
                        
                    } else {
                        var startTextPos = pos;
                        pos = S.indexOf(openBracket, pos) - 1; // Skip characters until '<'
                        if (pos === -2) {
                            pos = S.length;
                        }
                        var text = S.slice(startTextPos, pos + 1);
                        if (text.trim().length > 0) {
                            children.push(text);
                        }
                    }
                    pos++;
                }
                return children;
            }
            
            _order = 1;
            return simplefy(parseChildren());
        }
        
        function simplefy(children) {
            var node = {};
            
            if (children === undefined) {
                return {};
            }
            
            // Text node (e.g. <t>This is text.</t>)
            if (children.length === 1 && typeof children[0] == 'string') {
                return children[0];
            }
        
            // map each object
            children.forEach(function (child) {
        
                if (!node[child.tagName]) {
                    node[child.tagName] = [];
                }
        
                if (typeof child === 'object') {
                    var kids = simplefy(child.children);
                    if (child.attrs) {
                        kids.attrs = child.attrs;
                    }
                    
                    if (kids["attrs"] === undefined) {
                        kids["attrs"] = {"order": _order};
                    } else {
                        kids["attrs"]["order"] = _order;
                    }
                    _order++;
                    node[child.tagName].push(kids);
                }
            });
            
            for (var i in node) {
                if (node[i].length == 1) {
                    node[i] = node[i][0];
                }
            }
            
            return node;
        };
    };

    /*!
    JSZipUtils - A collection of cross-browser utilities to go along with JSZip.
    <http://stuk.github.io/jszip-utils>
    (c) 2014 Stuart Knightley, David Duponchel
    Dual licenced under the MIT license or GPLv3. See https://raw.github.com/Stuk/jszip-utils/master/LICENSE.markdown.
    */
    !function(a){"object"==typeof exports?module.exports=a():"function"==typeof define&&define.amd?define(a):"undefined"!=typeof window?window.JSZipUtils=a():"undefined"!=typeof global?global.JSZipUtils=a():"undefined"!=typeof self&&(self.JSZipUtils=a())}(function(){return function a(b,c,d){function e(g,h){if(!c[g]){if(!b[g]){var i="function"==typeof require&&require;if(!h&&i)return i(g,!0);if(f)return f(g,!0);throw new Error("Cannot find module '"+g+"'")}var j=c[g]={exports:{}};b[g][0].call(j.exports,function(a){var c=b[g][1][a];return e(c?c:a)},j,j.exports,a,b,c,d)}return c[g].exports}for(var f="function"==typeof require&&require,g=0;g<d.length;g++)e(d[g]);return e}({1:[function(a,b){"use strict";function c(){try{return new window.XMLHttpRequest}catch(a){}}function d(){try{return new window.ActiveXObject("Microsoft.XMLHTTP")}catch(a){}}var e={};e._getBinaryFromXHR=function(a){return a.response||a.responseText};var f=window.ActiveXObject?function(){return c()||d()}:c;e.getBinaryContent=function(a,b){try{var c=f();c.open("GET",a,!0),"responseType"in c&&(c.responseType="arraybuffer"),c.overrideMimeType&&c.overrideMimeType("text/plain; charset=x-user-defined"),c.onreadystatechange=function(){var d,f;if(4===c.readyState)if(200===c.status||0===c.status){d=null,f=null;try{d=e._getBinaryFromXHR(c)}catch(g){f=new Error(g)}b(f,d)}else b(new Error("Ajax error for "+a+" : "+this.status+" "+this.statusText),null)},c.send()}catch(d){b(new Error(d),null)}},b.exports=e},{}]},{},[1])(1)});                
    
    /**
     * Colorz (or Colz) is a Javascript "library" to help
     * in color conversion between the usual color-spaces
     * Hex - Rgb - Hsl / Hsv - Hsb
     *
     * It provides some helpers to output Canvas / CSS
     * color strings.
     *
     * by Carlos Cabo 2013
     * http://carloscabo.com
     *
     * Some formulas borrowed from Wikipedia or other authors.
    */

    (function(name, definition) {
        if (typeof define === "function") {
        define(definition);
        } else if (typeof module !== "undefined" && module.exports) {
        module.exports = definition();
        } else {
        var theModule = definition(),
            global = this,
            old = global[name];
        theModule.noConflict = function() {
            global[name] = old;
            return theModule;
        };
        global[name] = theModule;
        }
    })("colz", function() {
        var round = Math.round,
        toString = "toString",
        colz = colz || {},
        Rgb,
        Rgba,
        Hsl,
        Hsla,
        Color,
        ColorScheme,
        hexToRgb,
        componentToHex,
        rgbToHex,
        rgbToHsl,
        hue2rgb,
        hslToRgb,
        rgbToHsb,
        hsbToRgb,
        hsbToHsl,
        hsvToHsl,
        hsvToRgb,
        randomColor;
        Rgb = colz.Rgb = function(col) {
        this.r = col[0];
        this.g = col[1];
        this.b = col[2];
        };
        Rgb.prototype[toString] = function() {
        return "rgb(" + this.r + "," + this.g + "," + this.b + ")";
        };
        Rgba = colz.Rgba = function(col) {
        this.r = col[0];
        this.g = col[1];
        this.b = col[2];
        this.a = col[3];
        };
        Rgba.prototype[toString] = function() {
        return (
            "rgba(" + this.r + "," + this.g + "," + this.b + "," + this.a + ")"
        );
        };
        Hsl = colz.Hsl = function(col) {
        this.h = col[0];
        this.s = col[1];
        this.l = col[2];
        };
        Hsl.prototype[toString] = function() {
        return "hsl(" + this.h + "," + this.s + "%," + this.l + "%)";
        };
        Hsla = colz.Hsla = function(col) {
        this.h = col[0];
        this.s = col[1];
        this.l = col[2];
        this.a = col[3];
        };
        Hsla.prototype[toString] = function() {
        return (
            "hsla(" + this.h + "," + this.s + "%," + this.l + "%," + this.a + ")"
        );
        };
        Color = colz.Color = function() {
        this.hex = this.r = this.g = this.b = this.h = this.s = this.l = this.a = this.hsl = this.hsla = this.rgb = this.rgba = null;
        this.init(arguments);
        };
        var colorPrototype = Color.prototype;
        colorPrototype.init = function(arg) {
        var _this = this;
        if (typeof arg[0] === "string") {
            if (arg[0].charAt(0) !== "#") {
            arg[0] = "#" + arg[0];
            }
            if (arg[0].length < 7) {
            arg[0] =
                "#" +
                arg[0][1] +
                arg[0][1] +
                arg[0][2] +
                arg[0][2] +
                arg[0][3] +
                arg[0][3];
            }
            _this.hex = arg[0].toLowerCase();
            _this.rgb = new Rgb(hexToRgb(_this.hex));
            _this.r = _this.rgb.r;
            _this.g = _this.rgb.g;
            _this.b = _this.rgb.b;
            _this.a = 1;
            _this.rgba = new Rgba([_this.r, _this.g, _this.b, _this.a]);
        }
        if (typeof arg[0] === "number") {
            _this.r = arg[0];
            _this.g = arg[1];
            _this.b = arg[2];
            if (typeof arg[3] === "undefined") {
            _this.a = 1;
            } else {
            _this.a = arg[3];
            }
            _this.rgb = new Rgb([_this.r, _this.g, _this.b]);
            _this.rgba = new Rgba([_this.r, _this.g, _this.b, _this.a]);
            _this.hex = rgbToHex([_this.r, _this.g, _this.b]);
        }
        if (arg[0] instanceof Array) {
            _this.r = arg[0][0];
            _this.g = arg[0][1];
            _this.b = arg[0][2];
            if (typeof arg[0][3] === "undefined") {
            _this.a = 1;
            } else {
            _this.a = arg[0][3];
            }
            _this.rgb = new Rgb([_this.r, _this.g, _this.b]);
            _this.rgba = new Rgba([_this.r, _this.g, _this.b, _this.a]);
            _this.hex = rgbToHex([_this.r, _this.g, _this.b]);
        }
        _this.hsl = new Hsl(colz.rgbToHsl([_this.r, _this.g, _this.b]));
        _this.h = _this.hsl.h;
        _this.s = _this.hsl.s;
        _this.l = _this.hsl.l;
        _this.hsla = new Hsla([_this.h, _this.s, _this.l, _this.a]);
        };
        colorPrototype.setHue = function(newhue) {
        var _this = this;
        _this.h = newhue;
        _this.hsl.h = newhue;
        _this.hsla.h = newhue;
        _this.updateFromHsl();
        };
        colorPrototype.setSat = function(newsat) {
        var _this = this;
        _this.s = newsat;
        _this.hsl.s = newsat;
        _this.hsla.s = newsat;
        _this.updateFromHsl();
        };
        colorPrototype.setLum = function(newlum) {
        var _this = this;
        _this.l = newlum;
        _this.hsl.l = newlum;
        _this.hsla.l = newlum;
        _this.updateFromHsl();
        };
        colorPrototype.setAlpha = function(newalpha) {
        this.a = newalpha;
        this.hsla.a = newalpha;
        this.rgba.a = newalpha;
        };
        colorPrototype.updateFromHsl = function() {
        this.rgb = null;
        this.rgb = new Rgb(colz.hslToRgb([this.h, this.s, this.l]));
        this.r = this.rgb.r;
        this.g = this.rgb.g;
        this.b = this.rgb.b;
        this.rgba.r = this.rgb.r;
        this.rgba.g = this.rgb.g;
        this.rgba.b = this.rgb.b;
        this.hex = null;
        this.hex = rgbToHex([this.r, this.g, this.b]);
        };
        randomColor = colz.randomColor = function() {
        var r = "#" + Math.random().toString(16).slice(2, 8);
        return new Color(r);
        };
        hexToRgb = colz.hexToRgb = function(hex) {
        var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result
            ? [
                parseInt(result[1], 16),
                parseInt(result[2], 16),
                parseInt(result[3], 16)
            ]
            : null;
        };
        componentToHex = colz.componentToHex = function(c) {
        var hex = c.toString(16);
        return hex.length === 1 ? "0" + hex : hex;
        };
        rgbToHex = colz.rgbToHex = function() {
        var arg, r, g, b;
        arg = arguments;
        if (arg.length > 1) {
            r = arg[0];
            g = arg[1];
            b = arg[2];
        } else {
            r = arg[0][0];
            g = arg[0][1];
            b = arg[0][2];
        }
        return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
        };
        rgbToHsl = colz.rgbToHsl = function() {
        var arg, r, g, b, h, s, l, d, max, min;
        arg = arguments;
        if (typeof arg[0] === "number") {
            r = arg[0];
            g = arg[1];
            b = arg[2];
        } else {
            r = arg[0][0];
            g = arg[0][1];
            b = arg[0][2];
        }
        r /= 255;
        g /= 255;
        b /= 255;
        max = Math.max(r, g, b);
        min = Math.min(r, g, b);
        l = (max + min) / 2;
        if (max === min) {
            h = s = 0;
        } else {
            d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
            case r:
                h = (g - b) / d + (g < b ? 6 : 0);
                break;
            case g:
                h = (b - r) / d + 2;
                break;
            case b:
                h = (r - g) / d + 4;
                break;
            }
            h /= 6;
        }
        h = round(h * 360);
        s = round(s * 100);
        l = round(l * 100);
        return [h, s, l];
        };
        hue2rgb = colz.hue2rgb = function(p, q, t) {
        if (t < 0) {
            t += 1;
        }
        if (t > 1) {
            t -= 1;
        }
        if (t < 1 / 6) {
            return p + (q - p) * 6 * t;
        }
        if (t < 1 / 2) {
            return q;
        }
        if (t < 2 / 3) {
            return p + (q - p) * (2 / 3 - t) * 6;
        }
        return p;
        };
        hslToRgb = colz.hslToRgb = function() {
        var arg, r, g, b, h, s, l, q, p;
        arg = arguments;
        if (typeof arg[0] === "number") {
            h = arg[0] / 360;
            s = arg[1] / 100;
            l = arg[2] / 100;
        } else {
            h = arg[0][0] / 360;
            s = arg[0][1] / 100;
            l = arg[0][2] / 100;
        }
        if (s === 0) {
            r = g = b = l;
        } else {
            q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            p = 2 * l - q;
            r = colz.hue2rgb(p, q, h + 1 / 3);
            g = colz.hue2rgb(p, q, h);
            b = colz.hue2rgb(p, q, h - 1 / 3);
        }
        return [round(r * 255), round(g * 255), round(b * 255)];
        };
        rgbToHsb = colz.rgbToHsb = function(r, g, b) {
        var max, min, h, s, v, d;
        r = r / 255;
        g = g / 255;
        b = b / 255;
        max = Math.max(r, g, b);
        min = Math.min(r, g, b);
        v = max;
        d = max - min;
        s = max === 0 ? 0 : d / max;
        if (max === min) {
            h = 0;
        } else {
            switch (max) {
            case r:
                h = (g - b) / d + (g < b ? 6 : 0);
                break;
            case g:
                h = (b - r) / d + 2;
                break;
            case b:
                h = (r - g) / d + 4;
                break;
            }
            h /= 6;
        }
        h = round(h * 360);
        s = round(s * 100);
        v = round(v * 100);
        return [h, s, v];
        };
        hsbToRgb = colz.hsbToRgb = function(h, s, v) {
        var r, g, b, i, f, p, q, t;
        if (v === 0) {
            return [0, 0, 0];
        }
        s = s / 100;
        v = v / 100;
        h = h / 60;
        i = Math.floor(h);
        f = h - i;
        p = v * (1 - s);
        q = v * (1 - s * f);
        t = v * (1 - s * (1 - f));
        if (i === 0) {
            r = v;
            g = t;
            b = p;
        } else if (i === 1) {
            r = q;
            g = v;
            b = p;
        } else if (i === 2) {
            r = p;
            g = v;
            b = t;
        } else if (i === 3) {
            r = p;
            g = q;
            b = v;
        } else if (i === 4) {
            r = t;
            g = p;
            b = v;
        } else if (i === 5) {
            r = v;
            g = p;
            b = q;
        }
        r = Math.floor(r * 255);
        g = Math.floor(g * 255);
        b = Math.floor(b * 255);
        return [r, g, b];
        };
        hsbToHsl = colz.hsbToHsl = function(h, s, b) {
        return colz.rgbToHsl(colz.hsbToRgb(h, s, b));
        };
        hsvToHsl = colz.hsvToHsl = colz.hsbToHsl;
        hsvToRgb = colz.hsvToRgb = colz.hsbToRgb;
        ColorScheme = colz.ColorScheme = function(color_val, angle_array) {
        this.palette = [];
        if (angle_array === undefined && color_val instanceof Array) {
            this.createFromColors(color_val);
        } else {
            this.createFromAngles(color_val, angle_array);
        }
        };
        var colorSchemePrototype = ColorScheme.prototype;
        colorSchemePrototype.createFromColors = function(color_val) {
        for (var i in color_val) {
            if (color_val.hasOwnProperty(i)) {
            this.palette.push(new Color(color_val[i]));
            }
        }
        return this.palette;
        };
        colorSchemePrototype.createFromAngles = function(color_val, angle_array) {
        this.palette.push(new Color(color_val));
        for (var i in angle_array) {
            if (angle_array.hasOwnProperty(i)) {
            var tempHue = (this.palette[0].h + angle_array[i]) % 360;
            this.palette.push(
                new Color(
                colz.hslToRgb([tempHue, this.palette[0].s, this.palette[0].l])
                )
            );
            }
        }
        return this.palette;
        };
        ColorScheme.Compl = function(color_val) {
        return new ColorScheme(color_val, [180]);
        };
        ColorScheme.Triad = function(color_val) {
        return new ColorScheme(color_val, [120, 240]);
        };
        ColorScheme.Tetrad = function(color_val) {
        return new ColorScheme(color_val, [60, 180, 240]);
        };
        ColorScheme.Analog = function(color_val) {
        return new ColorScheme(color_val, [-45, 45]);
        };
        ColorScheme.Split = function(color_val) {
        return new ColorScheme(color_val, [150, 210]);
        };
        ColorScheme.Accent = function(color_val) {
        return new ColorScheme(color_val, [-45, 45, 180]);
        };
        return colz;
    });
}( jQuery ));
