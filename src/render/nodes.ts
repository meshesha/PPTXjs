// Ported from PPTXjs v1 (legacy/js/pptxjs.js). Internal rendering code: the
// node objects come from the simplified XML tree, so most values are `any`.
import type { WarpObj } from "./context";
import { genTable } from "./table";
import { genShape } from "./shape";
import { genChart } from "./chart";
import { asArray, getTextByPathList } from "../core/xml";
import { IsVideoLink, extractFileExtension, getMimeType } from "../core/package";
import { base64ArrayBuffer } from "../utils/base64";
import { escapeHtml } from "../utils/html";
import { angleToDegrees, slideFactor } from "./units";
// ported from legacy/js/pptxjs.js:781
export function processNodesInSlide(nodeKey?: any, nodeValue?: any, nodes?: any, warpObj?: any, source?: any, sType?: any): any {
    var result: any = "";

    switch (nodeKey) {
        case "p:sp":    // Shape, Text
            result = processSpNode(nodeValue, nodes, warpObj, source, sType);
            break;
        case "p:cxnSp":    // Shape, Text (with connection)
            result = processCxnSpNode(nodeValue, nodes, warpObj, source, sType);
            break;
        case "p:pic":    // Picture
            result = processPicNode(nodeValue, warpObj, source, sType);
            break;
        case "p:graphicFrame":    // Chart, Diagram, Table
            result = processGraphicFrameNode(nodeValue, warpObj, source, sType);
            break;
        case "p:grpSp":
            result = processGroupSpNode(nodeValue, warpObj, source);
            break;
        case "mc:AlternateContent": //Equations and formulas as Image
            //console.log("mc:AlternateContent nodeValue:" , nodeValue , "nodes:",nodes, "sType:",sType)
            var mcFallbackNode: any = getTextByPathList(nodeValue, ["mc:Fallback"]) ?? asArray(getTextByPathList(nodeValue, ["mc:Choice"]))[0];
            if (mcFallbackNode !== undefined) {
                result = processGroupSpNode(mcFallbackNode, warpObj, source);
            }
            break;
        default:
            //console.log("nodeKey: ", nodeKey)
    }

    // top-level slide shapes are tagged for hit-testing (editor, click handlers);
    // display:contents keeps the wrapper out of the layout
    if (source == "slide" && sType === undefined && result !== "") {
        var shapeId: any = shapeIdOf(nodeKey, nodeValue);
        if (shapeId !== undefined) {
            result = "<div class='pptx-shape' data-shape-id='" + shapeId + "' style='display:contents'>" + result + "</div>";
        }
    }
    return result;

}

const NV_PROPS: Record<string, string> = {
    "p:sp": "p:nvSpPr",
    "p:pic": "p:nvPicPr",
    "p:graphicFrame": "p:nvGraphicFramePr",
    "p:grpSp": "p:nvGrpSpPr",
    "p:cxnSp": "p:nvCxnSpPr",
};

/** `p:cNvPr/@id` of a shape node. */
export function shapeIdOf(nodeKey: string, node: any): string | undefined {
    var nv: string | undefined = NV_PROPS[nodeKey];
    return nv ? getTextByPathList(node, [nv, "p:cNvPr", "attrs", "id"]) : undefined;
}

// ported from legacy/js/pptxjs.js:813
export function processGroupSpNode(node?: any, warpObj?: any, source?: any): any {
    //console.log("processGroupSpNode: node: ", node)
    var xfrmNode: any = getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
    if (xfrmNode !== undefined) {
        var x: any = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * slideFactor;
        var y: any = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * slideFactor;
        var chx: any = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * slideFactor;
        var chy: any = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * slideFactor;
        var cx: any = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * slideFactor;
        var cy: any = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * slideFactor;
        var chcx: any = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * slideFactor;
        var chcy: any = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * slideFactor;
        var rotate: any = parseInt(xfrmNode["attrs"]["rot"])
        var rotStr: any = ""//;" border: 3px solid black;";
        // angleToDegrees(getTextByPathList(slideXfrmNode, ["attrs", "rot"]));
        // var rotX = 0;
        // var rotY = 0;
        var top: any = y - chy,
            left: any = x - chx,
            width: any = cx - chcx,
            height: any = cy - chcy;

        var sType: any = "group";
        if (!isNaN(rotate)) {
            rotate = angleToDegrees(rotate);
            rotStr += "transform: rotate(" + rotate + "deg) ; transform-origin: center;";
            // var cLin = Math.sqrt(Math.pow((chy), 2) + Math.pow((chx), 2));
            // var rdian = degreesToRadians(rotate);
            // rotX = cLin * Math.cos(rdian);
            // rotY = cLin * Math.sin(rdian);
            if (rotate != 0) {
                top = y;
                left = x;
                width = cx;
                height = cy;
                sType = "group-rotate";
            }

        }
    }
    var grpStyle: any = "";

    if (rotStr !== undefined && rotStr != "") {
        grpStyle += rotStr;
    }

    if (top !== undefined) {
        grpStyle += "top: " + top + "px;";
    }
    if (left !== undefined) {
        grpStyle += "left: " + left + "px;";
    }
    if (width !== undefined) {
        grpStyle += "width:" + width + "px;";
    }
    if (height !== undefined) {
        grpStyle += "height: " + height + "px;";
    }
    var order: any = node["attrs"]["order"];

    var result: any = "<div class='block group' style='z-index: " + order + ";" + grpStyle + "'>";

    // Procsee all child nodes
    for (var nodeKey in node) {
        if (node[nodeKey].constructor === Array) {
            for (var i: any = 0; i < node[nodeKey].length; i++) {
                result += processNodesInSlide(nodeKey, node[nodeKey][i], node, warpObj, source, sType);
            }
        } else {
            result += processNodesInSlide(nodeKey, node[nodeKey], node, warpObj, source, sType);
        }
    }

    result += "</div>";

    return result;
}

// ported from legacy/js/pptxjs.js:891
export function processSpNode(node?: any, pNode?: any, warpObj?: any, source?: any, sType?: any): any {

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

    var id: any = getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "id"]);
    var name: any = getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
    var idx: any = (getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]) === undefined) ? undefined : getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]);
    var type: any = (getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]) === undefined) ? undefined : getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
    var order: any = getTextByPathList(node, ["attrs", "order"]);
    var isUserDrawnBg: any;
    if (source == "slideLayoutBg" || source == "slideMasterBg") {
        var userDrawn: any = getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "attrs", "userDrawn"]);
        if (userDrawn == "1") {
            isUserDrawnBg = true;
        } else {
            isUserDrawnBg = false;
        }
    }
    var slideLayoutSpNode: any = undefined;
    var slideMasterSpNode: any = undefined;

    if (idx !== undefined) {
        slideLayoutSpNode = warpObj["slideLayoutTables"]["idxTable"][idx];
        if (type !== undefined) {
            slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
        } else {
            slideMasterSpNode = warpObj["slideMasterTables"]["idxTable"][idx];
        }
    } else {
        if (type !== undefined) {
            slideLayoutSpNode = warpObj["slideLayoutTables"]["typeTable"][type];
            slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
        }
    }

    if (type === undefined) {
        var txBoxVal: any = getTextByPathList(node, ["p:nvSpPr", "p:cNvSpPr", "attrs", "txBox"]);
        if (txBoxVal == "1") {
            type = "textBox";
        }
    }
    if (type === undefined) {
        type = getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (type === undefined) {
            //type = getTextByPathList(slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
            if (source == "diagramBg") {
                type = "diagram";
            } else {

                type = "obj"; //default type
            }
        }
    }
    //console.log("processSpNode type:", type, "idx:", idx);
    return genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source);
}

// ported from legacy/js/pptxjs.js:958
export function processCxnSpNode(node?: any, pNode?: any, warpObj?: any, source?: any, sType?: any): any {

    var id: any = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
    var name: any = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
    var idx: any = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
    var type: any = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
    //<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
    var order: any = node["attrs"]["order"];

    return genShape(node, pNode, undefined, undefined, id, name, idx, type, order, warpObj, undefined, sType, source);
}

// ported from legacy/js/pptxjs.js:8379
export function processPicNode(node?: any, warpObj?: any, source?: any, sType?: any): any {
    //console.log("processPicNode node:", node, "source:", source, "sType:", sType, "warpObj;", warpObj);
    var rtrnData: any = "";
    var mediaPicFlag: any = false;
    var order: any = node["attrs"]["order"];

    var rid: any = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
    var resObj: any;
    if (source == "slideMasterBg") {
        resObj = warpObj["masterResObj"];
    } else if (source == "slideLayoutBg") {
        resObj = warpObj["layoutResObj"];
    } else {
        //imgName = warpObj["slideResObj"][rid]["target"];
        resObj = warpObj["slideResObj"];
    }
    var imgName: any = resObj[rid]["target"];

    //console.log("processPicNode imgName:", imgName);
    var imgFileExt: any = extractFileExtension(imgName).toLowerCase();
    var zip: any = warpObj["zip"];
    var imgArrayBuffer: any = zip.bytes(imgName);
    if (imgArrayBuffer === null) {
        warpObj.ctx.warn("Missing image part: " + imgName);
        return "";
    }
    var mimeType: any = "";
    var xfrmNode: any = node["p:spPr"]["a:xfrm"];
    if (xfrmNode === undefined) {
        var idx: any = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
        var type: any = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (idx !== undefined) {
            xfrmNode = getTextByPathList(warpObj["slideLayoutTables"], ["idxTable", idx, "p:spPr", "a:xfrm"]);
        }
    }
    ///////////////////////////////////////Amir//////////////////////////////
    var rotate: any = 0;
    var rotateNode: any = getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
    if (rotateNode !== undefined) {
        rotate = angleToDegrees(rotateNode);
    }
    //video
    var vdoNode: any = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
    var vdoRid: any, vdoFile: any, vdoFileExt: any, vdoMimeType: any, uInt8Array: any, blob: any, vdoBlob: any, mediaSupportFlag: any = false, isVdeoLink: any = false;
    var mediaProcess: any = warpObj.ctx.mediaProcess;
    if (vdoNode !== undefined && mediaProcess) {
        vdoRid = vdoNode["attrs"]["r:link"];
        vdoFile = resObj[vdoRid]["target"];
        var checkIfLink: any = IsVideoLink(vdoFile);
        if (checkIfLink) {
            vdoFile = escapeHtml(vdoFile);
            //vdoBlob = vdoFile;
            isVdeoLink = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        } else {
            vdoFileExt = extractFileExtension(vdoFile).toLowerCase();
            if (vdoFileExt == "mp4" || vdoFileExt == "webm" || vdoFileExt == "ogg") {
                uInt8Array = zip.bytes(vdoFile);
                vdoMimeType = getMimeType(vdoFileExt);
                blob = new Blob([uInt8Array], {
                    type: vdoMimeType
                });
                vdoBlob = warpObj.ctx.createObjectUrl(blob);
                mediaSupportFlag = vdoBlob !== "";
                mediaPicFlag = true;
            }
        }
    }
    //Audio
    var audioNode: any = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
    var audioRid: any, audioFile: any, audioFileExt: any, audioMimeType: any, uInt8ArrayAudio: any, blobAudio: any, audioBlob: any;
    var audioPlayerFlag: any = false;
    var audioObjc: any;
    if (audioNode !== undefined && mediaProcess) {
        audioRid = audioNode["attrs"]["r:link"];
        audioFile = resObj[audioRid]["target"];
        audioFileExt = extractFileExtension(audioFile).toLowerCase();
        if (audioFileExt == "mp3" || audioFileExt == "wav" || audioFileExt == "ogg") {
            uInt8ArrayAudio = zip.bytes(audioFile);
            blobAudio = new Blob([uInt8ArrayAudio], { type: getMimeType(audioFileExt) });
            audioBlob = warpObj.ctx.createObjectUrl(blobAudio);
            var cx: any = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * 20;
            var cy: any = xfrmNode["a:ext"]["attrs"]["cy"];
            var x: any = parseInt(xfrmNode["a:off"]["attrs"]["x"]) / 2.5;
            var y: any = xfrmNode["a:off"]["attrs"]["y"];
            audioObjc = {
                "a:ext": {
                    "attrs": {
                        "cx": cx,
                        "cy": cy
                    }
                },
                "a:off": {
                    "attrs": {
                        "x": x,
                        "y": y

                    }
                }
            }
            audioPlayerFlag = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        }
    }
    //console.log(node)
    //////////////////////////////////////////////////////////////////////////
    mimeType = getMimeType(imgFileExt);
    rtrnData = "<div class='block content' style='" +
        ((mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, node, undefined, undefined) : getPosition(xfrmNode, node, undefined, undefined)) +
        ((mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)) +
        " z-index: " + order + ";" +
        "transform: rotate(" + rotate + "deg);'>";
    if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
        rtrnData += "<img src='data:" + mimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/>";
    } else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
        if (vdoNode !== undefined && !isVdeoLink) {
            rtrnData += "<video src='" + vdoBlob + "' controls preload='metadata' style='width: 100%; height: 100%'></video>";
        } else if (vdoNode !== undefined && isVdeoLink) {
            if (warpObj.ctx.externalMedia) {
                // third-party content: sandboxed, no referrer
                rtrnData += "<iframe src='" + vdoFile + "' sandbox='allow-scripts allow-same-origin allow-presentation' referrerpolicy='no-referrer' allow='fullscreen' style='width: 100%; height: 100%; border: 0'></iframe>";
            } else {
                rtrnData += "<img src='data:" + mimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/>" +
                    "<a class='pptx-media-link' href='" + vdoFile + "' target='_blank' rel='noopener noreferrer' title='" + vdoFile + "'>&#9654;</a>";
            }
        }
        if (audioNode !== undefined) {
            rtrnData += '<audio controls preload="metadata"><source src="' + audioBlob + '"></audio>';
            //'<button onclick="audio_player.play()">Play</button>'+
            //'<button onclick="audio_player.pause()">Pause</button>';
        }
    }
    if (!mediaSupportFlag && mediaPicFlag) {
        rtrnData += "<span style='color:red;font-size:40px;position: absolute;'>This media file Not supported by HTML5</span>";
    }
    if ((vdoNode !== undefined || audioNode !== undefined) && !mediaProcess && mediaSupportFlag) {
        /* debug log removed */
    }
    rtrnData += "</div>";
    //console.log(rtrnData)
    return rtrnData;
}

// ported from legacy/js/pptxjs.js:8514
export function processGraphicFrameNode(node?: any, warpObj?: any, source?: any, sType?: any): any {

    var result: any = "";
    var graphicTypeUri: any = getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);

    switch (graphicTypeUri) {
        case "http://schemas.openxmlformats.org/drawingml/2006/table":
            result = genTable(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/chart":
            result = genChart(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
            result = genDiagram(node, warpObj, source, sType);
            break;
        case "http://schemas.openxmlformats.org/presentationml/2006/ole":
            //result = genDiagram(node, warpObj, source, sType);
            var oleObjNode: any = getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent", "mc:Fallback","p:oleObj"]);
            
            if (oleObjNode === undefined) {
                oleObjNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "p:oleObj"]);
            }
            //console.log("node:", node, "oleObjNode:", oleObjNode)
            if (oleObjNode !== undefined){
                result = processGroupSpNode(oleObjNode, warpObj, source);
            }
            break;
        default:
    }

    return result;
}

// ported from legacy/js/pptxjs.js:8547
export function processSpPrNode(node?: any, warpObj?: any): any {

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

var is_first_br: any = false;

// ported from legacy/js/pptxjs.js:10320
export function genDiagram(node?: any, warpObj?: any, source?: any, sType?: any): any {
    //console.log(warpObj)
    //zip.xmlFresh(sldFileName)
    /**files define the diagram:
     * 1-colors#.xml,
     * 2-data#.xml, 
     * 3-layout#.xml,
     * 4-quickStyle#.xml.
     * 5-drawing#.xml, which Microsoft added as an extension for persisting diagram layout information.
     */
    ///get colors#.xml, data#.xml , layout#.xml , quickStyle#.xml
    var order: any = node["attrs"]["order"];
    var zip: any = warpObj["zip"];
    var xfrmNode: any = getTextByPathList(node, ["p:xfrm"]);
    var dgmRelIds: any = getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);
    //console.log(dgmRelIds)
    var dgmClrFileId: any = dgmRelIds["r:cs"];
    var dgmDataFileId: any = dgmRelIds["r:dm"];
    var dgmLayoutFileId: any = dgmRelIds["r:lo"];
    var dgmQuickStyleFileId: any = dgmRelIds["r:qs"];
    var dgmClrFileName: any = warpObj["slideResObj"][dgmClrFileId].target,
        dgmDataFileName: any = warpObj["slideResObj"][dgmDataFileId].target,
        dgmLayoutFileName: any = warpObj["slideResObj"][dgmLayoutFileId].target;
    var dgmQuickStyleFileName: any = warpObj["slideResObj"][dgmQuickStyleFileId].target;
    //console.log("dgmClrFileName: " , dgmClrFileName,", dgmDataFileName: ",dgmDataFileName,", dgmLayoutFileName: ",dgmLayoutFileName,", dgmQuickStyleFileName: ",dgmQuickStyleFileName);
    var dgmClr: any = zip.xmlFresh(dgmClrFileName);
    var dgmData: any = zip.xmlFresh(dgmDataFileName);
    var dgmLayout: any = zip.xmlFresh(dgmLayoutFileName);
    var dgmQuickStyle: any = zip.xmlFresh(dgmQuickStyleFileName);
    //console.log(dgmClr,dgmData,dgmLayout,dgmQuickStyle)
    ///get drawing#.xml
    // var dgmDrwFileName = "";
    // var dataModelExt = getTextByPathList(dgmData, ["dgm:dataModel", "dgm:extLst", "a:ext", "dsp:dataModelExt", "attrs"]);
    // if (dataModelExt !== undefined) {
    //     var dgmDrwFileId = dataModelExt["relId"];
    //     dgmDrwFileName = warpObj["slideResObj"][dgmDrwFileId]["target"];
    // }
    // var dgmDrwFile = "";
    // if (dgmDrwFileName != "") {
    //     dgmDrwFile = zip.xmlFresh(dgmDrwFileName);
    // }
    // var dgmDrwSpArray = getTextByPathList(dgmDrwFile, ["dsp:drawing", "dsp:spTree", "dsp:sp"]);
    //var dgmDrwSpArray = getTextByPathList(warpObj["digramFileContent"], ["dsp:drawing", "dsp:spTree", "dsp:sp"]);
    var dgmDrwSpArray: any = getTextByPathList(warpObj["digramFileContent"], ["p:drawing", "p:spTree", "p:sp"]);
    var rslt: any = "";
    if (dgmDrwSpArray !== undefined) {
        var dgmDrwSpArrayLen: any = dgmDrwSpArray.length;
        for (var i: any = 0; i < dgmDrwSpArrayLen; i++) {
            var dspSp: any = dgmDrwSpArray[i];
            // var dspSpObjToStr = JSON.stringify(dspSp);
            // var pSpStr = dspSpObjToStr.replace(/dsp:/g, "p:");
            // var pSpStrToObj = JSON.parse(pSpStr);
            //console.log("pSpStrToObj[" + i + "]: ", pSpStrToObj);
            //rslt += processSpNode(pSpStrToObj, node, warpObj, "diagramBg", sType)
            rslt += processSpNode(dspSp, node, warpObj, "diagramBg", sType)
        }
        // dgmDrwFile: "dsp:"-> "p:"
    }

    return "<div class='block diagram-content' style='" +
        getPosition(xfrmNode, node, undefined, undefined, sType) +
        getSize(xfrmNode, undefined, undefined) +
        "'>" + rslt + "</div>";
}

// ported from legacy/js/pptxjs.js:10385
export function getPosition(slideSpNode?: any, pNode?: any, slideLayoutSpNode?: any, slideMasterSpNode?: any, sType?: any): any {
    var off: any;
    var x: any = -1, y: any = -1;

    if (slideSpNode !== undefined) {
        off = slideSpNode["a:off"]["attrs"];
    }

    if (off === undefined && slideLayoutSpNode !== undefined) {
        off = slideLayoutSpNode["a:off"]["attrs"];
    } else if (off === undefined && slideMasterSpNode !== undefined) {
        off = slideMasterSpNode["a:off"]["attrs"];
    }
    var offX: any = 0, offY: any = 0;
    var grpX: any = 0, grpY: any = 0;
    if (sType == "group") {

        var grpXfrmNode: any = getTextByPathList(pNode, ["p:grpSpPr", "a:xfrm"]);
        if (xfrmNode !== undefined) {
            grpX = parseInt(grpXfrmNode["a:off"]["attrs"]["x"]) * slideFactor;
            grpY = parseInt(grpXfrmNode["a:off"]["attrs"]["y"]) * slideFactor;
            // var chx = parseInt(grpXfrmNode["a:chOff"]["attrs"]["x"]) * slideFactor;
            // var chy = parseInt(grpXfrmNode["a:chOff"]["attrs"]["y"]) * slideFactor;
            // var cx = parseInt(grpXfrmNode["a:ext"]["attrs"]["cx"]) * slideFactor;
            // var cy = parseInt(grpXfrmNode["a:ext"]["attrs"]["cy"]) * slideFactor;
            // var chcx = parseInt(grpXfrmNode["a:chExt"]["attrs"]["cx"]) * slideFactor;
            // var chcy = parseInt(grpXfrmNode["a:chExt"]["attrs"]["cy"]) * slideFactor;
            // var rotate = parseInt(grpXfrmNode["attrs"]["rot"])
        }
    }
    if (sType == "group-rotate" && pNode["p:grpSpPr"] !== undefined) {
        var xfrmNode: any = pNode["p:grpSpPr"]["a:xfrm"];
        // var ox = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * slideFactor;
        // var oy = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * slideFactor;
        var chx: any = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * slideFactor;
        var chy: any = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * slideFactor;

        offX = chx;
        offY = chy;
    }
    if (off === undefined) {
        return "";
    } else {
        x = parseInt(off["x"]) * slideFactor;
        y = parseInt(off["y"]) * slideFactor;
        // if (type = "body")
        //     console.log("getPosition: slideSpNode: ", slideSpNode, ", type: ", type, "x: ", x, "offX:", offX, "y:", y, "offY:", offY)
        return (isNaN(x) || isNaN(y)) ? "" : "top:" + (y - offY + grpY) + "px; left:" + (x - offX + grpX) + "px;";
    }

}

// ported from legacy/js/pptxjs.js:10437
export function getSize(slideSpNode?: any, slideLayoutSpNode?: any, slideMasterSpNode?: any): any {
    var ext: any = undefined;
    var w: any = -1, h: any = -1;

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
        w = parseInt(ext["cx"]) * slideFactor;
        h = parseInt(ext["cy"]) * slideFactor;
        return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
    }

}
