// Ported from PPTXjs v1 (legacy/js/pptxjs.js). Internal rendering code: the
// node objects come from the simplified XML tree, so most values are `any`.
import type { WarpObj } from "./context";
import { processNodesInSlide } from "./nodes";
import { getFillType, getPicFill } from "./fill";
import { getSolidFill } from "./color";
import { getTextByPathList } from "../core/xml";
import { angleToDegrees } from "./units";
// ported from legacy/js/pptxjs.js:11567
export function getBackground(warpObj?: any, slideSize?: any, index?: any): any {
    //var rslt = "";
    var slideContent: any = warpObj["slideContent"];
    var slideLayoutContent: any = warpObj["slideLayoutContent"];
    var slideMasterContent: any = warpObj["slideMasterContent"];

    var nodesSldLayout: any = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:spTree"]);
    var nodesSldMaster: any = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:spTree"]);
    // console.log("slideContent : ", slideContent)
    // console.log("slideLayoutContent : ", slideLayoutContent)
    // console.log("slideMasterContent : ", slideMasterContent)
    //console.log("warpObj : ", warpObj)
    var showMasterSp: any = getTextByPathList(slideLayoutContent, ["p:sldLayout", "attrs", "showMasterSp"]);
    // p:sld/@showMasterSp="0" hides the graphics of the layout AND the master (issue #38)
    var slideShowMasterSp: any = getTextByPathList(slideContent, ["p:sld", "attrs", "showMasterSp"]);
    var hideTemplateShapes: boolean = slideShowMasterSp == "0" || slideShowMasterSp == "false";
    //console.log("slideLayoutContent : ", slideLayoutContent, ", showMasterSp: ", showMasterSp)
    var bgColor: any = getSlideBackgroundFill(warpObj, index);
    var result: any = "<div class='slide-background-" + index + "' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
    var node_ph_type_ary: any = [];
    if (nodesSldLayout !== undefined && !hideTemplateShapes) {
        for (var nodeKey in nodesSldLayout) {
            if (nodesSldLayout[nodeKey].constructor === Array) {
                for (var i: any = 0; i < nodesSldLayout[nodeKey].length; i++) {
                    var ph_type: any = getTextByPathList(nodesSldLayout[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    // if (ph_type !== undefined && ph_type != "pic") {
                    //     node_ph_type_ary.push(ph_type);
                    // }
                    if (ph_type != "pic") {
                        result += processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], nodesSldLayout, warpObj, "slideLayoutBg"); //slideLayoutBg , slideMasterBg
                    }
                }
            } else {
                var ph_type: any = getTextByPathList(nodesSldLayout[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                // if (ph_type !== undefined && ph_type != "pic") {
                //     node_ph_type_ary.push(ph_type);
                // }
                if (ph_type != "pic") {
                    result += processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, warpObj, "slideLayoutBg"); //slideLayoutBg, slideMasterBg
                }
            }
        }
    }
    if (nodesSldMaster !== undefined && !hideTemplateShapes && (showMasterSp == "1" || showMasterSp === undefined)) {
        for (var nodeKey in nodesSldMaster) {
            if (nodesSldMaster[nodeKey].constructor === Array) {
                for (var i: any = 0; i < nodesSldMaster[nodeKey].length; i++) {
                    var ph_type: any = getTextByPathList(nodesSldMaster[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    //if (node_ph_type_ary.indexOf(ph_type) > -1) {
                    result += processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], nodesSldMaster, warpObj, "slideMasterBg"); //slideLayoutBg , slideMasterBg
                    //}
                }
            } else {
                var ph_type: any = getTextByPathList(nodesSldMaster[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                //if (node_ph_type_ary.indexOf(ph_type) > -1) {
                result += processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, warpObj, "slideMasterBg"); //slideLayoutBg, slideMasterBg
                //}
            }
        }
    }
    return result;

}

// ported from legacy/js/pptxjs.js:11627
export function getSlideBackgroundFill(warpObj?: any, index?: any): any {
    var slideContent: any = warpObj["slideContent"];
    var slideLayoutContent: any = warpObj["slideLayoutContent"];
    var slideMasterContent: any = warpObj["slideMasterContent"];

    //console.log("slideContent: ", slideContent)
    //console.log("slideLayoutContent: ", slideLayoutContent)
    //console.log("slideMasterContent: ", slideMasterContent)
    //getFillType(node)
    var bgPr: any = getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
    var bgRef: any = getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);
    //console.log("slideContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
    var bgcolor: any;
    if (bgPr !== undefined) {
        //bgcolor = "background-color: blue;";
        var bgFillTyp: any = getFillType(bgPr);

        if (bgFillTyp == "SOLID_FILL") {
            var sldFill: any = bgPr["a:solidFill"];
            var clrMapOvr: any;
            var sldClrMapOvr: any = getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                clrMapOvr = sldClrMapOvr;
            } else {
                var sldClrMapOvr: any = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                }

            }
            var sldBgClr: any = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
            //var sldTint = getColorOpacity(sldFill);
            //console.log("bgColor: ", bgColor)
            //bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
            bgcolor = "background: #" + sldBgClr + ";";

        } else if (bgFillTyp == "GRADIENT_FILL") {
            bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
        } else if (bgFillTyp == "PIC_FILL") {
            //console.log("PIC_FILL - ", bgFillTyp, bgPr, warpObj);
            bgcolor = getBgPicFill(bgPr, "slideBg", warpObj, undefined, index);

        }
        //console.log(slideContent,slideMasterContent,color_ary,tint_ary,rot,bgcolor)
    } else if (bgRef !== undefined) {
        //console.log("slideContent",bgRef)
        var clrMapOvr: any;
        var sldClrMapOvr: any = getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            clrMapOvr = sldClrMapOvr;
        } else {
            var sldClrMapOvr: any = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                clrMapOvr = sldClrMapOvr;
            } else {
                clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            }

        }
        var phClr: any = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);

        // if (bgRef["a:srgbClr"] !== undefined) {
        //     phClr = getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
        // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
        //     var schemeClr = getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);
        //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
        // }
        var idx: any = Number(bgRef["attrs"]["idx"]);


        if (idx == 0 || idx == 1000) {
            //no background
        } else if (idx > 0 && idx < 1000) {
            //fillStyleLst in themeContent
            //themeContent["a:fmtScheme"]["a:fillStyleLst"]
            //bgcolor = "background: red;";
        } else if (idx > 1000) {
            //bgFillStyleLst  in themeContent
            //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
            var trueIdx: any = idx - 1000;
            // themeContent["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
            var bgFillLst: any = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
            var sortblAry: any = [];
            Object.keys(bgFillLst).forEach(function (key?: any) {
                var bgFillLstTyp: any = bgFillLst[key];
                if (key != "attrs") {
                    if (bgFillLstTyp.constructor === Array) {
                        for (var i: any = 0; i < bgFillLstTyp.length; i++) {
                            var obj: any = {};
                            obj[key] = bgFillLstTyp[i];
                            obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                            obj["attrs"] = {
                                "order": bgFillLstTyp[i]["attrs"]["order"]
                            }
                            sortblAry.push(obj)
                        }
                    } else {
                        var obj: any = {};
                        obj[key] = bgFillLstTyp;
                        obj["idex"] = bgFillLstTyp["attrs"]["order"];
                        obj["attrs"] = {
                            "order": bgFillLstTyp["attrs"]["order"]
                        }
                        sortblAry.push(obj)
                    }
                }
            });
            var sortByOrder: any = sortblAry.slice(0);
            sortByOrder.sort(function (a?: any, b?: any) {
                return a.idex - b.idex;
            });
            var bgFillLstIdx: any = sortByOrder[trueIdx - 1];
            var bgFillTyp: any = getFillType(bgFillLstIdx);
            if (bgFillTyp == "SOLID_FILL") {
                var sldFill: any = bgFillLstIdx["a:solidFill"];
                var sldBgClr: any = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                //var sldTint = getColorOpacity(sldFill);
                //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                bgcolor = "background: #" + sldBgClr + ";";
                //console.log("slideMasterContent - sldFill",sldFill)
            } else if (bgFillTyp == "GRADIENT_FILL") {
                bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
            } else {
                /* debug log removed */
            }
        }

    }
    else {
        bgPr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr"]);
        bgRef = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgRef"]);
        //console.log("slideLayoutContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
        var clrMapOvr: any;
        var sldClrMapOvr: any = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            clrMapOvr = sldClrMapOvr;
        } else {
            clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
        }
        if (bgPr !== undefined) {
            var bgFillTyp: any = getFillType(bgPr);
            if (bgFillTyp == "SOLID_FILL") {
                var sldFill: any = bgPr["a:solidFill"];

                var sldBgClr: any = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                //var sldTint = getColorOpacity(sldFill);
                // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                bgcolor = "background: #" + sldBgClr + ";";
            } else if (bgFillTyp == "GRADIENT_FILL") {
                bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
            } else if (bgFillTyp == "PIC_FILL") {
                bgcolor = getBgPicFill(bgPr, "slideLayoutBg", warpObj, undefined, index);

            }
            //console.log("slideLayoutContent",bgcolor)
        } else if (bgRef !== undefined) {
            /* debug log removed */
            //bgcolor = "background: white;";
            var phClr: any = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
            var idx: any = Number(bgRef["attrs"]["idx"]);
            //console.log("phClr=", phClr, "idx=", idx)

            if (idx == 0 || idx == 1000) {
                //no background
            } else if (idx > 0 && idx < 1000) {
                //fillStyleLst in themeContent
                //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                //bgcolor = "background: red;";
            } else if (idx > 1000) {
                //bgFillStyleLst  in themeContent
                //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                var trueIdx: any = idx - 1000;
                var bgFillLst: any = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                var sortblAry: any = [];
                Object.keys(bgFillLst).forEach(function (key?: any) {
                    //console.log("cubicBezTo[" + key + "]:");
                    var bgFillLstTyp: any = bgFillLst[key];
                    if (key != "attrs") {
                        if (bgFillLstTyp.constructor === Array) {
                            for (var i: any = 0; i < bgFillLstTyp.length; i++) {
                                var obj: any = {};
                                obj[key] = bgFillLstTyp[i];
                                obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                obj["attrs"] = {
                                    "order": bgFillLstTyp[i]["attrs"]["order"]
                                }
                                sortblAry.push(obj)
                            }
                        } else {
                            var obj: any = {};
                            obj[key] = bgFillLstTyp;
                            obj["idex"] = bgFillLstTyp["attrs"]["order"];
                            obj["attrs"] = {
                                "order": bgFillLstTyp["attrs"]["order"]
                            }
                            sortblAry.push(obj)
                        }
                    }
                });
                var sortByOrder: any = sortblAry.slice(0);
                sortByOrder.sort(function (a?: any, b?: any) {
                    return a.idex - b.idex;
                });
                var bgFillLstIdx: any = sortByOrder[trueIdx - 1];
                var bgFillTyp: any = getFillType(bgFillLstIdx);
                if (bgFillTyp == "SOLID_FILL") {
                    var sldFill: any = bgFillLstIdx["a:solidFill"];
                    //console.log("sldFill: ", sldFill)
                    //var sldTint = getColorOpacity(sldFill);
                    //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                    var sldBgClr: any = getSolidFill(sldFill, clrMapOvr, phClr, warpObj);
                    //console.log("bgcolor: ", bgcolor)
                    bgcolor = "background: #" + sldBgClr + ";";
                } else if (bgFillTyp == "GRADIENT_FILL") {
                    //console.log("GRADIENT_FILL: ", bgFillLstIdx, phClr)
                    bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                } else if (bgFillTyp == "PIC_FILL") {
                    //theme rels
                    //console.log("PIC_FILL - ", bgFillTyp, bgFillLstIdx, bgFillLst, warpObj);
                    bgcolor = getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr, index);
                } else {
                    /* debug log removed */
                }
            }
        } else {
            bgPr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr"]);
            bgRef = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgRef"]);

            var clrMap: any = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            //console.log("slideMasterContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
            if (bgPr !== undefined) {
                var bgFillTyp: any = getFillType(bgPr);
                if (bgFillTyp == "SOLID_FILL") {
                    var sldFill: any = bgPr["a:solidFill"];
                    var sldBgClr: any = getSolidFill(sldFill, clrMap, undefined, warpObj);
                    // var sldTint = getColorOpacity(sldFill);
                    // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                    bgcolor = "background: #" + sldBgClr + ";";
                } else if (bgFillTyp == "GRADIENT_FILL") {
                    bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                } else if (bgFillTyp == "PIC_FILL") {
                    bgcolor = getBgPicFill(bgPr, "slideMasterBg", warpObj, undefined, index);
                }
            } else if (bgRef !== undefined) {
                //var obj={
                //    "a:solidFill": bgRef
                //}
                var phClr: any = getSolidFill(bgRef, clrMap, undefined, warpObj);
                // var phClr;
                // if (bgRef["a:srgbClr"] !== undefined) {
                //     phClr = getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
                // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                //     var schemeClr = getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);

                //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
                // }
                var idx: any = Number(bgRef["attrs"]["idx"]);
                //console.log("phClr=", phClr, "idx=", idx)

                if (idx == 0 || idx == 1000) {
                    //no background
                } else if (idx > 0 && idx < 1000) {
                    //fillStyleLst in themeContent
                    //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                    //bgcolor = "background: red;";
                } else if (idx > 1000) {
                    //bgFillStyleLst  in themeContent
                    //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                    var trueIdx: any = idx - 1000;
                    var bgFillLst: any = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    var sortblAry: any = [];
                    Object.keys(bgFillLst).forEach(function (key?: any) {
                        //console.log("cubicBezTo[" + key + "]:");
                        var bgFillLstTyp: any = bgFillLst[key];
                        if (key != "attrs") {
                            if (bgFillLstTyp.constructor === Array) {
                                for (var i: any = 0; i < bgFillLstTyp.length; i++) {
                                    var obj: any = {};
                                    obj[key] = bgFillLstTyp[i];
                                    obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp[i]["attrs"]["order"]
                                    }
                                    sortblAry.push(obj)
                                }
                            } else {
                                var obj: any = {};
                                obj[key] = bgFillLstTyp;
                                obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                obj["attrs"] = {
                                    "order": bgFillLstTyp["attrs"]["order"]
                                }
                                sortblAry.push(obj)
                            }
                        }
                    });
                    var sortByOrder: any = sortblAry.slice(0);
                    sortByOrder.sort(function (a?: any, b?: any) {
                        return a.idex - b.idex;
                    });
                    var bgFillLstIdx: any = sortByOrder[trueIdx - 1];
                    var bgFillTyp: any = getFillType(bgFillLstIdx);
                    //console.log("bgFillLstIdx: ", bgFillLstIdx, ", bgFillTyp: ", bgFillTyp, ", phClr: ", phClr);
                    if (bgFillTyp == "SOLID_FILL") {
                        var sldFill: any = bgFillLstIdx["a:solidFill"];
                        //console.log("sldFill: ", sldFill)
                        //var sldTint = getColorOpacity(sldFill);
                        //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                        var sldBgClr: any = getSolidFill(sldFill, clrMap, phClr, warpObj);
                        //console.log("bgcolor: ", bgcolor)
                        bgcolor = "background: #" + sldBgClr + ";";
                    } else if (bgFillTyp == "GRADIENT_FILL") {
                        //console.log("GRADIENT_FILL: ", bgFillLstIdx, phClr)
                        bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                    } else if (bgFillTyp == "PIC_FILL") {
                        //theme rels
                        // console.log("PIC_FILL - ", bgFillTyp, bgFillLstIdx, bgFillLst, warpObj);
                        bgcolor = getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr, index);
                    } else {
                        /* debug log removed */
                    }
                }
            }
        }
    }

    //console.log("bgcolor: ", bgcolor)
    return bgcolor;
}

// ported from legacy/js/pptxjs.js:11958
export function getBgGradientFill(bgPr?: any, phClr?: any, slideMasterContent?: any, warpObj?: any): any {
    var bgcolor: any = "";
    if (bgPr !== undefined) {
        var grdFill: any = bgPr["a:gradFill"];
        var gsLst: any = grdFill["a:gsLst"]["a:gs"];
        //var startColorNode, endColorNode;
        var color_ary: any = [];
        var pos_ary: any = [];
        //var tint_ary = [];
        for (var i: any = 0; i < gsLst.length; i++) {
            var lo_tint: any;
            var lo_color: any = "";
            var lo_color: any = getSolidFill(gsLst[i], slideMasterContent["p:sldMaster"]["p:clrMap"]["attrs"], phClr, warpObj);
            var pos: any = getTextByPathList(gsLst[i], ["attrs", "pos"])
            //console.log("pos: ", pos)
            if (pos !== undefined) {
                pos_ary[i] = pos / 1000 + "%";
            } else {
                pos_ary[i] = "";
            }
            //console.log("lo_color", lo_color)
            color_ary[i] = "#" + lo_color;
            //tint_ary[i] = (lo_tint !== undefined) ? parseInt(lo_tint) / 100000 : 1;
        }
        //get rot
        var lin: any = grdFill["a:lin"];
        var rot: any = 90;
        if (lin !== undefined) {
            rot = angleToDegrees(lin["attrs"]["ang"]);// + 270;
            //console.log("rot: ", rot)
            rot = rot + 90;
        }
        bgcolor = "background: linear-gradient(" + rot + "deg,";
        for (var i: any = 0; i < gsLst.length; i++) {
            if (i == gsLst.length - 1) {
                //if (phClr === undefined) {
                //bgcolor += "rgba(" + hexToRgbNew(color_ary[i]) + "," + tint_ary[i] + ")" + ");";
                bgcolor += color_ary[i] + " " + pos_ary[i] + ");";
                //} else {
                //bgcolor += "rgba(" + hexToRgbNew(phClr) + "," + tint_ary[i] + ")" + ");";
                // bgcolor += "" + phClr + ";";;
                //}
            } else {
                //if (phClr === undefined) {
                //bgcolor += "rgba(" + hexToRgbNew(color_ary[i]) + "," + tint_ary[i] + ")" + ", ";
                bgcolor += color_ary[i] + " " + pos_ary[i] + ", ";;
                //} else {
                //bgcolor += "rgba(" + hexToRgbNew(phClr) + "," + tint_ary[i] + ")" + ", ";
                // bgcolor += phClr + ", ";
                //}
            }
        }
    } else {
        if (phClr !== undefined) {
            //bgcolor = "rgba(" + hexToRgbNew(phClr) + ",0);";
            //bgcolor = phClr + ");";
            bgcolor = "background: #" + phClr + ";";
        }
    }
    return bgcolor;
}

// ported from legacy/js/pptxjs.js:12019
export function getBgPicFill(bgPr?: any, sorce?: any, warpObj?: any, phClr?: any, index?: any): any {
    //console.log("getBgPicFill bgPr", bgPr)
    var bgcolor: any;
    var picFillBase64: any = getPicFill(sorce, bgPr["a:blipFill"], warpObj);
    var ordr: any = bgPr["attrs"]["order"];
    var aBlipNode: any = bgPr["a:blipFill"]["a:blip"];
    //a:duotone
    var duotone: any = getTextByPathList(aBlipNode, ["a:duotone"]);
    if (duotone !== undefined) {
        //console.log("pic duotone: ", duotone)
        var clr_ary: any = [];
        // duotone.forEach(function (clr) {
        //     console.log("pic duotone clr: ", clr)
        // }) 
        Object.keys(duotone).forEach(function (clr_type?: any) {
            //console.log("pic duotone clr: clr_type: ", clr_type, duotone[clr_type])
            if (clr_type != "attrs") {
                var obj: any = {};
                obj[clr_type] = duotone[clr_type];
                clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
            }
            // Object.keys(duotone[clr_type]).forEach(function (clr) {
            //     if (clr != "order") {
            //         var obj = {};
            //         obj[clr_type] = duotone[clr_type][clr];
            //         clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
            //     }
            // })
        })
        //console.log("pic duotone clr_ary: ", clr_ary);
        //filter: url(file.svg#filter-element-id)
        //https://codepen.io/bhenbe/pen/QEZOvd
        //https://www.w3schools.com/cssref/css3_pr_filter.asp

        // var color1 = clr_ary[0];
        // var color2 = clr_ary[1];
        // var cssName = "";

        // var styleText_before_after = "content: '';" +
        //     "display: block;" +
        //     "width: 100%;" +
        //     "height: 100%;" +
        //     // "z-index: 1;" +
        //     "position: absolute;" +
        //     "top: 0;" +
        //     "left: 0;";

        // var cssName = "slide-background-" + index + "::before," + " .slide-background-" + index + "::after";
        // styleTable[styleText_before_after] = {
        //     "name": cssName,
        //     "text": styleText_before_after
        // };


        // var styleText_after = "background-color: #" + clr_ary[1] + ";" +
        //     "mix-blend-mode: darken;";

        // cssName = "slide-background-" + index + "::after";
        // styleTable[styleText_after] = {
        //     "name": cssName,
        //     "text": styleText_after
        // };

        // var styleText_before = "background-color: #" + clr_ary[0] + ";" +
        //     "mix-blend-mode: lighten;";

        // cssName = "slide-background-" + index + "::before";
        // styleTable[styleText_before] = {
        //     "name": cssName,
        //     "text": styleText_before
        // };

    }
    //a:alphaModFix
    var aphaModFixNode: any = getTextByPathList(aBlipNode, ["a:alphaModFix", "attrs"])
    var imgOpacity: any = "";
    if (aphaModFixNode !== undefined && aphaModFixNode["amt"] !== undefined && aphaModFixNode["amt"] != "") {
        var amt: any = parseInt(aphaModFixNode["amt"]) / 100000;
        //var opacity = amt;
        imgOpacity = "opacity:" + amt + ";";

    }
    //a:tile

    var tileNode: any = getTextByPathList(bgPr, ["a:blipFill", "a:tile", "attrs"])
    var prop_style: any = "";
    if (tileNode !== undefined && tileNode["sx"] !== undefined) {
        var sx: any = (parseInt(tileNode["sx"]) / 100000);
        var sy: any = (parseInt(tileNode["sy"]) / 100000);
        var tx: any = (parseInt(tileNode["tx"]) / 100000);
        var ty: any = (parseInt(tileNode["ty"]) / 100000);
        var algn: any = tileNode["algn"]; //tl(top left),t(top), tr(top right), l(left), ctr(center), r(right), bl(bottom left), b(bottm) , br(bottom right)
        var flip: any = tileNode["flip"]; //none,x,y ,xy

        prop_style += "background-repeat: round;"; //repeat|repeat-x|repeat-y|no-repeat|space|round|initial|inherit;
        //prop_style += "background-size: 300px 100px;"; size (w,h,sx, sy) -TODO
        //prop_style += "background-position: 50% 40%;"; //offset (tx, ty) -TODO
    }
    //a:srcRect
    //a:stretch => a:fillRect =>attrs (l:-17000, r:-17000)
    var stretch: any = getTextByPathList(bgPr, ["a:blipFill", "a:stretch"]);
    if (stretch !== undefined) {
        var fillRect: any = getTextByPathList(stretch, ["a:fillRect", "attrs"]);
        //console.log("getBgPicFill=>bgPr: ", bgPr)
        // var top = fillRect["t"], right = fillRect["r"], bottom = fillRect["b"], left = fillRect["l"];
        prop_style += "background-repeat: no-repeat;";
        prop_style += "background-position: center;";
        if (fillRect !== undefined) {
            //prop_style += "background-size: contain, cover;";
            prop_style += "background-size:  100% 100%;;";
        }
    }
    bgcolor = "background: url(" + picFillBase64 + ");  z-index: " + ordr + ";" + prop_style + imgOpacity;

    return bgcolor;
}
