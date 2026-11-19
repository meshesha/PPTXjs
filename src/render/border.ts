// Ported from PPTXjs v1 (legacy/js/pptxjs.js). Internal rendering code: the
// node objects come from the simplified XML tree, so most values are `any`.
import type { WarpObj } from "./context";
import { getFillType, getGradientFill, getPatternFill } from "./fill";
import { getSolidFill } from "./color";
import { getTextByPathList } from "../core/xml";
// ported from legacy/js/pptxjs.js:11398
export function getBorder(node?: any, pNode?: any, isSvgMode?: any, bType?: any, warpObj?: any): any {
    //console.log("getBorder", node, pNode, isSvgMode, bType)
    var cssText: any, lineNode: any, subNodeTxt: any;

    if (bType == "shape") {
        cssText = "border: ";
        lineNode = node["p:spPr"]["a:ln"];
        //subNodeTxt = "p:spPr";
        //node["p:style"]["a:lnRef"] = 
    } else if (bType == "text") {
        cssText = "";
        lineNode = node["a:rPr"]["a:ln"];
        //subNodeTxt = "a:rPr";
    }

    //var is_noFill = getTextByPathList(node, ["p:spPr", "a:noFill"]);
    var is_noFill: any = getTextByPathList(lineNode, ["a:noFill"]);
    if (is_noFill !== undefined) {
        return "hidden";
    }

    //console.log("lineNode: ", lineNode)
    if (lineNode == undefined) {
        var lnRefNode: any = getTextByPathList(node, ["p:style", "a:lnRef"])
        if (lnRefNode !== undefined){
            var lnIdx: any = getTextByPathList(lnRefNode, ["attrs", "idx"]);
            //console.log("lnIdx:", lnIdx, "lnStyleLst:", warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) -1])
            lineNode = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) - 1];
        }
    }
    if (lineNode == undefined) {
        //is table
        cssText = "";
        lineNode = node
    }

    var borderColor: any;
    if (lineNode !== undefined) {
        // Border width: 1pt = 12700, default = 0.75pt
        var borderWidth: any = parseInt(getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
        if (isNaN(borderWidth) || borderWidth < 1) {
            cssText += (4/3) + "px ";//"1pt ";
            borderWidth = isNaN(borderWidth) ? 0.75 : borderWidth; // OOXML default line width is 0.75pt
        } else {
            cssText += borderWidth + "px ";// + "pt ";
        }
        // Border type
        var borderType: any = getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
        if (borderType === undefined) {
            borderType = getTextByPathList(lineNode, ["attrs", "cmpd"]);
        }
        var strokeDasharray: any = "0";
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
            case "dbl":
                cssText += "double";
                strokeDasharray = "0";
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
        var fillTyp: any = getFillType(lineNode);
        //console.log("getBorder:node : fillTyp", fillTyp)
        if (fillTyp == "NO_FILL") {
            borderColor = isSvgMode ? "none" : "";//"background-color: initial;";
        } else if (fillTyp == "SOLID_FILL") {
            borderColor = getSolidFill(lineNode["a:solidFill"], undefined, undefined, warpObj);
        } else if (fillTyp == "GRADIENT_FILL") {
            borderColor = getGradientFill(lineNode["a:gradFill"], warpObj);
            //console.log("shpFill",shpFill,grndColor.color)
        } else if (fillTyp == "PATTERN_FILL") {
            borderColor = getPatternFill(lineNode["a:pattFill"], warpObj);
        }

    }

    //console.log("getBorder:node : borderColor", borderColor)
    // 2. drawingML namespace
    if (borderColor === undefined) {
        //var schemeClrNode = getTextByPathList(node, ["p:style", "a:lnRef", "a:schemeClr"]);
        // if (schemeClrNode !== undefined) {
        //     var schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);
        //     var borderColor = getSchemeColorFromTheme(schemeClr, undefined, undefined);
        // }
        var lnRefNode: any = getTextByPathList(node, ["p:style", "a:lnRef"]);
        //console.log("getBorder: lnRef : ", lnRefNode)
        if (lnRefNode !== undefined) {
            borderColor = getSolidFill(lnRefNode, undefined, undefined, warpObj);
        }

        // if (borderColor !== undefined) {
        //     var shade = getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
        //     if (shade !== undefined) {
        //         shade = parseInt(shade) / 10000;
        //         var color = tinycolor("#" + borderColor);
        //         borderColor = color.darken(shade).toHex8();//.replace("#", "");
        //     }
        // }

    }

    //console.log("getBorder: borderColor : ", borderColor)
    if (borderColor === undefined) {
        if (isSvgMode) {
            borderColor = "none";
        } else {
            borderColor = "hidden";
        }
    } else {
        borderColor = "#" + borderColor; //wrong if not solid fill - TODO

    }
    cssText += " " + borderColor + " ";//wrong if not solid fill - TODO

    if (isSvgMode) {
        return { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
    } else {
        return cssText + ";";
    }
    // } else {
    //     if (isSvgMode) {
    //         return { "color": 'none', "width": '0', "type": 'none', "strokeDasharray": '0' };
    //     } else {
    //         return "hidden";
    //     }
    // }
}
