// Ported from PPTXjs v1 (legacy/js/pptxjs.js). Internal rendering code: the
// node objects come from the simplified XML tree, so most values are `any`.
import type { WarpObj } from "./context";
import { imageSizeFromDataUrl } from "../utils/image-size";
import { getSolidFill } from "./color";
import { getTextByPathList, setTextByPathList } from "../core/xml";
import { getMimeType } from "../core/package";
import { base64ArrayBuffer } from "../utils/base64";
import { escapeHtml } from "../utils/html";
import { angleToDegrees } from "./units";
import { tinycolor } from "../utils/tinycolor";
// ported from legacy/js/pptxjs.js:12142
export function getShapeFill(node?: any, pNode?: any, isSvgMode?: any, warpObj?: any, source?: any): any {

    // 1. presentationML
    // p:spPr/ [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
    // From slide
    //Fill Type:
    //console.log("getShapeFill ShapeFill: ", node, ", isSvgMode; ", isSvgMode)
    var fillType: any = getFillType(getTextByPathList(node, ["p:spPr"]));
    //var noFill = getTextByPathList(node, ["p:spPr", "a:noFill"]);
    var fillColor: any;
    if (fillType == "NO_FILL") {
        return isSvgMode ? "none" : "";//"background-color: initial;";
    } else if (fillType == "SOLID_FILL") {
        var shpFill: any = node["p:spPr"]["a:solidFill"];
        fillColor = getSolidFill(shpFill, undefined, undefined, warpObj);
    } else if (fillType == "GRADIENT_FILL") {
        var shpFill: any = node["p:spPr"]["a:gradFill"];
        fillColor = getGradientFill(shpFill, warpObj);
        //console.log("shpFill",shpFill,grndColor.color)
    } else if (fillType == "PATTERN_FILL") {
        var shpFill: any = node["p:spPr"]["a:pattFill"];
        fillColor = getPatternFill(shpFill, warpObj);
    } else if (fillType == "PIC_FILL") {
        var shpFill: any = node["p:spPr"]["a:blipFill"];
        fillColor = getPicFill(source, shpFill, warpObj);
    }
    //console.log("getShapeFill ShapeFill: ", node, ", isSvgMode; ", isSvgMode, ", fillType: ", fillType, ", fillColor: ", fillColor, ", source: ", source)


    // 2. drawingML namespace
    if (fillColor === undefined) {
        var clrName: any = getTextByPathList(node, ["p:style", "a:fillRef"]);
        var idx: any = parseInt(getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
        if (idx == 0 || idx == 1000) {
            //no fill
            return isSvgMode ? "none" : "";
        } else if (idx > 0 && idx < 1000) {
            // <a:fillStyleLst> fill
        } else if (idx > 1000) {
            //<a:bgFillStyleLst>
        }
        fillColor = getSolidFill(clrName, undefined, undefined, warpObj);
    }
    // 3. is group fill
    if (fillColor === undefined) {
        var grpFill: any = getTextByPathList(node, ["p:spPr", "a:grpFill"]);
        if (grpFill !== undefined) {
            //fillColor = getSolidFill(clrName, undefined, undefined, undefined, warpObj);
            //get parent fill style - TODO
            //console.log("ShapeFill: grpFill: ", grpFill, ", pNode: ", pNode)
            var grpShpFill: any = pNode["p:grpSpPr"];
            var spShpNode: any = { "p:spPr": grpShpFill }
            return getShapeFill(spShpNode, node, isSvgMode, warpObj, source);
        } else if (fillType == "NO_FILL") {
            return isSvgMode ? "none" : "";
        }
    }
    //console.log("ShapeFill: fillColor: ", fillColor, ", fillType; ", fillType)

    if (fillColor !== undefined) {
        if (fillType == "GRADIENT_FILL") {
            if (isSvgMode) {
                // console.log("GRADIENT_FILL color", fillColor.color[0])
                return fillColor;
            } else {
                var colorAry: any = fillColor.color;
                var rot: any = fillColor.rot;

                var bgcolor: any = "background: linear-gradient(" + rot + "deg,";
                for (var i: any = 0; i < colorAry.length; i++) {
                    if (i == colorAry.length - 1) {
                        bgcolor += "#" + colorAry[i] + ");";
                    } else {
                        bgcolor += "#" + colorAry[i] + ", ";
                    }

                }
                return bgcolor;
            }
        } else if (fillType == "PIC_FILL") {
            if (isSvgMode) {
                return fillColor;
            } else {

                return "background-image:url(" + fillColor + ");";
            }
        } else if (fillType == "PATTERN_FILL") {
            /////////////////////////////////////////////////////////////Need to check -----------TODO
            // if (isSvgMode) {
            //     var color = tinycolor(fillColor);
            //     fillColor = color.toRgbString();

            //     return fillColor;
            // } else {
            var bgPtrn: any = "", bgSize: any = "", bgPos: any = "";
            bgPtrn = fillColor[0];
            if (fillColor[1] !== null && fillColor[1] !== undefined && fillColor[1] != "") {
                bgSize = " background-size:" + fillColor[1] + ";";
            }
            if (fillColor[2] !== null && fillColor[2] !== undefined && fillColor[2] != "") {
                bgPos = " background-position:" + fillColor[2] + ";";
            }
            return "background: " + bgPtrn + ";" + bgSize + bgPos;
            //}
        } else {
            if (isSvgMode) {
                var color: any = tinycolor(fillColor);
                fillColor = color.toRgbString();

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
            return "background-color: inherit;";
        }

    }

}

// ported from legacy/js/pptxjs.js:12268
export function getFillType(node?: any): any {
    //Need to test/////////////////////////////////////////////
    //SOLID_FILL
    //PIC_FILL
    //GRADIENT_FILL
    //PATTERN_FILL
    //NO_FILL
    var fillType: any = "";
    if (node === undefined || node === null || typeof node !== "object") {
        return fillType; // issue #33: called with a missing node
    }
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
    if (node["a:grpFill"] !== undefined) {
        fillType = "GROUP_FILL";
    }


    return fillType;
}

// ported from legacy/js/pptxjs.js:12298
export function getGradientFill(node?: any, warpObj?: any): any {
    //console.log("getGradientFill: node", node)
    var gsLst: any = node["a:gsLst"]["a:gs"];
    //get start color
    var color_ary: any = [];
    var tint_ary: any = [];
    for (var i: any = 0; i < gsLst.length; i++) {
        var lo_tint: any;
        var lo_color: any = getSolidFill(gsLst[i], undefined, undefined, warpObj);
        //console.log("lo_color",lo_color)
        color_ary[i] = lo_color;
    }
    //get rot
    var lin: any = node["a:lin"];
    var rot: any = 0;
    if (lin !== undefined) {
        rot = angleToDegrees(lin["attrs"]["ang"]) + 90;
    }
    return {
        "color": color_ary,
        "rot": rot
    }
}

// ported from legacy/js/pptxjs.js:12321
export function getPicFill(type?: any, node?: any, warpObj?: any): any {
    //Need to test/////////////////////////////////////////////
    //rId
    //TODO - Image Properties - Tile, Stretch, or Display Portion of Image
    //(http://officeopenxml.com/drwPic-tile.php)
    var img: any;
    var rId: any = node["a:blip"]["attrs"]["r:embed"];
    var imgPath: any;
    //console.log("getPicFill(...) rId: ", rId, ", warpObj: ", warpObj, ", type: ", type)
    if (type == "slideBg" || type == "slide") {
        imgPath = getTextByPathList(warpObj, ["slideResObj", rId, "target"]);
    } else if (type == "slideLayoutBg") {
        imgPath = getTextByPathList(warpObj, ["layoutResObj", rId, "target"]);
    } else if (type == "slideMasterBg") {
        imgPath = getTextByPathList(warpObj, ["masterResObj", rId, "target"]);
    } else if (type == "themeBg") {
        imgPath = getTextByPathList(warpObj, ["themeResObj", rId, "target"]);
    } else if (type == "diagramBg") {
        imgPath = getTextByPathList(warpObj, ["diagramResObj", rId, "target"]);
    }
    if (imgPath === undefined) {
        return undefined;
    }
    img = getTextByPathList(warpObj, ["loaded-images", imgPath]); //, type, rId
    if (img === undefined) {
        imgPath = escapeHtml(imgPath);


        var imgExt: any = imgPath.split(".").pop();
        if (imgExt == "xml") {
            return undefined;
        }
        var imgArrayBuffer: any = warpObj["zip"].bytes(imgPath);
        if (imgArrayBuffer === null) return undefined;
        var imgMimeType: any = getMimeType(imgExt);
        img = "data:" + imgMimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer);
        //warpObj["loaded-images"][imgPath] = img; //"defaultTextStyle": defaultTextStyle,
        setTextByPathList(warpObj, ["loaded-images", imgPath], img); //, type, rId
    }
    return img;
}

// ported from legacy/js/pptxjs.js:12361
export function getPatternFill(node?: any, warpObj?: any): any {
    //https://developer.mozilla.org/en-US/docs/Web/CSS/CSS_Images/Using_CSS_gradients
    //https://cssgradient.io/blog/css-gradient-text/
    //https://css-tricks.com/background-patterns-simplified-by-conic-gradients/
    //https://stackoverflow.com/questions/6705250/how-to-get-a-pattern-into-a-written-text-via-css
    //https://stackoverflow.com/questions/14072142/striped-text-in-css
    //https://css-tricks.com/stripes-css/
    //https://yuanchuan.dev/gradient-shapes/
    var fgColor: any = "", bgColor: any = "", prst: any = "";
    var bgClr: any = node["a:bgClr"];
    var fgClr: any = node["a:fgClr"];
    prst = node["attrs"]["prst"];
    fgColor = getSolidFill(fgClr, undefined, undefined, warpObj);
    bgColor = getSolidFill(bgClr, undefined, undefined, warpObj);
    //var angl_ary = getAnglefromParst(prst);
    //var ptrClr = "repeating-linear-gradient(" + angl + "deg,  #" + bgColor + ",#" + fgColor + " 2px);"
    //linear-gradient(0deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black), 
    //linear-gradient(90deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black);
    var linear_gradient: any = getLinerGrandient(prst, bgColor, fgColor);
    //console.log("getPatternFill: node:", node, ", prst: ", prst, ", fgColor: ", fgColor, ", bgColor:", bgColor, ', linear_gradient: ', linear_gradient)
    return linear_gradient;
}

// ported from legacy/js/pptxjs.js:12384
export function getLinerGrandient(prst?: any, bgColor?: any, fgColor?: any): any {
    // dashDnDiag (Dashed Downward Diagonal)-V
    // dashHorz (Dashed Horizontal)-V
    // dashUpDiag(Dashed Upward DIagonal)-V
    // dashVert(Dashed Vertical)-V
    // diagBrick(Diagonal Brick)-V
    // divot(Divot)-VX
    // dkDnDiag(Dark Downward Diagonal)-V
    // dkHorz(Dark Horizontal)-V
    // dkUpDiag(Dark Upward Diagonal)-V
    // dkVert(Dark Vertical)-V
    // dotDmnd(Dotted Diamond)-VX
    // dotGrid(Dotted Grid)-V
    // horzBrick(Horizontal Brick)-V
    // lgCheck(Large Checker Board)-V
    // lgConfetti(Large Confetti)-V
    // lgGrid(Large Grid)-V
    // ltDnDiag(Light Downward Diagonal)-V
    // ltHorz(Light Horizontal)-V
    // ltUpDiag(Light Upward Diagonal)-V
    // ltVert(Light Vertical)-V
    // narHorz(Narrow Horizontal)-V
    // narVert(Narrow Vertical)-V
    // openDmnd(Open Diamond)-V
    // pct10(10 %)-V
    // pct20(20 %)-V
    // pct25(25 %)-V
    // pct30(30 %)-V
    // pct40(40 %)-V
    // pct5(5 %)-V
    // pct50(50 %)-V
    // pct60(60 %)-V
    // pct70(70 %)-V
    // pct75(75 %)-V
    // pct80(80 %)-V
    // pct90(90 %)-V
    // smCheck(Small Checker Board) -V
    // smConfetti(Small Confetti)-V
    // smGrid(Small Grid) -V
    // solidDmnd(Solid Diamond)-V
    // sphere(Sphere)-V
    // trellis(Trellis)-VX
    // wave(Wave)-V
    // wdDnDiag(Wide Downward Diagonal)-V
    // wdUpDiag(Wide Upward Diagonal)-V
    // weave(Weave)-V
    // zigZag(Zig Zag)-V
    // shingle(Shingle)-V
    // plaid(Plaid)-V
    // cross (Cross)
    // diagCross(Diagonal Cross)
    // dnDiag(Downward Diagonal)
    // horz(Horizontal)
    // upDiag(Upward Diagonal)
    // vert(Vertical)
    switch (prst) {
        case "smGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "4px 4px"];
            break
        case "dotGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "8px 8px"];
            break
        case "lgGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1.5px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1.5px)  #" + bgColor + ";", "8px 8px"];
            break
        case "wdUpDiag":
            //return ["repeating-linear-gradient(-45deg,  #" + bgColor + ", #" + bgColor + " 1px,#" + fgColor + " 5px);"];
            return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            // return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 1px),  #" + fgColor + " 50%, transparent calc(50% + 1px),  transparent 100%) " +
            //     "#" + bgColor + ";", "6px 6px"];
            break
        case "dkUpDiag":
            return ["repeating-linear-gradient(-45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
            break
        case "ltUpDiag":
            return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
            break
        case "wdDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            break
        case "dkDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
            break
        case "ltDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
            break
        case "dkHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
            break
        case "ltHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            break
        case "narHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
            break
        case "dkVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
            break
        case "ltVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            break
        case "narVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
            break
        case "lgCheck":
        case "smCheck":
            var size: any = "";
            var pos: any = "";
            if (prst == "lgCheck") {
                size = "8px 8px";
                pos = "0 0, 4px 4px, 4px 4px, 8px 8px";
            } else {
                size = "4px 4px";
                pos = "0 0, 2px 2px, 2px 2px, 4px 4px";
            }
            return ["linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0) " +
                "#" + bgColor + ";", size, pos];
            break
        // case "smCheck":
        //     return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
        //         "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%)  " +
        //         "#" + bgColor + ";", "4px 4px"];
        //     break 

        case "dashUpDiag":
            return ["repeating-linear-gradient(152deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "dashDnDiag":
            return ["repeating-linear-gradient(45deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "diagBrick":
            return ["linear-gradient(45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                "linear-gradient(-45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                "linear-gradient(-45deg, transparent 65%,  #" + fgColor + " 80%, transparent 0) " +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "horzBrick":
            return ["linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px) " +
                "#" + fgColor + ";", "4px 4px", "0 0.15px, 0.3px 2.5px, 2px 2.15px, 2.35px 0.4px"];
            break

        case "dashVert":
            return ["linear-gradient(0deg,  #" + bgColor + " 30%, transparent 30%)," +
                "linear-gradient(90deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "dashHorz":
            return ["linear-gradient(90deg,  #" + bgColor + " 30%, transparent 30%)," +
                "linear-gradient(0deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "solidDmnd":
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", "8px 8px"];
            break
        case "openDmnd":
            return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
                "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%) " +
                "#" + bgColor + ";", "8px 8px"];
            break

        case "dotDmnd":
            return ["radial-gradient(#" + fgColor + " 15%, transparent 0), " +
                "radial-gradient(#" + fgColor + " 15%, transparent 0) " +
                "#" + bgColor + ";", "4px 4px", "0 0, 2px 2px"];
            break
        case "zigZag":
        case "wave":
            var size: any = "";
            if (prst == "zigZag") size = "0";
            else size = "1px";
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "lgConfetti":
        case "smConfetti":
            var size: any = "";
            if (prst == "lgConfetti") size = "4px 4px";
            else size = "2px 2px";
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px, " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px , " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", size];
            break
        // case "weave":
        //     return ["linear-gradient(45deg,  #" + bgColor + " 5%, transparent 25%) 50px 0, " +
        //         "linear-gradient(135deg,  #" + bgColor + " 25%, transparent 25%) 50px 0, " +
        //         "linear-gradient(45deg,  #" + bgColor + " 25%, transparent 25%) " +
        //         "#" + fgColor + ";", "4px 4px"];
        //     //background: linear-gradient(45deg, #dca 12%, transparent 0, transparent 88%, #dca 0),
        //     //linear-gradient(135deg, transparent 37 %, #a85 0, #a85 63 %, transparent 0),
        //     //linear-gradient(45deg, transparent 37 %, #dca 0, #dca 63 %, transparent 0) #753;
        //     // background-size: 25px 25px;
        //     break;

        case "plaid":
            return ["linear-gradient(0deg, transparent, transparent 25%, #" + fgColor + "33 25%, #" + fgColor + "33 50%)," +
                "linear-gradient(90deg, transparent, transparent 25%, #" + fgColor + "66 25%, #" + fgColor + "66 50%) " +
                "#" + bgColor + ";", "4px 4px"];
            /**
                background-color: #6677dd;
                background-image: 
                repeating-linear-gradient(0deg, transparent, transparent 35px, rgba(255, 255, 255, 0.2) 35px, rgba(255, 255, 255, 0.2) 70px), 
                repeating-linear-gradient(90deg, transparent, transparent 35px, rgba(255,255,255,0.4) 35px, rgba(255,255,255,0.4) 70px);
             */
            break;
        case "sphere":
            return ["radial-gradient(#" + fgColor + " 50%, transparent 50%)," +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "weave":
        case "shingle":
            return ["linear-gradient(45deg, #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px), " +
                "linear-gradient(-45deg,  #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px) 0 4px, " +
                "#" + bgColor + ";", "4px 8px"];
            break
        //background:
        //linear-gradient(45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px),
        //linear-gradient(-45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px)0 4px;
        //background-color:#708090;
        //background-size: 4px 8px;
        case "pct5":
        case "pct10":
        case "pct20":
        case "pct25":
        case "pct30":
        case "pct40":
        case "pct50":
        case "pct60":
        case "pct70":
        case "pct75":
        case "pct80":
        case "pct90":
        //case "dotDmnd":
        case "trellis":
        case "divot":
            var px_pr_ary: any;
            switch (prst) {
                case "pct5":
                    px_pr_ary = ["0.3px", "10%", "2px 2px"];
                    break
                case "divot":
                    px_pr_ary = ["0.3px", "40%", "4px 4px"];
                    break
                case "pct10":
                    px_pr_ary = ["0.3px", "20%", "2px 2px"];
                    break
                case "pct20":
                    //case "dotDmnd":
                    px_pr_ary = ["0.2px", "40%", "2px 2px"];
                    break
                case "pct25":
                    px_pr_ary = ["0.2px", "50%", "2px 2px"];
                    break
                case "pct30":
                    px_pr_ary = ["0.5px", "50%", "2px 2px"];
                    break
                case "pct40":
                    px_pr_ary = ["0.5px", "70%", "2px 2px"];
                    break
                case "pct50":
                    px_pr_ary = ["0.09px", "90%", "2px 2px"];
                    break
                case "pct60":
                    px_pr_ary = ["0.3px", "90%", "2px 2px"];
                    break
                case "pct70":
                case "trellis":
                    px_pr_ary = ["0.5px", "95%", "2px 2px"];
                    break
                case "pct75":
                    px_pr_ary = ["0.65px", "100%", "2px 2px"];
                    break
                case "pct80":
                    px_pr_ary = ["0.85px", "100%", "2px 2px"];
                    break
                case "pct90":
                    px_pr_ary = ["1px", "100%", "2px 2px"];
                    break
            }
            return ["radial-gradient(#" + fgColor + " " + px_pr_ary[0] + ", transparent " + px_pr_ary[1] + ")," +
                "#" + bgColor + ";", px_pr_ary[2]];
            break
        default:
            return [0, 0];
    }
}

// ported from legacy/js/pptxjs.js:13574
export function getSvgGradient(w?: any, h?: any, angl?: any, color_arry?: any, shpId?: any): any {
    var stopsArray: any = getMiddleStops(color_arry - 2);

    var svgAngle: any = '',
        svgHeight: any = h,
        svgWidth: any = w,
        svg: any = '',
        xy_ary: any = SVGangle(angl, svgHeight, svgWidth),
        x1: any = xy_ary[0],
        y1: any = xy_ary[1],
        x2: any = xy_ary[2],
        y2: any = xy_ary[3];

    var sal: any = stopsArray.length,
        sr: any = sal < 20 ? 100 : 1000;
    svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"';
    svgAngle = '<linearGradient id="linGrd_' + shpId + '"' + svgAngle + '>\n';
    svg += svgAngle;

    for (var i: any = 0; i < sal; i++) {
        var tinClr: any = tinycolor("#" + color_arry[i]);
        var alpha: any = tinClr.getAlpha();
        //console.log("color: ", color_arry[i], ", rgba: ", tinClr.toHexString(), ", alpha: ", alpha)
        svg += '<stop offset="' + Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr + '" style="stop-color:' + tinClr.toHexString() + '; stop-opacity:' + (alpha) + ';"';
        svg += '/>\n'
    }

    svg += '</linearGradient>\n' + '';

    return svg
}

// ported from legacy/js/pptxjs.js:13605
export function getMiddleStops(s?: any): any {
    var sArry: any = ['0%', '100%'];
    if (s == 0) {
        return sArry;
    } else {
        var i: any = s;
        while (i--) {
            var middleStop: any = 100 - ((100 / (s + 1)) * (i + 1)), // AM: Ex - For 3 middle stops, progression will be 25%, 50%, and 75%, plus 0% and 100% at the ends.
                middleStopString: any = middleStop + "%";
            sArry.splice(-1, 0, middleStopString);
        } // AM: add into stopsArray before 100%
    }
    return sArry
}

// ported from legacy/js/pptxjs.js:13619
export function SVGangle(deg?: any, svgHeight?: any, svgWidth?: any): any {
    var w: any = parseFloat(svgWidth),
        h: any = parseFloat(svgHeight),
        ang: any = parseFloat(deg),
        o: any = 2,
        n: any = 2,
        wc: any = w / 2,
        hc: any = h / 2,
        tx1: any = 2,
        ty1: any = 2,
        tx2: any = 2,
        ty2: any = 2,
        k: any = (((ang % 360) + 360) % 360),
        j: any = (360 - k) * Math.PI / 180,
        i: any = Math.tan(j),
        l: any = hc - i * wc;

    if (k == 0) {
        tx1 = w,
            ty1 = hc,
            tx2 = 0,
            ty2 = hc
    } else if (k < 90) {
        n = w,
            o = 0
    } else if (k == 90) {
        tx1 = wc,
            ty1 = 0,
            tx2 = wc,
            ty2 = h
    } else if (k < 180) {
        n = 0,
            o = 0
    } else if (k == 180) {
        tx1 = 0,
            ty1 = hc,
            tx2 = w,
            ty2 = hc
    } else if (k < 270) {
        n = 0,
            o = h
    } else if (k == 270) {
        tx1 = wc,
            ty1 = h,
            tx2 = wc,
            ty2 = 0
    } else {
        n = w,
            o = h;
    }
    // AM: I could not quite figure out what m, n, and o are supposed to represent from the original code on visualcsstools.com.
    var m: any = o + (n / i),
        tx1: any = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1,
        ty1: any = ty1 == 2 ? i * tx1 + l : ty1,
        tx2: any = tx2 == 2 ? w - tx1 : tx2,
        ty2: any = ty2 == 2 ? h - ty1 : ty2,
        x1: any = Math.round(tx2 / w * 100 * 100) / 100,
        y1: any = Math.round(ty2 / h * 100 * 100) / 100,
        x2: any = Math.round(tx1 / w * 100 * 100) / 100,
        y2: any = Math.round(ty1 / h * 100 * 100) / 100;
    return [x1, y1, x2, y2];
}

// ported from legacy/js/pptxjs.js:13681
export function getSvgImagePattern(node?: any, fill?: any, shpId?: any, warpObj?: any): any {
    var pic_dim: any = getBase64ImageDimensions(fill);
    var width: any = pic_dim[0];
    var height: any = pic_dim[1];
    //console.log("getSvgImagePattern node:", node);
    var blipFillNode: any = node["p:spPr"]["a:blipFill"];
    var tileNode: any = getTextByPathList(blipFillNode, ["a:tile", "attrs"])
    if (tileNode !== undefined && tileNode["sx"] !== undefined) {
        var sx: any = (parseInt(tileNode["sx"]) / 100000) * width;
        var sy: any = (parseInt(tileNode["sy"]) / 100000) * height;
    }

    var blipNode: any = node["p:spPr"]["a:blipFill"]["a:blip"];
    var tialphaModFixNode: any = getTextByPathList(blipNode, ["a:alphaModFix", "attrs"])
    var imgOpacity: any = "";
    if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
        var amt: any = parseInt(tialphaModFixNode["amt"]) / 100000;
        var opacity: any = amt;
        var imgOpacity: any = "opacity='" + opacity + "'";

    }
    if (sx !== undefined && sx != 0) {
        var ptrn: any = '<pattern id="imgPtrn_' + shpId + '" x="0" y="0"  width="' + sx + '" height="' + sy + '" patternUnits="userSpaceOnUse">';
    } else {
        var ptrn: any = '<pattern id="imgPtrn_' + shpId + '"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
    }
    var duotoneNode: any = getTextByPathList(blipNode, ["a:duotone"])
    var fillterNode: any = "";
    var filterUrl: any = "";
    if (duotoneNode !== undefined) {
        //console.log("pic duotoneNode: ", duotoneNode)
        var clr_ary: any = [];
        Object.keys(duotoneNode).forEach(function (clr_type?: any) {
            //Object.keys(duotoneNode[clr_type]).forEach(function (clr) {
            //console.log("blip pic duotone clr: ", duotoneNode[clr_type][clr], clr)
            if (clr_type != "attrs") {
                var obj: any = {};
                obj[clr_type] = duotoneNode[clr_type];
                //console.log("blip pic duotone obj: ", obj)
                var hexClr: any = getSolidFill(obj, undefined, undefined, warpObj)
                //clr_ary.push();

                var color: any = tinycolor("#" + hexClr);
                clr_ary.push(color.toRgb()); // { r: 255, g: 0, b: 0, a: 1 }
            }
            // })
        })

        if (clr_ary.length == 2) {

            fillterNode = '<filter id="svg_image_duotone"> ' +
                '<feColorMatrix type="matrix" values=".33 .33 .33 0 0' +
                '.33 .33 .33 0 0' +
                '.33 .33 .33 0 0' +
                '0 0 0 1 0">' +
                '</feColorMatrix>' +
                '<feComponentTransfer color-interpolation-filters="sRGB">' +
                //clr_ary.forEach(function(clr){
                '<feFuncR type="table" tableValues="' + clr_ary[0].r / 255 + ' ' + clr_ary[1].r / 255 + '"></feFuncR>' +
                '<feFuncG type="table" tableValues="' + clr_ary[0].g / 255 + ' ' + clr_ary[1].g / 255 + '"></feFuncG>' +
                '<feFuncB type="table" tableValues="' + clr_ary[0].b / 255 + ' ' + clr_ary[1].b / 255 + '"></feFuncB>' +
                //});
                '</feComponentTransfer>' +
                ' </filter>';
        }

        filterUrl = 'filter="url(#svg_image_duotone)"';

        ptrn += fillterNode;
    }

    fill = escapeHtml(fill);
    if (sx !== undefined && sx != 0) {
        ptrn += '<image  xlink:href="' + fill + '" x="0" y="0" width="' + sx + '" height="' + sy + '" ' + imgOpacity + ' ' + filterUrl + '></image>';
    } else {
        ptrn += '<image  xlink:href="' + fill + '" preserveAspectRatio="none" width="1" height="1" ' + imgOpacity + ' ' + filterUrl + '></image>';
    }
    ptrn += '</pattern>';

    //console.log("getSvgImagePattern(...) pic_dim:", pic_dim, ", fillColor: ", fill, ", blipNode: ", blipNode, ",sx: ", sx, ", sy: ", sy, ", clr_ary: ", clr_ary, ", ptrn: ", ptrn)

    return ptrn;
}

// ported from legacy/js/pptxjs.js:13765
export function getBase64ImageDimensions(imgSrc?: any): any {
    // v1 busy-waited on `new Image()`; read the size from the image header instead
    return imageSizeFromDataUrl(String(imgSrc ?? ""));
}
