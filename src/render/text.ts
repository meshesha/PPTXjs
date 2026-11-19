// Ported from PPTXjs v1 (legacy/js/pptxjs.js). Internal rendering code: the
// node objects come from the simplified XML tree, so most values are `any`.
import type { WarpObj } from "./context";
import { getBorder } from "./border";
import { getBgPicFill } from "./background";
import { getFillType, getGradientFill, getPatternFill } from "./fill";
import { getSchemeColorFromTheme, getSolidFill } from "./color";
import { decodeXmlEntities, getTextByPathList } from "../core/xml";
import { getMimeType } from "../core/package";
import { base64ArrayBuffer } from "../utils/base64";
import { escapeHtml, sanitizeUrl } from "../utils/html";
import { fontSizeFactor, slideFactor } from "./units";
import { hasDirectionPolicy, isRtlLanguage, resolveParagraphRtl } from "./bidi";
import { dingbatToUnicode } from "../dingbat";
import { applyNumericBullets } from "./numbering";

// v1 kept this as an implicit global shared between genTextBody and genSpanElement.
let is_first_br: any = false;
// ported from legacy/js/pptxjs.js:8569
export function genTextBody(textBodyNode?: any, spNode?: any, slideLayoutSpNode?: any, slideMasterSpNode?: any, type?: any, idx?: any, warpObj?: any, tbl_col_width?: any): any {
    var text: any = "";
    var slideMasterTextStyles: any = warpObj["slideMasterTextStyles"];

    if (textBodyNode === undefined) {
        return text;
    }
    //rtl : <p:txBody>
    //          <a:bodyPr wrap="square" rtlCol="1">

    var pFontStyle: any = getTextByPathList(spNode, ["p:style", "a:fontRef"]);
    //console.log("genTextBody spNode: ", getTextByPathList(spNode,["p:spPr","a:xfrm","a:ext"]));

    //var lstStyle = textBodyNode["a:lstStyle"];
    
    var apNode: any = textBodyNode["a:p"];
    if (apNode.constructor !== Array) {
        apNode = [apNode];
    }

    for (var i: any = 0; i < apNode.length; i++) {
        var pNode: any = apNode[i];
        var rNode: any = pNode["a:r"];
        var fldNode: any = pNode["a:fld"];
        var brNode: any = pNode["a:br"];
        if (rNode !== undefined) {
            rNode = (rNode.constructor === Array) ? rNode : [rNode];
        }
        if (rNode !== undefined && fldNode !== undefined) {
            fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode];
            rNode = rNode.concat(fldNode)
        }
        if (rNode !== undefined && brNode !== undefined) {
            is_first_br = true;
            brNode = (brNode.constructor === Array) ? brNode : [brNode];
            brNode.forEach(function (item?: any, indx?: any) {
                item.type = "br";
            });
            if (brNode.length > 1) {
                brNode.shift();
            }
            rNode = rNode.concat(brNode)
            //console.log("single a:p  rNode:", rNode, "brNode:", brNode )
            rNode.sort(function (a?: any, b?: any) {
                return a.attrs.order - b.attrs.order;
            });
            //console.log("sorted rNode:",rNode)
        }
        //rtlStr = "";//"dir='"+isRTL+"'";
        var styleText: any = "";
        var marginsVer: any = getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
        if (marginsVer != "") {
            styleText = marginsVer;
        }
        if (type == "body" || type == "obj" || type == "shape") {
            styleText += "font-size: 0px;";
            //styleText += "line-height: 0;";
            styleText += "font-weight: 100;";
            styleText += "font-style: normal;";
        }
        var cssName: any = "";

        if (styleText in warpObj.ctx.styleTable) {
            cssName = warpObj.ctx.styleTable[styleText]["name"];
        } else {
            cssName = "_css_" + (Object.keys(warpObj.ctx.styleTable).length + 1);
            warpObj.ctx.styleTable[styleText] = {
                "name": cssName,
                "text": styleText
            };
        }
        //console.log("textBodyNode: ", textBodyNode["a:lstStyle"])
        var prg_width_node: any = getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
        var prg_height_node: any;// = getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cy"]);
        var sld_prg_width: any = ((prg_width_node !== undefined) ? ("width:" + (parseInt(prg_width_node) * slideFactor) + "px;") : "width:inherit;");
        var sld_prg_height: any = ((prg_height_node !== undefined) ? ("height:" + (parseInt(prg_height_node) * slideFactor) + "px;") : "");
        var prg_dir: any = getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
        var prg_align: any = getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj);
        if (prg_align === undefined && prg_dir == "pregraph-rtl" && isPolicyRtl(pNode, idx, type, warpObj)) {
            // RTL decided by the textDirection option (not the file) and no alignment
            // anywhere: align to the start (right) like any RTL text
            prg_align = "h-right-rtl";
        }
        text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + "' class='slide-prgrph " + (prg_align ?? "") + " " +
            prg_dir + " " + cssName + "' >";
        var buText_ary: any = genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
        var isBullate: any = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
        var bu_width: any = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? buText_ary[1] + buText_ary[2] : 0;
        text += (buText_ary[0] !== undefined) ? buText_ary[0]:"";
        //get text margin 
        var margin_ary: any = getPregraphMargn(pNode, idx, type, isBullate, warpObj);
        var margin: any = margin_ary[0];
        var mrgin_val: any = margin_ary[1];
        if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0){
            //sorce : table text
            prg_width_node = tbl_col_width;
        }

        var prgrph_text: any = "";
        //var prgr_txt_art = [];
        var total_text_len: any = 0;
        if (rNode === undefined && pNode !== undefined) {
            // without r
            var prgr_text: any = genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj, isBullate);
            if (isBullate) {
                total_text_len += warpObj.ctx.measureHtmlWidth(prgr_text);
            }
            prgrph_text += prgr_text;
        } else if (rNode !== undefined) {
            // with multi r
            for (var j: any = 0; j < rNode.length; j++) {
                var prgr_text: any = genSpanElement(rNode[j], j, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj, isBullate);
                if (isBullate) {
                    total_text_len += warpObj.ctx.measureHtmlWidth(prgr_text);
                }
                prgrph_text += prgr_text;
            }
        }

        prg_width_node = parseInt(prg_width_node) * slideFactor - bu_width - mrgin_val;
        if (isBullate) {
            //get prg_width_node if there is a bulltes
            //console.log("total_text_len: ", total_text_len, "prg_width_node:", prg_width_node)

            // only shrink when real text measurement is available (browser DOM)
            if (warpObj.ctx.canMeasure && total_text_len < prg_width_node ){
                prg_width_node = total_text_len + bu_width + 1; // +1px: sub-pixel rounding must not force a wrap
            }
        }
        var prg_width: any = ((prg_width_node !== undefined) ? ("width:" + (prg_width_node )) + "px;" : "width:inherit;");
        // the text keeps the paragraph's base direction (v1 reset it to LTR with
        // "direction: initial", which misplaced punctuation/numbers in RTL text)
        text += "<div style='height: 100%;direction: inherit;overflow-wrap:break-word;word-wrap: break-word;" + prg_width + margin + "' >";
        text += prgrph_text;
        text += "</div>";
        text += "</div>";
    }

    return applyNumericBullets(text);
}

// ported from legacy/js/pptxjs.js:8709
export function genBuChar(node?: any, i?: any, spNode?: any, textBodyNode?: any, pFontStyle?: any, idx?: any, type?: any, warpObj?: any): any {
    //console.log("genBuChar node: ", node, ", spNode: ", spNode, ", pFontStyle: ", pFontStyle, "type", type)
    ///////////////////////////////////////Amir///////////////////////////////
    var sldMstrTxtStyles: any = warpObj["slideMasterTextStyles"];
    var lstStyle: any = textBodyNode["a:lstStyle"];

    var rNode: any = getTextByPathList(node, ["a:r"]);
    if (rNode !== undefined && rNode.constructor === Array) {
        rNode = rNode[0]; //bullet only to first "a:r"
    }
    var lvl: any = parseInt(getTextByPathList(node["a:pPr"], ["attrs", "lvl"])) + 1;
    if (isNaN(lvl)) {
        lvl = 1;
    }
    var lvlStr: any = "a:lvl" + lvl + "pPr";
    var dfltBultColor: any, dfltBultSize: any, bultColor: any, bultSize: any, color_tye: any;

    if (rNode !== undefined) {
        dfltBultColor = getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
        color_tye = dfltBultColor[2];
        dfltBultSize = getFontSize(rNode, textBodyNode, pFontStyle, lvl, type, warpObj);
    } else {
        return "";
    }
    //console.log("Bullet Size: " + bultSize);

    var bullet: any = "", marRStr: any = "", marLStr: any = "", margin_val: any=0, font_val: any=0;
    /////////////////////////////////////////////////////////////////


    var pPrNode: any = node["a:pPr"];
    var BullNONE: any = getTextByPathList(pPrNode, ["a:buNone"]);
    if (BullNONE !== undefined) {
        return "";
    }

    var buType: any = "TYPE_NONE";

    var layoutMasterNode: any = getLayoutAndMasterNode(node, idx, type, warpObj);
    var pPrNodeLaout: any = layoutMasterNode.nodeLaout;
    var pPrNodeMaster: any = layoutMasterNode.nodeMaster;

    var buChar: any = getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
    var buNum: any = getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
    var buStart: any = getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "startAt"]);
    var buPic: any = getTextByPathList(pPrNode, ["a:buBlip"]);
    if (buChar !== undefined) {
        buType = "TYPE_BULLET";
    }
    if (buNum !== undefined) {
        buType = "TYPE_NUMERIC";
    }
    if (buPic !== undefined) {
        buType = "TYPE_BULPIC";
    }

    var buFontSize: any = getTextByPathList(pPrNode, ["a:buSzPts", "attrs", "val"]);
    if (buFontSize === undefined) {
        buFontSize = getTextByPathList(pPrNode, ["a:buSzPct", "attrs", "val"]);
        if (buFontSize !== undefined) {
            var prcnt: any = parseInt(buFontSize) / 100000;
            //dfltBultSize = XXpt
            //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
            var dfltBultSizeNoPt: any = parseInt(dfltBultSize);
            bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
        }
    } else {
        bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
    }

    //get definde bullet COLOR
    var buClrNode: any = getTextByPathList(pPrNode, ["a:buClr"]);


    if (buChar === undefined && buNum === undefined && buPic === undefined) {

        if (lstStyle !== undefined) {
            BullNONE = getTextByPathList(lstStyle, [lvlStr,"a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }
            buType = "TYPE_NONE";
            buChar = getTextByPathList(lstStyle, [lvlStr,"a:buChar", "attrs", "char"]);
            buNum = getTextByPathList(lstStyle, [lvlStr,"a:buAutoNum", "attrs", "type"]);
            buPic = getTextByPathList(lstStyle, [lvlStr,"a:buBlip"]);
            if (buChar !== undefined) {
                buType = "TYPE_BULLET";
            }
            if (buNum !== undefined) {
                buType = "TYPE_NUMERIC";
            }
            if (buPic !== undefined) {
                buType = "TYPE_BULPIC";
            }
            if (buChar !== undefined || buNum !== undefined || buPic !== undefined) {
                pPrNode = lstStyle[lvlStr];
            }
        }
    }
    if (buChar === undefined && buNum === undefined && buPic === undefined) {
        //check in slidelayout and masterlayout - TODO
        if (pPrNodeLaout !== undefined) {
            BullNONE = getTextByPathList(pPrNodeLaout, ["a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }
            buType = "TYPE_NONE";
            buChar = getTextByPathList(pPrNodeLaout, ["a:buChar", "attrs", "char"]);
            buNum = getTextByPathList(pPrNodeLaout, ["a:buAutoNum", "attrs", "type"]);
            buPic = getTextByPathList(pPrNodeLaout, ["a:buBlip"]);
            if (buChar !== undefined) {
                buType = "TYPE_BULLET";
            }
            if (buNum !== undefined) {
                buType = "TYPE_NUMERIC";
            }
            if (buPic !== undefined) {
                buType = "TYPE_BULPIC";
            }
        }
        if (buChar === undefined && buNum === undefined && buPic === undefined) {
            //masterlayout

            if (pPrNodeMaster !== undefined) {
                BullNONE = getTextByPathList(pPrNodeMaster, ["a:buNone"]);
                if (BullNONE !== undefined) {
                    return "";
                }
                buType = "TYPE_NONE";
                buChar = getTextByPathList(pPrNodeMaster, ["a:buChar", "attrs", "char"]);
                buNum = getTextByPathList(pPrNodeMaster, ["a:buAutoNum", "attrs", "type"]);
                buPic = getTextByPathList(pPrNodeMaster, ["a:buBlip"]);
                if (buChar !== undefined) {
                    buType = "TYPE_BULLET";
                }
                if (buNum !== undefined) {
                    buType = "TYPE_NUMERIC";
                }
                if (buPic !== undefined) {
                    buType = "TYPE_BULPIC";
                }
            }

        }

    }
    //rtl
    if (buStart === undefined) {
        buStart = getTextByPathList(lstStyle, [lvlStr, "a:buAutoNum", "attrs", "startAt"]) ??
            getTextByPathList(pPrNodeLaout, ["a:buAutoNum", "attrs", "startAt"]) ??
            getTextByPathList(pPrNodeMaster, ["a:buAutoNum", "attrs", "startAt"]);
    }
    var getRtlVal: any = getTextByPathList(pPrNode, ["attrs", "rtl"]);
    if (getRtlVal === undefined) {
        getRtlVal = getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (getRtlVal === undefined && type != "shape") {
            getRtlVal = getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    var isRTL: any = resolveParagraphRtl(node, getRtlVal, warpObj);
    //align
    var alignNode: any = getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
    if (alignNode === undefined) {
        alignNode = getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
        if (alignNode === undefined) {
            alignNode = getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
        }
    }
    //indent?
    var indentNode: any = getTextByPathList(pPrNode, ["attrs", "indent"]);
    if (indentNode === undefined) {
        indentNode = getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
        if (indentNode === undefined) {
            indentNode = getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
        }
    }
    var indent: any = 0;
    if (indentNode !== undefined) {
        indent = parseInt(indentNode) * slideFactor;
    }
    //marL
    var marLNode: any = getTextByPathList(pPrNode, ["attrs", "marL"]);
    if (marLNode === undefined) {
        marLNode = getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
        if (marLNode === undefined) {
            marLNode = getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
        }
    }
    //console.log("genBuChar() isRTL", isRTL, "alignNode:", alignNode)
    if (marLNode !== undefined) {
        var marginLeft: any = parseInt(marLNode) * slideFactor;
        if (isRTL) {// && alignNode == "r") {
            marLStr = "padding-right:";// "margin-right: ";
        } else {
            marLStr = "padding-left:";//"margin-left: ";
        }
        margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
        marLStr += margin_val + "px;";
    }
    
    //marR?
    var marRNode: any = getTextByPathList(pPrNode, ["attrs", "marR"]);
    if (marRNode === undefined && marLNode === undefined) {
        //need to check if this posble - TODO
        marRNode = getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
        if (marRNode === undefined) {
            marRNode = getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
        }
    }
    if (marRNode !== undefined) {
        var marginRight: any = parseInt(marRNode) * slideFactor;
        if (isRTL) {// && alignNode == "r") {
            marLStr = "padding-right:";// "margin-right: ";
        } else {
            marLStr = "padding-left:";//"margin-left: ";
        }
        marRStr += ((marginRight + indent < 0) ? 0 : (marginRight + indent)) + "px;";
    }

    if (buType != "TYPE_NONE") {
        //var buFontAttrs = getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
    }
    //console.log("Bullet Type: " + buType);
    //console.log("NumericTypr: " + buNum);
    //console.log("buChar: " + (buChar === undefined?'':buChar.charCodeAt(0)));
    //get definde bullet COLOR
    if (buClrNode === undefined){
        //lstStyle
        buClrNode = getTextByPathList(lstStyle, [lvlStr, "a:buClr"]);
    }
    if (buClrNode === undefined) {
        buClrNode = getTextByPathList(pPrNodeLaout, ["a:buClr"]);
        if (buClrNode === undefined) {
            buClrNode = getTextByPathList(pPrNodeMaster, ["a:buClr"]);
        }
    }
    var defBultColor: any;
    if (buClrNode !== undefined) {
        defBultColor = getSolidFill(buClrNode, undefined, undefined, warpObj);
    } else {
        if (pFontStyle !== undefined) {
            //console.log("genBuChar pFontStyle: ", pFontStyle)
            defBultColor = getSolidFill(pFontStyle, undefined, undefined, warpObj);
        }
    }
    if (defBultColor === undefined || defBultColor == "NONE") {
        bultColor = dfltBultColor;
    } else {
        bultColor = [defBultColor, "", "solid"];
        color_tye = "solid";
    }
    //console.log("genBuChar node:", node, "pPrNode", pPrNode, " buClrNode: ", buClrNode, "defBultColor:", defBultColor,"dfltBultColor:" , dfltBultColor , "bultColor:", bultColor)

    //console.log("genBuChar: buClrNode: ", buClrNode, "bultColor", bultColor)
    //get definde bullet SIZE
    if (buFontSize === undefined) {
        buFontSize = getTextByPathList(pPrNodeLaout, ["a:buSzPts", "attrs", "val"]);
        if (buFontSize === undefined) {
            buFontSize = getTextByPathList(pPrNodeLaout, ["a:buSzPct", "attrs", "val"]);
            if (buFontSize !== undefined) {
                var prcnt: any = parseInt(buFontSize) / 100000;
                //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                var dfltBultSizeNoPt: any = parseInt(dfltBultSize);
                bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
            }
        }else{
            bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
        }
    }
    if (buFontSize === undefined) {
        buFontSize = getTextByPathList(pPrNodeMaster, ["a:buSzPts", "attrs", "val"]);
        if (buFontSize === undefined) {
            buFontSize = getTextByPathList(pPrNodeMaster, ["a:buSzPct", "attrs", "val"]);
            if (buFontSize !== undefined) {
                var prcnt: any = parseInt(buFontSize) / 100000;
                //dfltBultSize = XXpt
                //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                var dfltBultSizeNoPt: any = parseInt(dfltBultSize);
                bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
            }
        } else {
            bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
        }
    }
    if (buFontSize === undefined) {
        bultSize = dfltBultSize;
    }
    font_val = parseInt(bultSize);
    ////////////////////////////////////////////////////////////////////////
    if (buType == "TYPE_BULLET") {
        // the bullet font is inherited independently of the bullet char (issue #40)
        var typefaceNode: any = getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]) ??
            getTextByPathList(lstStyle, [lvlStr, "a:buFont", "attrs", "typeface"]) ??
            getTextByPathList(pPrNodeLaout, ["a:buFont", "attrs", "typeface"]) ??
            getTextByPathList(pPrNodeMaster, ["a:buFont", "attrs", "typeface"]);
        var typeface: any = "";
        if (typefaceNode !== undefined) {
            typeface = "font-family: " + typefaceNode;
        }
        // var marginLeft = parseInt(getTextByPathList(marLNode)) * slideFactor;
        // var marginRight = parseInt(getTextByPathList(marRNode)) * slideFactor;
        // if (isNaN(marginLeft)) {
        //     marginLeft = 328600 * slideFactor;
        // }
        // if (isNaN(marginRight)) {
        //     marginRight = 0;
        // }

        bullet = "<div style='height: 100%;" + typeface + ";" +
            marLStr + marRStr +
            "font-size:" + bultSize + ";" ;
        
        //bullet += "display: table-cell;";
        //"line-height: 0px;";
        if (color_tye == "solid") {
            if (bultColor[0] !== undefined && bultColor[0] != "") {
                bullet += "color:#" + bultColor[0] + "; ";
            }
            if (bultColor[1] !== undefined && bultColor[1] != "" && bultColor[1] != ";") {
                bullet += "text-shadow:" + bultColor[1] + ";";
            }
            //no highlight/background-color to bullet
            // if (bultColor[3] !== undefined && bultColor[3] != "") {
            //     styleText += "background-color: #" + bultColor[3] + ";";
            // }
        } else if (color_tye == "pattern" || color_tye == "pic" || color_tye == "gradient") {
            if (color_tye == "pattern") {
                bullet += "background:" + bultColor[0][0] + ";";
                if (bultColor[0][1] !== null && bultColor[0][1] !== undefined && bultColor[0][1] != "") {
                    bullet += "background-size:" + bultColor[0][1] + ";";//" 2px 2px;" +
                }
                if (bultColor[0][2] !== null && bultColor[0][2] !== undefined && bultColor[0][2] != "") {
                    bullet += "background-position:" + bultColor[0][2] + ";";//" 2px 2px;" +
                }
                // bullet += "-webkit-background-clip: text;" +
                //     "background-clip: text;" +
                //     "color: transparent;" +
                //     "-webkit-text-stroke: " + bultColor[1].border + ";" +
                //     "filter: " + bultColor[1].effcts + ";";
            } else if (color_tye == "pic") {
                bullet += bultColor[0] + ";";
                // bullet += "-webkit-background-clip: text;" +
                //     "background-clip: text;" +
                //     "color: transparent;" +
                //     "-webkit-text-stroke: " + bultColor[1].border + ";";

            } else if (color_tye == "gradient") {

                var colorAry: any = bultColor[0].color;
                var rot: any = bultColor[0].rot;

                bullet += "background: linear-gradient(" + rot + "deg,";
                for (var i: any = 0; i < colorAry.length; i++) {
                    if (i == colorAry.length - 1) {
                        bullet += "#" + colorAry[i] + ");";
                    } else {
                        bullet += "#" + colorAry[i] + ", ";
                    }
                }
                // bullet += "color: transparent;" +
                //     "-webkit-background-clip: text;" +
                //     "background-clip: text;" +
                //     "-webkit-text-stroke: " + bultColor[1].border + ";";
            }
            bullet += "-webkit-background-clip: text;" +
                "background-clip: text;" +
                "color: transparent;";
            if (bultColor[1].border !== undefined && bultColor[1].border !== "") {
                bullet += "-webkit-text-stroke: " + bultColor[1].border + ";";
            }
            if (bultColor[1].effcts !== undefined && bultColor[1].effcts !== "") {
                bullet += "filter: " + bultColor[1].effcts + ";";
            }
        }

        if (isRTL) {
            //bullet += "display: inline-block;white-space: nowrap ;direction:rtl"; // float: right;  
            bullet += "white-space: nowrap ;direction:rtl"; // display: table-cell;;
        }
        var isIE11: any = false; // IE11 is no longer supported
        var htmlBu: any = buChar;

        if (!isIE11) {
            //ie11 does not support unicode ?
            htmlBu = getHtmlBullet(typefaceNode, buChar);
        }
        bullet += "'><div style='line-height: " + (font_val/2) + "px;'>" + htmlBu + "</div></div>"; //font_val
        //} 
        // else {
        //     marginLeft = 328600 * slideFactor * lvl;

        //     bullet = "<div style='" + marLStr + "'>" + buChar + "</div>";
        // }
    } else if (buType == "TYPE_NUMERIC") { ///////////Amir///////////////////////////////
        //if (buFontAttrs !== undefined) {
        // var marginLeft = parseInt(getTextByPathList(pPrNode, ["attrs", "marL"])) * slideFactor;
        // var marginRight = parseInt(buFontAttrs["pitchFamily"]);

        // if (isNaN(marginLeft)) {
        //     marginLeft = 328600 * slideFactor;
        // }
        // if (isNaN(marginRight)) {
        //     marginRight = 0;
        // }
        //var typeface = buFontAttrs["typeface"];

        bullet = "<div style='height: 100%;" + marLStr + marRStr +
            "color:#" + bultColor[0] + ";" +
            "font-size:" + bultSize + ";";// +
        //"line-height: 0px;";
        if (isRTL) {
            bullet += "display: inline-block;white-space: nowrap ;direction:rtl;"; // float: right;
        } else {
            bullet += "display: inline-block;white-space: nowrap ;direction:ltr;"; //float: left;
        }
        bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "'" + (buStart !== undefined ? " data-bulltstart = '" + parseInt(buStart) + "'" : "") + " class='numeric-bullet-style'></div>";
        // } else {
        //     marginLeft = 328600 * slideFactor * lvl;
        //     bullet = "<div style='margin-left: " + marginLeft + "px;";
        //     if (isRTL) {
        //         bullet += " float: right; direction:rtl;";
        //     } else {
        //         bullet += " float: left; direction:ltr;";
        //     }
        //     bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></div>";
        // }

    } else if (buType == "TYPE_BULPIC") { //PIC BULLET
        // var marginLeft = parseInt(getTextByPathList(pPrNode, ["attrs", "marL"])) * slideFactor;
        // var marginRight = parseInt(getTextByPathList(pPrNode, ["attrs", "marR"])) * slideFactor;

        // if (isNaN(marginRight)) {
        //     marginRight = 0;
        // }
        // //console.log("marginRight: "+marginRight)
        // //buPic
        // if (isNaN(marginLeft)) {
        //     marginLeft = 328600 * slideFactor;
        // } else {
        //     marginLeft = 0;
        // }
        //var buPicId = getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
        var buPicId: any = getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
        var svgPicPath: any = "";
        var buImg: any;
        if (buPicId !== undefined) {
            //svgPicPath = warpObj["slideResObj"][buPicId]["target"];
            //buImg = warpObj["zip"].file(svgPicPath).asText();
            //}else{
            //buPicId = getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
            var imgPath: any = warpObj["slideResObj"][buPicId]["target"];
            //console.log("imgPath: ", imgPath);
            var imgArrayBuffer: any = warpObj["zip"].bytes(imgPath);
            var imgExt: any = imgPath.split(".").pop();
            var imgMimeType: any = getMimeType(imgExt);
            buImg = "<img src='data:" + imgMimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%;'/>"// height: 100%
            //console.log("imgPath: "+imgPath+"\nimgMimeType: "+imgMimeType)
        }
        if (buPicId === undefined) {
            buImg = "&#8227;";
        }
        bullet = "<div style='height: 100%;" + marLStr + marRStr +
            "width:" + bultSize + ";display: inline-block; ";// +
        //"line-height: 0px;";
        if (isRTL) {
            bullet += "display: inline-block;white-space: nowrap ;direction:rtl;"; //direction:rtl; float: right;
        }
        bullet += "'>" + buImg + "  </div>";
        //////////////////////////////////////////////////////////////////////////////////////
    }
    // else {
    //     bullet = "<div style='margin-left: " + 328600 * slideFactor * lvl + "px" +
    //         "; margin-right: " + 0 + "px;'></div>";
    // }
    //console.log("genBuChar: width: ", $(bullet).outerWidth())
    return [bullet, margin_val, font_val];//$(bullet).outerWidth()];
}

// ported from legacy/js/pptxjs.js:9179
export function getHtmlBullet(typefaceNode?: any, buChar?: any): any {
    // http://www.alanwood.net/demos/wingdings.html
    var ch: string = decodeXmlEntities(String(buChar ?? ""));
    var font: string = decodeXmlEntities(String(typefaceNode ?? ""));
    var code: number = ch.codePointAt(0) ?? 0x2022;
    // Symbol fonts store glyphs in the private-use area U+F020..U+F0FF (issue #40)
    if (code >= 0xf020 && code <= 0xf0ff) code -= 0xf000;
    if (font === "Symbol") {
        var sym: number | undefined = SYMBOL_FONT[code];
        return "&#" + (sym ?? code) + ";";
    }
    if (font === "Wingdings" || font === "Wingdings 2" || font === "Wingdings 3" || font === "Webdings") {
        switch (code) {
            case 0xa7: return "&#9632;"; // ■ black square
            case 0x71: return "&#10065;"; // ❑
            case 0x76: return "&#10070;"; // ❖
            case 0xd8: return "&#11162;"; // ⮚
            case 0xfc: return "&#10004;"; // ✔
        }
        var wingCharCode: any = dingbatToUnicode(font, String.fromCharCode(code));
        if (wingCharCode !== null) return "&#" + wingCharCode + ";";
    }
    return "&#" + code + ";";
}

/** Symbol font code points that differ from Latin-1, mapped to Unicode. */
const SYMBOL_FONT: Record<number, number> = {
    0x22: 0x2200, 0x24: 0x2203, 0x27: 0x220b, 0x2a: 0x2217, 0x2d: 0x2212, 0x40: 0x2245,
    0xa5: 0x221e, 0xa7: 0x2663, 0xa8: 0x2666, 0xa9: 0x2665, 0xaa: 0x2660, 0xab: 0x2194,
    0xac: 0x2190, 0xad: 0x2191, 0xae: 0x2192, 0xaf: 0x2193, 0xb7: 0x2022, 0xb4: 0x00d7,
    0xb8: 0x00f7, 0xb9: 0x2260, 0xba: 0x2261, 0xbb: 0x2248, 0xbc: 0x2026, 0xc5: 0x2295,
    0xd6: 0x221a, 0xd7: 0x22c5, 0xd8: 0x00ac, 0xd9: 0x2227, 0xda: 0x2228, 0xdb: 0x21d4,
    0xdc: 0x21d0, 0xdd: 0x21d1, 0xde: 0x21d2, 0xdf: 0x21d3, 0xe0: 0x25ca, 0xf0: 0x2022,
};

// ported from legacy/js/pptxjs.js:9209
export function getDingbatToUnicode(typefaceNode?: any, buChar?: any): any {
    return dingbatToUnicode(String(typefaceNode), String(buChar));
}

// ported from legacy/js/pptxjs.js:9228
export function getLayoutAndMasterNode(node?: any, idx?: any, type?: any, warpObj?: any): any {
    var pPrNodeLaout: any, pPrNodeMaster: any;
    var pPrNode: any = node["a:pPr"];
    //lvl
    var lvl: any = 1;
    var lvlNode: any = getTextByPathList(pPrNode, ["attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }
    if (idx !== undefined) {
        //slidelayout
        pPrNodeLaout = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", "a:lvl" + lvl + "pPr"]);
        if (pPrNodeLaout === undefined) {
            pPrNodeLaout = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr"]);
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
            }
        }
    }
    if (type !== undefined) {
        //slidelayout
        var lvlStr: any = "a:lvl" + lvl + "pPr";
        if (pPrNodeLaout === undefined) {
            pPrNodeLaout = getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
        }
        //masterlayout
        if (type == "title" || type == "ctrTitle") {
            pPrNodeMaster = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr]);
        } else if (type == "body" || type == "obj" || type == "subTitle") {
            pPrNodeMaster = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr]);
        } else if (type == "shape" || type == "diagram") {
            pPrNodeMaster = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr]);
        } else if (type == "textBox") {
            pPrNodeMaster = getTextByPathList(warpObj, ["defaultTextStyle", lvlStr]);
        } else {
            pPrNodeMaster = getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
        }
    }
    return {
        "nodeLaout": pPrNodeLaout,
        "nodeMaster": pPrNodeMaster
    };
}

// ported from legacy/js/pptxjs.js:9271
export function genSpanElement(node?: any, rIndex?: any, pNode?: any, textBodyNode?: any, pFontStyle?: any, slideLayoutSpNode?: any, idx?: any, type?: any, rNodeLength?: any, warpObj?: any, isBullate?: any): any {
    //https://codepen.io/imdunn/pen/GRgwaye ?
    var text_style: any = "";
    var lstStyle: any = textBodyNode["a:lstStyle"];
    var slideMasterTextStyles: any = warpObj["slideMasterTextStyles"];

    var text: any = node["a:t"];
    //var text_count = text.length;

    var openElemnt: any = "<span";//"<bdi";
    var closeElemnt: any = "</span>";// "</bdi>";
    var styleText: any = "";
    if (text === undefined && node["type"] !== undefined) {
        if (is_first_br) {
            //openElemnt = "<br";
            //closeElemnt = "";
            //return "<br style='font-size: initial'>"
            is_first_br = false;
            return "<span class='line-break-br' ></span>";
        } else {
            // styleText += "display: block;";
            // openElemnt = "<span";
            // closeElemnt = "</span>";
        }

        styleText += "display: block;";
        //openElemnt = "<span";
        //closeElemnt = "</span>";
    } else {

        is_first_br = true;
    }
    if (typeof text !== 'string') {
        text = getTextByPathList(node, ["a:fld", "a:t"]);
        if (typeof text !== 'string') {
            text = "&nbsp;";
            //return "<span class='text-block '>&nbsp;</span>";
        }
        // if (text === undefined) {
        //     return "";
        // }
    }

    var pPrNode: any = pNode["a:pPr"];
    //lvl
    var lvl: any = 1;
    var lvlNode: any = getTextByPathList(pPrNode, ["attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }
    //console.log("genSpanElement node: ", node, "rIndex: ", rIndex, ", pNode: ", pNode, ",pPrNode: ", pPrNode, "pFontStyle:", pFontStyle, ", idx: ", idx, "type:", type, warpObj);
    var layoutMasterNode: any = getLayoutAndMasterNode(pNode, idx, type, warpObj);
    var pPrNodeLaout: any = layoutMasterNode.nodeLaout;
    var pPrNodeMaster: any = layoutMasterNode.nodeMaster;

    //Language
    var lang: any = getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
    var isRtlLan: any = isRtlLanguage(lang);
    //rtl
    var getRtlVal: any = getTextByPathList(pPrNode, ["attrs", "rtl"]);
    if (getRtlVal === undefined) {
        getRtlVal = getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (getRtlVal === undefined && type != "shape") {
            getRtlVal = getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    var isRTL: any = resolveParagraphRtl(pNode, getRtlVal, warpObj);
    var dirStr: any = isRTL ? "rtl" : "ltr";

    var linkID: any = getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
    var linkTooltip: any = "";
    var defLinkClr: any;
    if (linkID !== undefined) {
        linkTooltip = getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]) ?? "";
        if (linkTooltip !== "") {
            linkTooltip = "title='" + linkTooltip + "'";
        }
        defLinkClr = getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);

        var linkClrNode: any = getTextByPathList(node, ["a:rPr", "a:solidFill"]);// getTextByPathList(node, ["a:rPr", "a:solidFill"]);
        var rPrlinkClr: any = getSolidFill(linkClrNode, undefined, undefined, warpObj);


        //console.log("genSpanElement defLinkClr: ", defLinkClr, "rPrlinkClr:", rPrlinkClr)
        if (rPrlinkClr !== undefined && rPrlinkClr != "") {
            defLinkClr = rPrlinkClr;
        }

    }
    /////////////////////////////////////////////////////////////////////////////////////
    //getFontColor
    var fontClrPr: any = getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
    var fontClrType: any = fontClrPr[2];
    //console.log("genSpanElement fontClrPr: ", fontClrPr, "linkID", linkID);
    if (fontClrType == "solid") {
        if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
            styleText += "color: #" + fontClrPr[0] + ";";
        }
        else if (linkID !== undefined && defLinkClr !== undefined) {
            styleText += "color: #" + defLinkClr + ";";
        }

        if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
            styleText += "text-shadow:" + fontClrPr[1] + ";";
        }
        if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
            styleText += "background-color: #" + fontClrPr[3] + ";";
        }
    } else if (fontClrType == "pattern" || fontClrType == "pic" || fontClrType == "gradient") {
        if (fontClrType == "pattern") {
            styleText += "background:" + fontClrPr[0][0] + ";";
            if (fontClrPr[0][1] !== null && fontClrPr[0][1] !== undefined && fontClrPr[0][1] != "") {
                styleText += "background-size:" + fontClrPr[0][1] + ";";//" 2px 2px;" +
            }
            if (fontClrPr[0][2] !== null && fontClrPr[0][2] !== undefined && fontClrPr[0][2] != "") {
                styleText += "background-position:" + fontClrPr[0][2] + ";";//" 2px 2px;" +
            }
            // styleText += "-webkit-background-clip: text;" +
            //     "background-clip: text;" +
            //     "color: transparent;" +
            //     "-webkit-text-stroke: " + fontClrPr[1].border + ";" +
            //     "filter: " + fontClrPr[1].effcts + ";";
        } else if (fontClrType == "pic") {
            styleText += fontClrPr[0] + ";";
            // styleText += "-webkit-background-clip: text;" +
            //     "background-clip: text;" +
            //     "color: transparent;" +
            //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";
        } else if (fontClrType == "gradient") {

            var colorAry: any = fontClrPr[0].color;
            var rot: any = fontClrPr[0].rot;

            styleText += "background: linear-gradient(" + rot + "deg,";
            for (var i: any = 0; i < colorAry.length; i++) {
                if (i == colorAry.length - 1) {
                    styleText += "#" + colorAry[i] + ");";
                } else {
                    styleText += "#" + colorAry[i] + ", ";
                }
            }
            // styleText += "-webkit-background-clip: text;" +
            //     "background-clip: text;" +
            //     "color: transparent;" +
            //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";

        }
        styleText += "-webkit-background-clip: text;" +
            "background-clip: text;" +
            "color: transparent;";
        if (fontClrPr[1].border !== undefined && fontClrPr[1].border !== "") {
            styleText += "-webkit-text-stroke: " + fontClrPr[1].border + ";";
        }
        if (fontClrPr[1].effcts !== undefined && fontClrPr[1].effcts !== "") {
            styleText += "filter: " + fontClrPr[1].effcts + ";";
        }
    }
    var font_size: any = getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
    //text_style += "font-size:" + font_size + ";"
    
    text_style += "font-size:" + font_size + ";" +
        // marLStr +
        "font-family:" + getFontType(node, type, warpObj, pFontStyle) + ";" +
        "font-weight:" + getFontBold(node, type, slideMasterTextStyles) + ";" +
        "font-style:" + getFontItalic(node, type, slideMasterTextStyles) + ";" +
        "text-decoration:" + getFontDecoration(node, type, slideMasterTextStyles) + ";" +
        "text-align:" + getTextHorizontalAlign(node, pNode, type, warpObj) + ";" +
        "vertical-align:" + getTextVerticalAlign(node, type, slideMasterTextStyles) + ";";
    //rNodeLength
    //console.log("genSpanElement node:", node, "lang:", lang, "isRtlLan:", isRtlLan, "span parent dir:", dirStr)
    // runs follow the paragraph direction; only RTL-language runs are marked
    // (v1 forced direction:ltr on every other run)
    if (isRtlLan) {
        styleText += "direction:rtl;";
    }
    // } else if (dirStr == "rtl" && isRtlLan ) {
    //     styleText += "direction:rtl;";

    // } else if (dirStr == "ltr" && !isRtlLan ) {
    //     styleText += "direction:ltr;";
    // } else if (dirStr == "ltr" && isRtlLan){
    //     styleText += "direction:ltr;";
    // }else{
    //     styleText += "direction:inherit;";
    // }

    // if (dirStr == "rtl" && !isRtlLan) { //|| rIndex === undefined
    //     styleText += "direction:ltr;";
    // } else if (dirStr == "rtl" && isRtlLan) {
    //     styleText += "direction:rtl;";
    // } else if (dirStr == "ltr" && !isRtlLan) {
    //     styleText += "direction:ltr;";
    // } else if (dirStr == "ltr" && isRtlLan) {
    //     styleText += "direction:rtl;";
    // } else {
    //     styleText += "direction:inherit;";
    // }

    //     //"direction:" + dirStr + ";";
    //if (rNodeLength == 1 || rIndex == 0 ){
    //styleText += "display: table-cell;white-space: nowrap;";
    //}
    var highlight: any = getTextByPathList(node, ["a:rPr", "a:highlight"]);
    if (highlight !== undefined) {
        styleText += "background-color:#" + getSolidFill(highlight, undefined, undefined, warpObj) + ";";
        //styleText += "Opacity:" + getColorOpacity(highlight) + ";";
    }

    //letter-spacing:
    var spcNode: any = getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
    if (spcNode === undefined) {
        spcNode = getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
        if (spcNode === undefined) {
            spcNode = getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
        }
    }
    if (spcNode !== undefined) {
        var ltrSpc: any = parseInt(spcNode) / 100; //pt
        styleText += "letter-spacing: " + ltrSpc + "px;";// + "pt;";
    }

    //Text Cap Types
    var capNode: any = getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
    if (capNode === undefined) {
        capNode = getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
        if (capNode === undefined) {
            capNode = getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
        }
    }
    if (capNode == "small" || capNode == "all") {
        styleText += "text-transform: uppercase";
    }
    //styleText += "word-break: break-word;";
    //console.log("genSpanElement node: ", node, ", capNode: ", capNode, ",pPrNodeLaout: ", pPrNodeLaout, ", pPrNodeMaster: ", pPrNodeMaster, "warpObj:", warpObj);

    var cssName: any = "";

    if (styleText in warpObj.ctx.styleTable) {
        cssName = warpObj.ctx.styleTable[styleText]["name"];
    } else {
        cssName = "_css_" + (Object.keys(warpObj.ctx.styleTable).length + 1);
        warpObj.ctx.styleTable[styleText] = {
            "name": cssName,
            "text": styleText
        };
    }
    var linkColorSyle: any = "";
    if (fontClrType == "solid" && linkID !== undefined) {
        linkColorSyle = "style='color: inherit;'";
    }

    var linkRes: any = (linkID !== undefined && linkID != "") ? warpObj["slideResObj"][linkID] : undefined;
    if (linkRes !== undefined) {
        var linkURL: any;
        var linkTarget: any;
        if (linkRes["type"] == "slide") {
            // link to another slide of the deck: in-page anchor handled by the viewer
            var linkedSlide: any = /(d+).xml$/.exec(linkRes["target"]);
            linkURL = "#" + warpObj.ctx.idPrefix + "-slide-" + (linkedSlide ? linkedSlide[1] : "");
            linkTarget = " data-pptx-slide-link='" + (linkedSlide ? linkedSlide[1] : "") + "'";
        } else {
            // only safe schemes; javascript:, data:, vbscript: ... are dropped (security)
            linkURL = sanitizeUrl(linkRes["target"]);
            linkTarget = " target='_blank' rel='noopener noreferrer'";
        }
        return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'><a href='" + linkURL + "' " + linkColorSyle + "  " + linkTooltip + linkTarget + ">" +
                text + "</a>" + closeElemnt;
    } else {
        // spaces/tabs are kept as-is; `.text-block { white-space: pre-wrap }` preserves them
        // while still wrapping at word boundaries (v1 replaced every space with &nbsp;)
        return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'>" + text + closeElemnt;
    }

}

// ported from legacy/js/pptxjs.js:9537
export function getPregraphMargn(pNode?: any, idx?: any, type?: any, isBullate?: any, warpObj?: any): any{
    if (!isBullate){
        return ["",0];
    }
    var marLStr: any = "", marRStr: any = "" , maginVal: any = 0;
    var pPrNode: any = pNode["a:pPr"];
    var layoutMasterNode: any = getLayoutAndMasterNode(pNode, idx, type, warpObj);
    var pPrNodeLaout: any = layoutMasterNode.nodeLaout;
    var pPrNodeMaster: any = layoutMasterNode.nodeMaster;
    
    // var lang = getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
    // var isRtlLan = (lang !== undefined && rtl_langs_array.indexOf(lang) !== -1) ? true : false;
    //rtl
    var getRtlVal: any = getTextByPathList(pPrNode, ["attrs", "rtl"]);
    if (getRtlVal === undefined) {
        getRtlVal = getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (getRtlVal === undefined && type != "shape") {
            getRtlVal = getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    var isRTL: any = resolveParagraphRtl(pNode, getRtlVal, warpObj);
    var dirStr: any = isRTL ? "rtl" : "ltr";

    //align
    var alignNode: any = getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
    if (alignNode === undefined) {
        alignNode = getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
        if (alignNode === undefined) {
            alignNode = getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
        }
    }
    //indent?
    var indentNode: any = getTextByPathList(pPrNode, ["attrs", "indent"]);
    if (indentNode === undefined) {
        indentNode = getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
        if (indentNode === undefined) {
            indentNode = getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
        }
    }
    var indent: any = 0;
    if (indentNode !== undefined) {
        indent = parseInt(indentNode) * slideFactor;
    }
    //
    //marL
    var marLNode: any = getTextByPathList(pPrNode, ["attrs", "marL"]);
    if (marLNode === undefined) {
        marLNode = getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
        if (marLNode === undefined) {
            marLNode = getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
        }
    }
    var marginLeft: any = 0;
    if (marLNode !== undefined) {
        marginLeft = parseInt(marLNode) * slideFactor;
    }
    if ((indentNode !== undefined || marLNode !== undefined)) {
        //var lvlIndent = defTabSz * lvl;

        if (isRTL) {// && alignNode == "r") {
            //marLStr = "margin-right: ";
            marLStr = "padding-right: ";
        } else {
            //marLStr = "margin-left: ";
            marLStr = "padding-left: ";
        }
        if (isBullate) {
            maginVal = Math.abs(0 - indent);
            marLStr += maginVal + "px;";  // (minus bullate numeric lenght/size - TODO
        } else {
            maginVal = Math.abs(marginLeft + indent);
            marLStr += maginVal + "px;";  // (minus bullate numeric lenght/size - TODO
        }
    }

    //marR?
    var marRNode: any = getTextByPathList(pPrNode, ["attrs", "marR"]);
    if (marRNode === undefined && marLNode === undefined) {
        //need to check if this posble - TODO
        marRNode = getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
        if (marRNode === undefined) {
            marRNode = getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
        }
    }
    if (marRNode !== undefined && isBullate) {
        var marginRight: any = parseInt(marRNode) * slideFactor;
        if (isRTL) {// && alignNode == "r") {
            //marRStr = "margin-right: ";
            marRStr = "padding-right: ";
        } else {
            //marRStr = "margin-left: ";
            marRStr = "padding-left: ";
        }
        marRStr += Math.abs(0 - indent) + "px;";
    }


    return [marLStr, maginVal];
}

// ported from legacy/js/pptxjs.js:10458
export function getVerticalMargins(pNode?: any, textBodyNode?: any, type?: any, idx?: any, warpObj?: any): any {
    //margin-top ; 
    //a:pPr => a:spcBef => a:spcPts (/100) | a:spcPct (/?)
    //margin-bottom
    //a:pPr => a:spcAft => a:spcPts (/100) | a:spcPct (/?)
    //+
    //a:pPr =>a:lnSpc => a:spcPts (/?) | a:spcPct (/?)
    //console.log("getVerticalMargins ", pNode, type,idx, warpObj)
    //var lstStyle = textBodyNode["a:lstStyle"];
    var lvl: any = 1
    var spcBefNode: any = getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
    var spcAftNode: any = getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
    var lnSpcNode: any = getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
    var lnSpcNodeType: any = "Pct";
    if (lnSpcNode === undefined) {
        lnSpcNode = getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
        if (lnSpcNode !== undefined) {
            lnSpcNodeType = "Pts";
        }
    }
    var lvlNode: any = getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }
    var fontSize: any;
    if (getTextByPathList(pNode, ["a:r"]) !== undefined) {
        var fontSizeStr: any = getFontSize(pNode["a:r"], textBodyNode,undefined, lvl, type, warpObj);
        if (fontSizeStr != "inherit") {
            fontSize = parseInt(fontSizeStr); //pt
        }
    }
    //var spcBef = "";
    //console.log("getVerticalMargins 1", fontSizeStr, fontSize, lnSpcNode, parseInt(lnSpcNode) / 100000, spcBefNode, spcAftNode)
    // if(spcBefNode !== undefined){
    //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
    // }
    // else{
    //    //i did not found case with percentage 
    //     spcBefNode = getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPct","attrs","val"]);
    //     if(spcBefNode !== undefined){
    //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
    //     }
    // }
    //var spcAft = "";
    // if(spcAftNode !== undefined){
    //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
    // }
    // else{
    //    //i did not found case with percentage 
    //     spcAftNode = getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPct","attrs","val"]);
    //     if(spcAftNode !== undefined){
    //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
    //     }
    // }
    // if(spcAftNode !== undefined){
    //     //check in layout and then in master
    // }
    var isInLayoutOrMaster: any = true;
    if(type == "shape" || type == "textBox"){
        isInLayoutOrMaster = false;
    }
    if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
        //check in layout
        if (idx !== undefined) {
            var laypPrNode: any = getTextByPathList(warpObj, ["slideLayoutTables", "idxTable", idx, "p:txBody", "a:p", (lvl - 1), "a:pPr"]);

            if (spcBefNode === undefined) {
                spcBefNode = getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                // if(spcBefNode !== undefined){
                //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
                // } 
                // else{
                //    //i did not found case with percentage 
                //     spcBefNode = getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPct","attrs","val"]);
                //     if(spcBefNode !== undefined){
                //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
                //     }
                // }
            }

            if (spcAftNode === undefined) {
                spcAftNode = getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                // if(spcAftNode !== undefined){
                //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
                // }
                // else{
                //    //i did not found case with percentage 
                //     spcAftNode = getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPct","attrs","val"]);
                //     if(spcAftNode !== undefined){
                //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
                //     }
                // }
            }

            if (lnSpcNode === undefined) {
                lnSpcNode = getTextByPathList(laypPrNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                if (lnSpcNode === undefined) {
                    lnSpcNode = getTextByPathList(laypPrNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                    if (lnSpcNode !== undefined) {
                        lnSpcNodeType = "Pts";
                    }
                }
            }
        }

    }
    if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
        //check in master
        //slideMasterTextStyles
        var slideMasterTextStyles: any = warpObj["slideMasterTextStyles"];
        var dirLoc: any = "";
        var lvl: any = "a:lvl" + lvl + "pPr";
        switch (type) {
            case "title":
            case "ctrTitle":
                dirLoc = "p:titleStyle";
                break;
            case "body":
            case "obj":
            case "dt":
            case "ftr":
            case "sldNum":
            case "textBox":
            // case "shape":
                dirLoc = "p:bodyStyle";
                break;
            case "shape":
            //case "textBox":
            default:
                dirLoc = "p:otherStyle";
        }
        // if (type == "shape" || type == "textBox") {
        //     lvl = "a:lvl1pPr";
        // }
        var inLvlNode: any = getTextByPathList(slideMasterTextStyles, [dirLoc, lvl]);
        if (inLvlNode !== undefined) {
            if (spcBefNode === undefined) {
                spcBefNode = getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                // if(spcBefNode !== undefined){
                //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
                // } 
                // else{
                //    //i did not found case with percentage 
                //     spcBefNode = getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPct","attrs","val"]);
                //     if(spcBefNode !== undefined){
                //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
                //     }
                // }
            }

            if (spcAftNode === undefined) {
                spcAftNode = getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                // if(spcAftNode !== undefined){
                //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
                // }
                // else{
                //    //i did not found case with percentage 
                //     spcAftNode = getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPct","attrs","val"]);
                //     if(spcAftNode !== undefined){
                //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
                //     }
                // }
            }

            if (lnSpcNode === undefined) {
                lnSpcNode = getTextByPathList(inLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                if (lnSpcNode === undefined) {
                    lnSpcNode = getTextByPathList(inLvlNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                    if (lnSpcNode !== undefined) {
                        lnSpcNodeType = "Pts";
                    }
                }
            }
        }
    }
    var spcBefor: any = 0, spcAfter: any = 0, spcLines: any = 0;
    var marginTopBottomStr: any = "";
    if (spcBefNode !== undefined) {
        spcBefor = parseInt(spcBefNode) / 100;
    }
    if (spcAftNode !== undefined) {
        spcAfter = parseInt(spcAftNode) / 100;
    }
    
    if (lnSpcNode !== undefined && fontSize !== undefined) {
        if (lnSpcNodeType == "Pts") {
            marginTopBottomStr += "padding-top: " + ((parseInt(lnSpcNode) / 100) - fontSize) + "px;";//+ "pt;";
        } else {
            var fct: any = parseInt(lnSpcNode) / 100000;
            spcLines = fontSize * (fct - 1) - fontSize;// fontSize *
            var pTop: any = (fct > 1) ? spcLines : 0;
            var pBottom: any = (fct > 1) ? fontSize : 0;
            // marginTopBottomStr += "padding-top: " + spcLines + "pt;";
            // marginTopBottomStr += "padding-bottom: " + pBottom + "pt;";
            marginTopBottomStr += "padding-top: " + pBottom + "px;";// + "pt;";
            marginTopBottomStr += "padding-bottom: " + spcLines + "px;";// + "pt;";
        }
    }

    //if (spcBefNode !== undefined || lnSpcNode !== undefined) {
    marginTopBottomStr += "margin-top: " + (spcBefor - 1) + "px;";// + "pt;"; //margin-top: + spcLines // minus 1 - to fix space
    //}
    if (spcAftNode !== undefined || lnSpcNode !== undefined) {
        //marginTopBottomStr += "margin-bottom: " + ((spcAfter - fontSize < 0) ? 0 : (spcAfter - fontSize)) + "pt;"; //margin-bottom: + spcLines
        //marginTopBottomStr += "margin-bottom: " + spcAfter * (1 / 4) + "px;";// + "pt;";
        marginTopBottomStr += "margin-bottom: " + spcAfter  + "px;";// + "pt;";
    }

    //console.log("getVerticalMargins 2 fontSize:", fontSize, "lnSpcNode:", lnSpcNode, "spcLines:", spcLines, "spcBefor:", spcBefor, "spcAfter:", spcAfter)
    //console.log("getVerticalMargins 3 ", marginTopBottomStr, pNode, warpObj)

    //return spcAft + spcBef;
    return marginTopBottomStr;
}

// ported from legacy/js/pptxjs.js:10672
export function getHorizontalAlign(node?: any, textBodyNode?: any, idx?: any, type?: any, prg_dir?: any, warpObj?: any): any {
    var algn: any = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
    if (algn === undefined) {
        //var layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
        // var pPrNodeLaout = layoutMasterNode.nodeLaout;
        // var pPrNodeMaster = layoutMasterNode.nodeMaster;
        var lvlIdx: any = 1;
        var lvlNode: any = getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvlIdx = parseInt(lvlNode) + 1;
        }
        var lvlStr: any = "a:lvl" + lvlIdx + "pPr";

        var lstStyle: any = textBodyNode["a:lstStyle"];
        algn = getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);

        if (algn === undefined && idx !== undefined ) {
            //slidelayout
            algn = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
            if (algn === undefined) {
                algn = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                if (algn === undefined) {
                    algn = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvlIdx - 1), "a:pPr", "attrs", "algn"]);
                }
            }
        }
        if (algn === undefined) {
            if (type !== undefined) {
                //slidelayout
                algn = getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);

                if (algn === undefined) {
                    //masterlayout
                    if (type == "title" || type == "ctrTitle") {
                        algn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                    } else if (type == "body" || type == "obj" || type == "subTitle") {
                        algn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                    } else if (type == "shape" || type == "diagram") {
                        algn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr, "attrs", "algn"]);
                    } else if (type == "textBox") {
                        algn = getTextByPathList(warpObj, ["defaultTextStyle", lvlStr, "attrs", "algn"]);
                    } else {
                        algn = getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                    }
                }
            } else {
                algn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
            }
        }
    }

    if (algn === undefined) {
        if (type == "title" || type == "subTitle" || type == "ctrTitle") {
            return "h-mid";
        } else if (type == "sldNum") {
            return "h-right";
        }
    }
    if (algn !== undefined) {
        switch (algn) {
            case "l":
                if (prg_dir == "pregraph-rtl"){
                    //return "h-right";
                    return "h-left-rtl";
                }else{
                    return "h-left";
                }
                break;
            case "r":
                if (prg_dir == "pregraph-rtl") {
                    //return "h-left";
                    return "h-right-rtl";
                }else{
                    return "h-right";
                }
                break;
            case "ctr":
                return "h-mid";
                break;
            case "just":
            case "dist":
            default:
                return "h-" + algn;
        }
    }
    //return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
}

/** True when a paragraph has no explicit direction in the file (layout/master included). */
function isPolicyRtl(node: any, idx: any, type: any, warpObj: any): boolean {
    if (getTextByPathList(node, ["a:pPr", "attrs", "rtl"]) !== undefined) return false;
    var lm: any = getLayoutAndMasterNode(node, idx, type, warpObj);
    if (getTextByPathList(lm.nodeLaout, ["attrs", "rtl"]) !== undefined) return false;
    if (type != "shape" && getTextByPathList(lm.nodeMaster, ["attrs", "rtl"]) !== undefined) return false;
    return true;
}

// ported from legacy/js/pptxjs.js:10759
export function getPregraphDir(node?: any, textBodyNode?: any, idx?: any, type?: any, warpObj?: any): any {
    var rtl: any = getTextByPathList(node, ["a:pPr", "attrs", "rtl"]);
    //console.log("getPregraphDir node:", node, "textBodyNode", textBodyNode, "rtl:", rtl, "idx", idx, "type", type, "warpObj", warpObj)
  

    if (rtl === undefined) {
        var layoutMasterNode: any = getLayoutAndMasterNode(node, idx, type, warpObj);
        var pPrNodeLaout: any = layoutMasterNode.nodeLaout;
        var pPrNodeMaster: any = layoutMasterNode.nodeMaster;
        rtl = getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (rtl === undefined && type != "shape") {
            rtl = getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }

    if (rtl == "1") {
        return "pregraph-rtl";
    } else if (rtl == "0") {
        return "pregraph-ltr";
    }
    // no explicit direction: presentation default / textDirection option
    if (hasDirectionPolicy(warpObj)) {
        return resolveParagraphRtl(node, undefined, warpObj) ? "pregraph-rtl" : "pregraph-ltr";
    }
    return "pregraph-inherit";

    // var contentDir = getContentDir(type, warpObj);
    // console.log("getPregraphDir node:", node["a:r"], "rtl:", rtl, "idx", idx, "type", type, "contentDir:", contentDir)

    // if (contentDir == "content"){
    //     return "pregraph-ltr";
    // } else if (contentDir == "content-rtl"){ 
    //     return "pregraph-rtl";
    // }
    // return "";
}

// ported from legacy/js/pptxjs.js:10791
export function getVerticalAlign(node?: any, slideLayoutSpNode?: any, slideMasterSpNode?: any, type?: any): any {

    //X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
    var anchor: any = getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
    //console.log("getVerticalAlign anchor:", anchor, "slideLayoutSpNode: ", slideLayoutSpNode)
    if (anchor === undefined) {
        //console.log("getVerticalAlign type:", type," node:", node, "slideLayoutSpNode:", slideLayoutSpNode, "slideMasterSpNode:", slideMasterSpNode)
        anchor = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
        if (anchor === undefined) {
            anchor = getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
            if (anchor === undefined) {
                //"If this attribute is omitted, then a value of t, or top is implied."
                anchor = "t";//getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
            }
        }
    }
    //console.log("getVerticalAlign:", node, slideLayoutSpNode, slideMasterSpNode, type, anchor)
    return (anchor === "ctr")?"v-mid" : ((anchor === "b") ? "v-down" : "v-up");
}

// ported from legacy/js/pptxjs.js:10811
export function getContentDir(node?: any, type?: any, warpObj?: any): any {
    return "content";
    var defRtl: any = getTextByPathList(node, ["p:txBody", "a:lstStyle", "a:defPPr", "attrs", "rtl"]);
    if (defRtl !== undefined) {
        if (defRtl == "1"){
            return "content-rtl";
        } else if (defRtl == "0") {
            return "content";
        }
    }
    //var lvl1Rtl = getTextByPathList(node, ["p:txBody", "a:lstStyle", "lvl1pPr", "attrs", "rtl"]);
    // if (lvl1Rtl !== undefined) {
    //     if (lvl1Rtl == "1") {
    //         return "content-rtl";
    //     } else if (lvl1Rtl == "0") {
    //         return "content";
    //     }
    // }
    var rtlCol: any = getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "rtlCol"]);
    if (rtlCol !== undefined) {
        if (rtlCol == "1") {
            return "content-rtl";
        } else if (rtlCol == "0") {
            return "content";
        }
    }
    //console.log("getContentDir node:", node, "rtlCol:", rtlCol)

    if (type === undefined) {
        return "content";
    }
    var slideMasterTextStyles: any = warpObj["slideMasterTextStyles"];
    var dirLoc: any = "";

    switch (type) {
        case "title":
        case "ctrTitle":
            dirLoc = "p:titleStyle";
            break;
        case "body":
        case "dt":
        case "ftr":
        case "sldNum":
        case "textBox":
            dirLoc = "p:bodyStyle";
            break;
        case "shape":
            dirLoc = "p:otherStyle";
    }
    if (slideMasterTextStyles !== undefined && dirLoc !== "") {
        var dirVal: any = getTextByPathList(slideMasterTextStyles[dirLoc], ["a:lvl1pPr", "attrs", "rtl"]);
        if (dirVal == "1") {
            return "content-rtl";
        }
    } 
    // else {
    //     if (type == "textBox") {
    //         var dirVal = getTextByPathList(warpObj, ["defaultTextStyle", "a:lvl1pPr", "attrs", "rtl"]);
    //         if (dirVal == "1") {
    //             return "content-rtl";
    //         }
    //     }
    // }
    return "content";
    //console.log("getContentDir() type:", type, "slideMasterTextStyles:", slideMasterTextStyles,"dirNode:",dirVal)
}

// ported from legacy/js/pptxjs.js:10878
export function getFontType(node?: any, type?: any, warpObj?: any, pFontStyle?: any): any {
    var typeface: any = getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

    if (typeface === undefined) {
        var fontIdx: any = "";
        var fontGrup: any = "";
        if (pFontStyle !== undefined) {
            fontIdx = getTextByPathList(pFontStyle, ["attrs", "idx"]);
        }
        var fontSchemeNode: any = getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fontScheme"]);
        if (fontIdx == "") {
            if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                fontIdx = "major";
            } else {
                fontIdx = "minor";
            }
        }
        fontGrup = "a:" + fontIdx + "Font";
        typeface = getTextByPathList(fontSchemeNode, [fontGrup, "a:latin", "attrs", "typeface"]);
    }

    return (typeface === undefined) ? "inherit" : typeface;
}

// ported from legacy/js/pptxjs.js:10902
export function getFontColorPr(node?: any, pNode?: any, lstStyle?: any, pFontStyle?: any, lvl?: any, idx?: any, type?: any, warpObj?: any): any {
    //text border using: text-shadow: -1px 0 black, 0 1px black, 1px 0 black, 0 -1px black;
    //{getFontColor(..) return color} -> getFontColorPr(..) return array[color,textBordr/shadow]
    //https://stackoverflow.com/questions/2570972/css-font-border
    //https://www.w3schools.com/cssref/css3_pr_text-shadow.asp
    //themeContent
    //console.log("getFontColorPr>> type:", type, ", node: ", node)
    var rPrNode: any = getTextByPathList(node, ["a:rPr"]);
    var filTyp: any, color: any, textBordr: any, colorType: any = "", highlightColor: any = "";
    //console.log("getFontColorPr type:", type, ", node: ", node, "pNode:", pNode, "pFontStyle:", pFontStyle)
    if (rPrNode !== undefined) {
        filTyp = getFillType(rPrNode);
        if (filTyp == "SOLID_FILL") {
            var solidFillNode: any = rPrNode["a:solidFill"];// getTextByPathList(node, ["a:rPr", "a:solidFill"]);
            color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
            var highlightNode: any = rPrNode["a:highlight"];
            if (highlightNode !== undefined) {
                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
            }
            colorType = "solid";
        } else if (filTyp == "PATTERN_FILL") {
            var pattFill: any = rPrNode["a:pattFill"];// getTextByPathList(node, ["a:rPr", "a:pattFill"]);
            color = getPatternFill(pattFill, warpObj);
            colorType = "pattern";
        } else if (filTyp == "PIC_FILL") {
            color = getBgPicFill(rPrNode, "slideBg", warpObj, undefined, undefined);
            //color = getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
            colorType = "pic";
        } else if (filTyp == "GRADIENT_FILL") {
            var shpFill: any = rPrNode["a:gradFill"];
            color = getGradientFill(shpFill, warpObj);
            colorType = "gradient";
        } 
    }
    if (color === undefined && getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]) !== undefined) {
        //lstStyle
        var lstStyledefRPr: any = getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]);
        filTyp = getFillType(lstStyledefRPr);
        if (filTyp == "SOLID_FILL") {
            var solidFillNode: any = lstStyledefRPr["a:solidFill"];// getTextByPathList(node, ["a:rPr", "a:solidFill"]);
            color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
            var highlightNode: any = lstStyledefRPr["a:highlight"];
            if (highlightNode !== undefined) {
                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
            }
            colorType = "solid";
        } else if (filTyp == "PATTERN_FILL") {
            var pattFill: any = lstStyledefRPr["a:pattFill"];// getTextByPathList(node, ["a:rPr", "a:pattFill"]);
            color = getPatternFill(pattFill, warpObj);
            colorType = "pattern";
        } else if (filTyp == "PIC_FILL") {
            color = getBgPicFill(lstStyledefRPr, "slideBg", warpObj, undefined, undefined);
            //color = getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
            colorType = "pic";
        } else if (filTyp == "GRADIENT_FILL") {
            var shpFill: any = lstStyledefRPr["a:gradFill"];
            color = getGradientFill(shpFill, warpObj);
            colorType = "gradient";
        }

    }
    if (color === undefined) {
        var sPstyle: any = getTextByPathList(pNode, ["p:style", "a:fontRef"]);
        if (sPstyle !== undefined) {
            color = getSolidFill(sPstyle, undefined, undefined, warpObj);
            if (color !== undefined) {
                colorType = "solid";
            }
            var highlightNode: any = sPstyle["a:highlight"]; //is "a:highlight" node in 'a:fontRef' ?
            if (highlightNode !== undefined) {
                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
            }
        }
        if (color === undefined) {
            if (pFontStyle !== undefined) {
                color = getSolidFill(pFontStyle, undefined, undefined, warpObj);
                if (color !== undefined) {
                    colorType = "solid";
                }
            }
        }
    }
    //console.log("getFontColorPr node", node, "colorType: ", colorType,"color: ",color)

    if (color === undefined) {

        var layoutMasterNode: any = getLayoutAndMasterNode(pNode, idx, type, warpObj);
        var pPrNodeLaout: any = layoutMasterNode.nodeLaout;
        var pPrNodeMaster: any = layoutMasterNode.nodeMaster;

        if (pPrNodeLaout !== undefined) {
            var defRpRLaout: any = getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:solidFill"]);
            if (defRpRLaout !== undefined) {
                color = getSolidFill(defRpRLaout, undefined, undefined, warpObj);
                var highlightNode: any = getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:highlight"]);
                if (highlightNode !== undefined) {
                    highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                }
                colorType = "solid";
            }
        }
        if (color === undefined) {

            if (pPrNodeMaster !== undefined) {
                var defRprMaster: any = getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:solidFill"]);
                if (defRprMaster !== undefined) {
                    color = getSolidFill(defRprMaster, undefined, undefined, warpObj);
                    var highlightNode: any = getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:highlight"]);
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                }
            }
        }
    }
    var txtEffects: any = [];
    var txtEffObj: any = {}
    //textBordr
    var txtBrdrNode: any = getTextByPathList(node, ["a:rPr", "a:ln"]);
    var textBordr: any = "";
    if (txtBrdrNode !== undefined && txtBrdrNode["a:noFill"] === undefined) {
        var txBrd: any = getBorder(node, pNode, false, "text", warpObj);
        var txBrdAry: any = txBrd.split(" ");
        //var brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("pt")))) + "px";
        var brdSize: any = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("px")))) + "px";
        var brdClr: any = txBrdAry[2];
        //var brdTyp = txBrdAry[1]; //not in use
        //console.log("getFontColorPr txBrdAry:", txBrdAry)
        if (colorType == "solid") {
            textBordr = "-" + brdSize + " 0 " + brdClr + ", 0 " + brdSize + " " + brdClr + ", " + brdSize + " 0 " + brdClr + ", 0 -" + brdSize + " " + brdClr;
            // if (oShadowStr != "") {
            //     textBordr += "," + oShadowStr;
            // } else {
            //     textBordr += ";";
            // }
            txtEffects.push(textBordr);
        } else {
            //textBordr = brdSize + " " + brdClr;
            txtEffObj.border = brdSize + " " + brdClr;
        }
    }
    // else {
    //     //if no border but exist/not exist shadow
    //     if (colorType == "solid") {
    //         textBordr = oShadowStr;
    //     } else {
    //         //TODO
    //     }
    // }
    var txtGlowNode: any = getTextByPathList(node, ["a:rPr", "a:effectLst", "a:glow"]);
    var oGlowStr: any = "";
    if (txtGlowNode !== undefined) {
        var glowClr: any = getSolidFill(txtGlowNode, undefined, undefined, warpObj);
        var rad: any = (txtGlowNode["attrs"]["rad"]) ? (txtGlowNode["attrs"]["rad"] * slideFactor) : 0;
        oGlowStr = "0 0 " + rad + "px #" + glowClr +
            ", 0 0 " + rad + "px #" + glowClr +
            ", 0 0 " + rad + "px #" + glowClr +
            ", 0 0 " + rad + "px #" + glowClr +
            ", 0 0 " + rad + "px #" + glowClr +
            ", 0 0 " + rad + "px #" + glowClr +
            ", 0 0 " + rad + "px #" + glowClr;
        if (colorType == "solid") {
            txtEffects.push(oGlowStr);
        } else {
            // txtEffObj.glow = {
            //     radiuse: rad,
            //     color: glowClr
            // } 
            txtEffects.push(
                "drop-shadow(0 0 " + rad / 3 + "px #" + glowClr + ") " +
                "drop-shadow(0 0 " + rad * 2 / 3 + "px #" + glowClr + ") " +
                "drop-shadow(0 0 " + rad + "px #" + glowClr + ")"
            );
        }
    }
    var txtShadow: any = getTextByPathList(node, ["a:rPr", "a:effectLst", "a:outerShdw"]);
    var oShadowStr: any = "";
    if (txtShadow !== undefined) {
        //https://developer.mozilla.org/en-US/docs/Web/CSS/filter-function/drop-shadow()
        //https://stackoverflow.com/questions/60468487/css-text-with-linear-gradient-shadow-and-text-outline
        //https://css-tricks.com/creating-playful-effects-with-css-text-shadows/
        //https://designshack.net/articles/css/12-fun-css-text-shadows-you-can-copy-and-paste/

        var shadowClr: any = getSolidFill(txtShadow, undefined, undefined, warpObj);
        var outerShdwAttrs: any = txtShadow["attrs"];
        // algn: "bl"
        // dir: "2640000"
        // dist: "38100"
        // rotWithShape: "0/1" - Specifies whether the shadow rotates with the shape if the shape is rotated.
        //blurRad (Blur Radius) - Specifies the blur radius of the shadow.
        //kx (Horizontal Skew) - Specifies the horizontal skew angle.
        //ky (Vertical Skew) - Specifies the vertical skew angle.
        //sx (Horizontal Scaling Factor) - Specifies the horizontal scaling slideFactor; negative scaling causes a flip.
        //sy (Vertical Scaling Factor) - Specifies the vertical scaling slideFactor; negative scaling causes a flip.
        var algn: any = outerShdwAttrs["algn"];
        var dir: any = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
        var dist: any = parseInt(outerShdwAttrs["dist"]) * slideFactor;//(px) //* (3 / 4); //(pt)
        var rotWithShape: any = outerShdwAttrs["rotWithShape"];
        var blurRad: any = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * slideFactor + "px") : "";
        var sx: any = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
        var sy: any = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
        var vx: any = dist * Math.sin(dir * Math.PI / 180);
        var hx: any = dist * Math.cos(dir * Math.PI / 180);

        //console.log("getFontColorPr outerShdwAttrs:", outerShdwAttrs, ", shadowClr:", shadowClr, ", algn: ", algn, ",dir: ", dir, ", dist: ", dist, ",rotWithShape: ", rotWithShape, ", color: ", color)

        if (!isNaN(vx) && !isNaN(hx)) {
            oShadowStr = hx + "px " + vx + "px " + blurRad + " #" + shadowClr;// + ";";
            if (colorType == "solid") {
                txtEffects.push(oShadowStr);
            } else {

                // txtEffObj.oShadow = {
                //     hx: hx,
                //     vx: vx,
                //     radius: blurRad,
                //     color: shadowClr
                // }

                //txtEffObj.oShadow = hx + "px " + vx + "px " + blurRad + " #" + shadowClr;

                txtEffects.push("drop-shadow(" + hx + "px " + vx + "px " + blurRad + " #" + shadowClr + ")");
            }
        }
        //console.log("getFontColorPr vx:", vx, ", hx: ", hx, ", sx: ", sx, ", sy: ", sy, ",oShadowStr: ", oShadowStr)
    }
    //console.log("getFontColorPr>>> color:", color)
    // if (color === undefined || color === "FFF") {
    //     color = "#000";
    // } else {
    //     color = "" + color;
    // }
    var text_effcts: any = "", txt_effects: any;
    if (colorType == "solid") {
        if (txtEffects.length > 0) {
            text_effcts = txtEffects.join(",");
        }
        txt_effects = text_effcts + ";"
    } else {
        if (txtEffects.length > 0) {
            text_effcts = txtEffects.join(" ");
        }
        txtEffObj.effcts = text_effcts;
        txt_effects = txtEffObj
    }
    //console.log("getFontColorPr txt_effects:", txt_effects)

    //return [color, textBordr, colorType];
    return [color, txt_effects, colorType, highlightColor];
}

// ported from legacy/js/pptxjs.js:11153
export function getFontSize(node?: any, textBodyNode?: any, pFontStyle?: any, lvl?: any, type?: any, warpObj?: any): any {
    // if(type == "sldNum")
    //console.log("getFontSize node:", node, "lstStyle", lstStyle, "lvl:", lvl, 'type:', type, "warpObj:", warpObj)
    var lstStyle: any = (textBodyNode !== undefined)? textBodyNode["a:lstStyle"] : undefined;
    var lvlpPr: any = "a:lvl" + lvl + "pPr";
    var fontSize: any = undefined;
    var sz: any, kern: any;
    if (node["a:rPr"] !== undefined) {
        fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
    }
    if (isNaN(fontSize) || fontSize === undefined && node["a:fld"] !== undefined) {
        sz = getTextByPathList(node["a:fld"], ["a:rPr", "attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
    }
    if ((isNaN(fontSize) || fontSize === undefined) && node["a:t"] === undefined) {
        sz = getTextByPathList(node["a:endParaRPr"], [ "attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
    }
    if ((isNaN(fontSize) || fontSize === undefined) && lstStyle !== undefined) {
        sz = getTextByPathList(lstStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
    }
    //a:spAutoFit
    var isAutoFit: any = false;
    var isKerning: any = false;
    if (textBodyNode !== undefined){
        var spAutoFitNode: any = getTextByPathList(textBodyNode, ["a:bodyPr", "a:spAutoFit"]);
        // if (spAutoFitNode === undefined) {
        //     spAutoFitNode = getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit"]);
        // }
        if (spAutoFitNode !== undefined){
            isAutoFit = true;
            isKerning = true;
        }
    }
    if (isNaN(fontSize) || fontSize === undefined) {
        // if (type == "shape" || type == "textBox") {
        //     type = "body";
        //     lvlpPr = "a:lvl1pPr";
        // }
        sz = getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
        kern = getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
        if (isKerning && kern !== undefined && !isNaN(fontSize) && (fontSize - parseInt(kern) / 100) > 0){
            fontSize = fontSize - parseInt(kern) / 100;
        }
    }

    if (isNaN(fontSize) || fontSize === undefined) {
        // if (type == "shape" || type == "textBox") {
        //     type = "body";
        //     lvlpPr = "a:lvl1pPr";
        // }
        sz = getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
        kern = getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
        if (sz === undefined) {
            if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                sz = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            } else if (type == "body" || type == "obj" || type == "dt" || type == "sldNum" || type === "textBox") {
                sz = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            }
            else if (type == "shape") {
                //textBox and shape text does not indent
                sz = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                isKerning = false;
            }

            if (sz === undefined) {
                sz = getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = (kern === undefined)? getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "kern"]) : undefined;
                isKerning = false;
            }
            //  else if (type === undefined || type == "shape") {
            //     sz = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
            //     kern = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            // } 
            // else if (type == "textBox") {
            //     sz = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
            //     kern = getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            // }
        } 
        fontSize = parseInt(sz) / 100;
        if (isKerning && kern !== undefined && !isNaN(fontSize) && ((fontSize - parseInt(kern) / 100) > parseInt(kern) / 100 )) {
            fontSize = fontSize - parseInt(kern) / 100;
            //fontSize =  parseInt(kern) / 100;
        }
    }

    var baseline: any = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
    if (baseline !== undefined && !isNaN(fontSize)) {
        var baselineVl: any = parseInt(baseline) / 100000;
        //fontSize -= 10; 
        // fontSize = fontSize * baselineVl;
        fontSize -= baselineVl;
    }

    if (!isNaN(fontSize)){
        var normAutofit: any = getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit", "attrs", "fontScale"]);
        if (normAutofit !== undefined && normAutofit != 0){
            //console.log("fontSize", fontSize, "normAutofit: ", normAutofit, normAutofit/100000)
            fontSize = Math.round(fontSize * (normAutofit / 100000))
        }
    }

    return isNaN(fontSize) ? ((type == "br") ? "initial" : "inherit") : (fontSize * fontSizeFactor + "px");// + "pt");
}

// ported from legacy/js/pptxjs.js:11263
export function getFontBold(node?: any, type?: any, slideMasterTextStyles?: any): any {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "inherit";
}

// ported from legacy/js/pptxjs.js:11267
export function getFontItalic(node?: any, type?: any, slideMasterTextStyles?: any): any {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "inherit";
}

// ported from legacy/js/pptxjs.js:11271
export function getFontDecoration(node?: any, type?: any, slideMasterTextStyles?: any): any {
    ///////////////////////////////Amir///////////////////////////////
    if (node["a:rPr"] !== undefined) {
        var underLine: any = node["a:rPr"]["attrs"]["u"] !== undefined ? node["a:rPr"]["attrs"]["u"] : "none";
        var strikethrough: any = node["a:rPr"]["attrs"]["strike"] !== undefined ? node["a:rPr"]["attrs"]["strike"] : 'noStrike';
        //console.log("strikethrough: "+strikethrough);

        if (underLine != "none" && strikethrough == "noStrike") {
            return "underline";
        } else if (underLine == "none" && strikethrough != "noStrike") {
            return "line-through";
        } else if (underLine != "none" && strikethrough != "noStrike") {
            return "underline line-through";
        } else {
            return "inherit";
        }
    } else {
        return "inherit";
    }
    /////////////////////////////////////////////////////////////////
    //return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "inherit";
}

// ported from legacy/js/pptxjs.js:11294
export function getTextHorizontalAlign(node?: any, pNode?: any, type?: any, warpObj?: any): any {
    //console.log("getTextHorizontalAlign: type: ", type, ", node: ", node)
    var getAlgn: any = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
    if (getAlgn === undefined) {
        getAlgn = getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);
    }
    if (getAlgn === undefined) {
        if (type == "title" || type == "ctrTitle" || type == "subTitle") {
            var lvlIdx: any = 1;
            var lvlNode: any = getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvlIdx = parseInt(lvlNode) + 1;
            }
            var lvlStr: any = "a:lvl" + lvlIdx + "pPr";
            getAlgn = getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
            if (getAlgn === undefined) {
                getAlgn = getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                if (getAlgn === undefined) {
                    getAlgn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                    if (getAlgn === undefined && type === "subTitle") {
                        getAlgn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                    }
                }
            }
        } else if (type == "body") {
            getAlgn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", "a:lvl1pPr", "attrs", "algn"]);
        } else {
            getAlgn = getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", "a:lvl1pPr", "attrs", "algn"]);
        }

    }

    var align: any = "inherit";
    if (getAlgn !== undefined) {
        switch (getAlgn) {
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
                align = "inherit";
        }
    }
    return align;
}

// ported from legacy/js/pptxjs.js:11351
export function getTextVerticalAlign(node?: any, type?: any, slideMasterTextStyles?: any): any {
    var baseline: any = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
    return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
}
