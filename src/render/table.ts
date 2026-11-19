// Ported from PPTXjs v1 (legacy/js/pptxjs.js). Internal rendering code: the
// node objects come from the simplified XML tree, so most values are `any`.
import type { WarpObj } from "./context";
import { getPosition, getSize } from "./nodes";
import { genTextBody } from "./text";
import { getBorder } from "./border";
import { getShapeFill } from "./fill";
import { getSolidFill } from "./color";
import { getTextByPathList } from "../core/xml";
import { slideFactor } from "./units";
// ported from legacy/js/pptxjs.js:9666
export function genTable(node?: any, warpObj?: any): any {
    var order: any = node["attrs"]["order"];
    var tableNode: any = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
    var xfrmNode: any = getTextByPathList(node, ["p:xfrm"]);
    /////////////////////////////////////////Amir////////////////////////////////////////////////
    var getTblPr: any = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
    var getColsGrid: any = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
    var tblDir: any = "";
    if (getTblPr !== undefined) {
        var isRTL: any = getTblPr["attrs"]["rtl"];
        tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
    }
    var firstRowAttr: any = getTblPr["attrs"]["firstRow"]; //associated element <a:firstRow> in the table styles
    var firstColAttr: any = getTblPr["attrs"]["firstCol"]; //associated element <a:firstCol> in the table styles
    var lastRowAttr: any = getTblPr["attrs"]["lastRow"]; //associated element <a:lastRow> in the table styles
    var lastColAttr: any = getTblPr["attrs"]["lastCol"]; //associated element <a:lastCol> in the table styles
    var bandRowAttr: any = getTblPr["attrs"]["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
    var bandColAttr: any = getTblPr["attrs"]["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
    //console.log("getTblPr: ", getTblPr);
    var tblStylAttrObj: any = {
        isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
        isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
        isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
        isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
        isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
        isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
    }

    var thisTblStyle: any;
    var tbleStyleId: any = getTblPr["a:tableStyleId"];
    if (tbleStyleId !== undefined) {
        var tbleStylList: any = warpObj.ctx.tableStyles["a:tblStyleLst"]["a:tblStyle"];
        if (tbleStylList !== undefined) {
            if (tbleStylList.constructor === Array) {
                for (var k: any = 0; k < tbleStylList.length; k++) {
                    if (tbleStylList[k]["attrs"]["styleId"] == tbleStyleId) {
                        thisTblStyle = tbleStylList[k];
                    }
                }
            } else {
                if (tbleStylList["attrs"]["styleId"] == tbleStyleId) {
                    thisTblStyle = tbleStylList;
                }
            }
        }
    }
    if (thisTblStyle !== undefined) {
        thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
        warpObj["thisTbiStyle"] = thisTblStyle;
    }
    var tblStyl: any = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
    var tblBorderStyl: any = getTextByPathList(tblStyl, ["a:tcBdr"]);
    var tbl_borders: any = "";
    if (tblBorderStyl !== undefined) {
        tbl_borders = getTableBorders(tblBorderStyl, warpObj);
    }
    var tbl_bgcolor: any = "";
    var tbl_opacity: any = 1;
    var tbl_bgFillschemeClr: any = getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
    //console.log( "thisTblStyle:", thisTblStyle, "warpObj:", warpObj)
    if (tbl_bgFillschemeClr !== undefined) {
        tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
    }
    if (tbl_bgFillschemeClr === undefined) {
        tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
        tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
    }
    if (tbl_bgcolor !== "") {
        tbl_bgcolor = "background-color: #" + tbl_bgcolor + ";";
    }
    ////////////////////////////////////////////////////////////////////////////////////////////
    var tableHtml: any = "<table " + tblDir + " style='border-collapse: collapse;" +
        getPosition(xfrmNode, node, undefined, undefined) +
        getSize(xfrmNode, undefined, undefined) +
        " z-index: " + order + ";" +
        tbl_borders + ";" +
        tbl_bgcolor + "'>";

    var trNodes: any = tableNode["a:tr"];
    if (trNodes.constructor !== Array) {
        trNodes = [trNodes];
    }
    //if (trNodes.constructor === Array) {
        //multi rows
        var totalrowSpan: any = 0;
        var rowSpanAry: any = [];
        for (var i: any = 0; i < trNodes.length; i++) {
            //////////////rows Style ////////////Amir
            var rowHeightParam: any = trNodes[i]["attrs"]["h"];
            var rowHeight: any = 0;
            var rowsStyl: any = "";
            if (rowHeightParam !== undefined) {
                rowHeight = parseInt(rowHeightParam) * slideFactor;
                rowsStyl += "height:" + rowHeight + "px;";
            }
            var fillColor: any = "";
            var row_borders: any = "";
            var fontClrPr: any = "";
            var fontWeight: any = "";
            var band_1H_fillColor: any;
            var band_2H_fillColor: any;

            if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                var bgFillschemeClr: any = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor: any = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                var rowTxtStyl: any = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    var local_fontColor: any = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontColor !== undefined) {
                        fontClrPr = local_fontColor;
                    }

                    var local_fontWeight: any = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight != "") {
                        fontWeight = local_fontWeight
                    }
                }
            }

            if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                var bgFillschemeClr: any = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor: any = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                var borderStyl: any = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    var local_row_borders: any = getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                var rowTxtStyl: any = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    var local_fontClrPr: any = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }
                    var local_fontWeight: any = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }

            } else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
                fillColor = "";
                row_borders = undefined;
                if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                    // console.log("i: ", i, 'thisTblStyle["a:band2H"]:', thisTblStyle["a:band2H"])
                    //check if there is a row bg
                    var bgFillschemeClr: any = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        var local_fillColor: any = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== "") {
                            fillColor = local_fillColor;
                            band_2H_fillColor = local_fillColor;
                        }
                    }


                    var borderStyl: any = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        var local_row_borders: any = getTableBorders(borderStyl, warpObj);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    var rowTxtStyl: any = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        var local_fontClrPr: any = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }
                    }

                    var local_fontWeight: any = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }
                if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                    var bgFillschemeClr: any = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        var local_fillColor: any = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== undefined) {
                            fillColor = local_fillColor;
                            band_1H_fillColor = local_fillColor;
                        }
                    }
                    var borderStyl: any = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        var local_row_borders: any = getTableBorders(borderStyl, warpObj);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    var rowTxtStyl: any = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        var local_fontClrPr: any = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }
                        var local_fontWeight: any = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                        if (local_fontWeight != "") {
                            fontWeight = local_fontWeight;
                        }
                    }
                }

            }
            //last row
            if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                var bgFillschemeClr: any = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor: any = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                    // var local_colorOpacity = getColorOpacity(bgFillschemeClr);
                    // if(local_colorOpacity !== undefined){
                    //     colorOpacity = local_colorOpacity;
                    // }
                }
                var borderStyl: any = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    var local_row_borders: any = getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                var rowTxtStyl: any = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    var local_fontClrPr: any = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }

                    var local_fontWeight: any = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }
            }
            rowsStyl += ((row_borders !== undefined) ? row_borders : "");
            rowsStyl += ((fontClrPr !== undefined) ? " color: #" + fontClrPr + ";" : "");
            rowsStyl += ((fontWeight != "") ? " font-weight:" + fontWeight + ";" : "");
            if (fillColor !== undefined && fillColor != "") {
                //rowsStyl += "background-color: rgba(" + hexToRgbNew(fillColor) + "," + colorOpacity + ");";
                rowsStyl += "background-color: #" + fillColor + ";";
            }
            tableHtml += "<tr style='" + rowsStyl + "'>";
            ////////////////////////////////////////////////

            var tcNodes: any = trNodes[i]["a:tc"];
            if (tcNodes !== undefined) {
                if (tcNodes.constructor === Array) {
                    //multi columns
                    var j: any = 0;
                    if (rowSpanAry.length == 0) {
                        rowSpanAry = new Array(tcNodes.length).fill(0);
                    }
                    var totalColSpan: any = 0;
                    while (j < tcNodes.length) {
                        if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                            var a_sorce: any;
                            //j=0 : first col
                            if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                                a_sorce = "a:firstCol";
                                if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) &&
                                    getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                    a_sorce = "a:seCell";
                                } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                    getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                    a_sorce = "a:neCell";
                                }
                            } else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                                !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                                !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                                j != (tcNodes.length - 1)) {

                                if ((j % 2) != 0) {

                                    var aBandNode: any = getTextByPathList(thisTblStyle, ["a:band2V"]);
                                    if (aBandNode === undefined) {
                                        aBandNode = getTextByPathList(thisTblStyle, ["a:band1V"]);
                                        if (aBandNode !== undefined) {
                                            a_sorce = "a:band2V";
                                        }
                                    } else {
                                        a_sorce = "a:band2V";
                                    }

                                }
                            }

                            if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                                a_sorce = "a:lastCol";
                                if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                    a_sorce = "a:swCell";
                                } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
                                    a_sorce = "a:nwCell";
                                }
                            }

                            var cellParmAry: any = getTableCellParams(tcNodes[j], getColsGrid, i , j , thisTblStyle, a_sorce, warpObj)
                            var text: any = cellParmAry[0];
                            var colStyl: any = cellParmAry[1];
                            var cssName: any = cellParmAry[2];
                            var rowSpan: any = cellParmAry[3];
                            var colSpan: any = cellParmAry[4];



                            if (rowSpan !== undefined) {
                                totalrowSpan++;
                                rowSpanAry[j] = parseInt(rowSpan) - 1;
                                tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' rowspan ='" +
                                    parseInt(rowSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                            } else if (colSpan !== undefined) {
                                tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' colspan = '" +
                                    parseInt(colSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                totalColSpan = parseInt(colSpan) - 1;
                            } else {
                                tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' style = '" + colStyl + "'>" + text + "</td>";
                            }

                        } else {
                            if (rowSpanAry[j] != 0) {
                                rowSpanAry[j] -= 1;
                            }
                            if (totalColSpan != 0) {
                                totalColSpan--;
                            }
                        }
                        j++;
                    }
                } else {
                    //single column 

                    var a_sorce: any;
                    if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                        a_sorce = "a:firstCol";

                    } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                        var aBandNode: any = getTextByPathList(thisTblStyle, ["a:band2V"]);
                        if (aBandNode === undefined) {
                            aBandNode = getTextByPathList(thisTblStyle, ["a:band1V"]);
                            if (aBandNode !== undefined) {
                                a_sorce = "a:band2V";
                            }
                        } else {
                            a_sorce = "a:band2V";
                        }
                    }

                    if (tblStylAttrObj["isLstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                        a_sorce = "a:lastCol";
                    }


                    var cellParmAry: any = getTableCellParams(tcNodes, getColsGrid , i , undefined , thisTblStyle, a_sorce, warpObj)
                    var text: any = cellParmAry[0];
                    var colStyl: any = cellParmAry[1];
                    var cssName: any = cellParmAry[2];
                    var rowSpan: any = cellParmAry[3];

                    if (rowSpan !== undefined) {
                        tableHtml += "<td  class='" + cssName + "' rowspan='" + parseInt(rowSpan) + "' style = '" + colStyl + "'>" + text + "</td>";
                    } else {
                        tableHtml += "<td class='" + cssName + "' style='" + colStyl + "'>" + text + "</td>";
                    }
                }
            }
            tableHtml += "</tr>";
        }
        //////////////////////////////////////////////////////////////////////////////////

    tableHtml += "</table>"; // v1 left the table unclosed
    return tableHtml;
}

// ported from legacy/js/pptxjs.js:10057
export function getTableCellParams(tcNodes?: any, getColsGrid?: any, row_idx?: any, col_idx?: any, thisTblStyle?: any, cellSource?: any, warpObj?: any): any {
    //thisTblStyle["a:band1V"] => thisTblStyle[cellSource]
    //text, cell-width, cell-borders, 
    //var text = genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj);//warpObj.ctx.tableStyles
    var rowSpan: any = getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
    var colSpan: any = getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
    var vMerge: any = getTextByPathList(tcNodes, ["attrs", "vMerge"]);
    var hMerge: any = getTextByPathList(tcNodes, ["attrs", "hMerge"]);
    var colStyl: any = "word-wrap: break-word;";
    var colWidth: any;
    var celFillColor: any = "";
    var col_borders: any = "";
    var colFontClrPr: any = "";
    var colFontWeight: any = "";
    var lin_bottm: any = "",
        lin_top: any = "",
        lin_left: any = "",
        lin_right: any = "",
        lin_bottom_left_to_top_right: any = "",
        lin_top_left_to_bottom_right: any = "";
    
    var colSapnInt: any = parseInt(colSpan);
    var total_col_width: any = 0;
    if (!isNaN(colSapnInt) && colSapnInt > 1){
        for (var k: any = 0; k < colSapnInt ; k++) {
            total_col_width += parseInt(getTextByPathList(getColsGrid[col_idx + k], ["attrs", "w"]));
        }
    }else{
        total_col_width = getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
    }
    

    var text: any = genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj, total_col_width);//warpObj.ctx.tableStyles

    if (total_col_width != 0 /*&& row_idx == 0*/) {
        colWidth = parseInt(total_col_width) * slideFactor;
        colStyl += "width:" + colWidth + "px;";
    }

    //cell bords
    lin_bottm = getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
    if (lin_bottm === undefined && cellSource !== undefined) {
        if (cellSource !== undefined)
            lin_bottm = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
        if (lin_bottm === undefined) {
            lin_bottm = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
        }
    }
    lin_top = getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
    if (lin_top === undefined) {
        if (cellSource !== undefined)
            lin_top = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
        if (lin_top === undefined) {
            lin_top = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
        }
    }
    lin_left = getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
    if (lin_left === undefined) {
        if (cellSource !== undefined)
            lin_left = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
        if (lin_left === undefined) {
            lin_left = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
        }
    }
    lin_right = getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
    if (lin_right === undefined) {
        if (cellSource !== undefined)
            lin_right = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
        if (lin_right === undefined) {
            lin_right = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
        }
    }
    lin_bottom_left_to_top_right = getTextByPathList(tcNodes, ["a:tcPr", "a:lnBlToTr"]);
    lin_top_left_to_bottom_right = getTextByPathList(tcNodes, ["a:tcPr", "a:InTlToBr"]);

    if (lin_bottm !== undefined && lin_bottm != "") {
        var bottom_line_border: any = getBorder(lin_bottm, undefined, false, "", warpObj)
        if (bottom_line_border != "") {
            colStyl += "border-bottom:" + bottom_line_border + ";";
        }
    }
    if (lin_top !== undefined && lin_top != "") {
        var top_line_border: any = getBorder(lin_top, undefined, false, "", warpObj);
        if (top_line_border != "") {
            colStyl += "border-top: " + top_line_border + ";";
        }
    }
    if (lin_left !== undefined && lin_left != "") {
        var left_line_border: any = getBorder(lin_left, undefined, false, "", warpObj)
        if (left_line_border != "") {
            colStyl += "border-left: " + left_line_border + ";";
        }
    }
    if (lin_right !== undefined && lin_right != "") {
        var right_line_border: any = getBorder(lin_right, undefined, false, "", warpObj)
        if (right_line_border != "") {
            colStyl += "border-right:" + right_line_border + ";";
        }
    }

    //cell fill color custom
    var getCelFill: any = getTextByPathList(tcNodes, ["a:tcPr"]);
    if (getCelFill !== undefined && getCelFill != "") {
        var cellObj: any = {
            "p:spPr": getCelFill
        };
        celFillColor = getShapeFill(cellObj, undefined, false, warpObj, "slide")
    }

    //cell fill color theme
    if (celFillColor == "" || celFillColor == "background-color: inherit;") {
        var bgFillschemeClr: any;
        if (cellSource !== undefined)
            bgFillschemeClr = getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
        if (bgFillschemeClr !== undefined) {
            var local_fillColor: any = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
            if (local_fillColor !== undefined) {
                celFillColor = " background-color: #" + local_fillColor + ";";
            }
        }
    }
    var cssName: any = "";
    if (celFillColor !== undefined && celFillColor != "") {
        if (celFillColor in warpObj.ctx.styleTable) {
            cssName = warpObj.ctx.styleTable[celFillColor]["name"];
        } else {
            cssName = "_tbl_cell_css_" + (Object.keys(warpObj.ctx.styleTable).length + 1);
            warpObj.ctx.styleTable[celFillColor] = {
                "name": cssName,
                "text": celFillColor
            };
        }

    }

    //border
    // var borderStyl = getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:tcBdr"]);
    // if (borderStyl !== undefined) {
    //     var local_col_borders = getTableBorders(borderStyl, warpObj);
    //     if (local_col_borders != "") {
    //         col_borders = local_col_borders;
    //     }
    // }
    // if (col_borders != "") {
    //     colStyl += col_borders;
    // }

    //Text style
    var rowTxtStyl: any;
    if (cellSource !== undefined) {
        rowTxtStyl = getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
    }
    // if (rowTxtStyl === undefined) {
    //     rowTxtStyl = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
    // }
    if (rowTxtStyl !== undefined) {
        var local_fontClrPr: any = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
        if (local_fontClrPr !== undefined) {
            colFontClrPr = local_fontClrPr;
        }
        var local_fontWeight: any = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
        if (local_fontWeight !== "") {
            colFontWeight = local_fontWeight;
        }
    }
    colStyl += ((colFontClrPr !== "") ? "color: #" + colFontClrPr + ";" : "");
    colStyl += ((colFontWeight != "") ? " font-weight:" + colFontWeight + ";" : "");

    return [text, colStyl, cssName, rowSpan, colSpan];
}

// ported from legacy/js/pptxjs.js:11356
export function getTableBorders(node?: any, warpObj?: any): any {
    var borderStyle: any = "";
    if (node["a:bottom"] !== undefined) {
        var obj: any = {
            "p:spPr": {
                "a:ln": node["a:bottom"]["a:ln"]
            }
        }
        var borders: any = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-bottom");
    }
    if (node["a:top"] !== undefined) {
        var obj: any = {
            "p:spPr": {
                "a:ln": node["a:top"]["a:ln"]
            }
        }
        var borders: any = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-top");
    }
    if (node["a:right"] !== undefined) {
        var obj: any = {
            "p:spPr": {
                "a:ln": node["a:right"]["a:ln"]
            }
        }
        var borders: any = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-right");
    }
    if (node["a:left"] !== undefined) {
        var obj: any = {
            "p:spPr": {
                "a:ln": node["a:left"]["a:ln"]
            }
        }
        var borders: any = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-left");
    }

    return borderStyle;
}
