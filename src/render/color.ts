// Ported from PPTXjs v1 (legacy/js/pptxjs.js). Internal rendering code: the
// node objects come from the simplified XML tree, so most values are `any`.
import type { WarpObj } from "./context";
import { getTextByPathList } from "../core/xml";
import { tinycolor } from "../utils/tinycolor";
// ported from legacy/js/pptxjs.js:12688
export function getSolidFill(node?: any, clrMap?: any, phClr?: any, warpObj?: any): any {

    if (node === undefined) {
        return undefined;
    }

    //console.log("getSolidFill node: ", node)
    var color: any = "";
    var clrNode: any;
    if (node["a:srgbClr"] !== undefined) {
        clrNode = node["a:srgbClr"];
        color = getTextByPathList(clrNode, ["attrs", "val"]); //#...
    } else if (node["a:schemeClr"] !== undefined) { //a:schemeClr
        clrNode = node["a:schemeClr"];
        var schemeClr: any = getTextByPathList(clrNode, ["attrs", "val"]);
        color = getSchemeColorFromTheme("a:" + schemeClr, clrMap, phClr, warpObj);
        //console.log("schemeClr: ", schemeClr, "color: ", color)
    } else if (node["a:scrgbClr"] !== undefined) {
        clrNode = node["a:scrgbClr"];
        //<a:scrgbClr r="50%" g="50%" b="50%"/>  //Need to test/////////////////////////////////////////////
        var defBultColorVals: any = clrNode["attrs"];
        var red: any = (defBultColorVals["r"].indexOf("%") != -1) ? defBultColorVals["r"].split("%").shift() : defBultColorVals["r"];
        var green: any = (defBultColorVals["g"].indexOf("%") != -1) ? defBultColorVals["g"].split("%").shift() : defBultColorVals["g"];
        var blue: any = (defBultColorVals["b"].indexOf("%") != -1) ? defBultColorVals["b"].split("%").shift() : defBultColorVals["b"];
        //var scrgbClr = red + "," + green + "," + blue;
        color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100));
        //console.log("scrgbClr: " + scrgbClr);

    } else if (node["a:prstClr"] !== undefined) {
        clrNode = node["a:prstClr"];
        //<a:prstClr val="black"/>  //Need to test/////////////////////////////////////////////
        var prstClr: any = getTextByPathList(clrNode, ["attrs", "val"]); //node["a:prstClr"]["attrs"]["val"];
        color = getColorName2Hex(prstClr);
        //console.log("blip prstClr: ", prstClr, " => hexClr: ", color);
    } else if (node["a:hslClr"] !== undefined) {
        clrNode = node["a:hslClr"];
        //<a:hslClr hue="14400000" sat="100%" lum="50%"/>  //Need to test/////////////////////////////////////////////
        var defBultColorVals: any = clrNode["attrs"];
        var hue: any = Number(defBultColorVals["hue"]) / 100000;
        var sat: any = Number((defBultColorVals["sat"].indexOf("%") != -1) ? defBultColorVals["sat"].split("%").shift() : defBultColorVals["sat"]) / 100;
        var lum: any = Number((defBultColorVals["lum"].indexOf("%") != -1) ? defBultColorVals["lum"].split("%").shift() : defBultColorVals["lum"]) / 100;
        //var hslClr = defBultColorVals["hue"] + "," + defBultColorVals["sat"] + "," + defBultColorVals["lum"];
        var hsl2rgb: any = hslToRgb(hue, sat, lum);
        color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
        //defBultColor = cnvrtHslColor2Hex(hslClr); //TODO
        // console.log("hslClr: " + hslClr);
    } else if (node["a:sysClr"] !== undefined) {
        clrNode = node["a:sysClr"];
        //<a:sysClr val="windowText" lastClr="000000"/>  //Need to test/////////////////////////////////////////////
        var sysClr: any = getTextByPathList(clrNode, ["attrs", "lastClr"]);
        if (sysClr !== undefined) {
            color = sysClr;
        }
    }
    //console.log("color: [%cstart]", "color: #" + color, tinycolor(color).toHslString(), color)

    //fix color -------------------------------------------------------- TODO 
    //
    //1. "alpha":
    //Specifies the opacity as expressed by a percentage value.
    // [Example: The following represents a green solid fill which is 50 % opaque
    // < a: solidFill >
    //     <a:srgbClr val="00FF00">
    //         <a:alpha val="50%" />
    //     </a:srgbClr>
    // </a: solidFill >
    var isAlpha: any = false;
    var alpha: any = parseInt(getTextByPathList(clrNode, ["a:alpha", "attrs", "val"])) / 100000;
    //console.log("alpha: ", alpha)
    if (!isNaN(alpha)) {
        // var al_color = new colz.Color(color);
        // al_color.setAlpha(alpha);
        // var ne_color = al_color.rgba.toString();
        // color = (rgba2hex(ne_color))
        var al_color: any = tinycolor(color);
        al_color.setAlpha(alpha);
        color = al_color.toHex8()
        isAlpha = true;
        //console.log("al_color: ", al_color, ", color: ", color)
    }
    //2. "alphaMod":
    // Specifies the opacity as expressed by a percentage relative to the input color.
    //     [Example: The following represents a green solid fill which is 50 % opaque
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:alphaMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //3. "alphaOff":
    // Specifies the opacity as expressed by a percentage offset increase or decrease to the
    // input color.Increases never increase the opacity beyond 100 %, decreases never decrease
    // the opacity below 0 %.
    // [Example: The following represents a green solid fill which is 90 % opaque
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:alphaOff val="-10%" />
    //         </a:srgbClr>
    //     </a: solidFill >

    //4. "blue":
    //Specifies the value of the blue component.The assigned value is specified as a
    //percentage with 0 % indicating minimal blue and 100 % indicating maximum blue.
    //  [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //      to value RRGGBB = (00, FF, FF)
    //          <a: solidFill >
    //              <a:srgbClr val="00FF00">
    //                  <a:blue val="100%" />
    //              </a:srgbClr>
    //          </a: solidFill >
    //5. "blueMod"
    // Specifies the blue component as expressed by a percentage relative to the input color
    // component.Increases never increase the blue component beyond 100 %, decreases
    // never decrease the blue component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    //     to value RRGGBB = (00, 00, 80)
    //     < a: solidFill >
    //         <a:srgbClr val="0000FF">
    //             <a:blueMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //6. "blueOff"
    // Specifies the blue component as expressed by a percentage offset increase or decrease
    // to the input color component.Increases never increase the blue component
    // beyond 100 %, decreases never decrease the blue component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    // to value RRGGBB = (00, 00, CC)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:blueOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >

    //7. "comp" - This element specifies that the color rendered should be the complement of its input color with the complement
    // being defined as such.Two colors are called complementary if, when mixed they produce a shade of grey.For
    // instance, the complement of red which is RGB(255, 0, 0) is cyan.(<a:comp/>)

    //8. "gamma" - This element specifies that the output color rendered by the generating application should be the sRGB gamma
    //              shift of the input color.

    //9. "gray" - This element specifies a grayscale of its input color, taking into relative intensities of the red, green, and blue
    //              primaries.

    //10. "green":
    // Specifies the value of the green component. The assigned value is specified as a
    // percentage with 0 % indicating minimal green and 100 % indicating maximum green.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    // to value RRGGBB = (00, FF, FF)
    //     < a: solidFill >
    //         <a:srgbClr val="0000FF">
    //             <a:green val="100%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //11. "greenMod":
    // Specifies the green component as expressed by a percentage relative to the input color
    // component.Increases never increase the green component beyond 100 %, decreases
    // never decrease the green component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    // to value RRGGBB = (00, 80, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:greenMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //12. "greenOff":
    // Specifies the green component as expressed by a percentage offset increase or decrease
    // to the input color component.Increases never increase the green component
    // beyond 100 %, decreases never decrease the green component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    // to value RRGGBB = (00, CC, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:greenOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >

    //13. "hue" (This element specifies a color using the HSL color model):
    // This element specifies the input color with the specified hue, but with its saturation and luminance unchanged.
    // < a: solidFill >
    //     <a:hslClr hue="14400000" sat="100%" lum="50%">
    // </a:solidFill>
    // <a:solidFill>
    //     <a:hslClr hue="0" sat="100%" lum="50%">
    //         <a:hue val="14400000"/>
    //     <a:hslClr/>
    // </a:solidFill>

    //14. "hueMod" (This element specifies a color using the HSL color model):
    // Specifies the hue as expressed by a percentage relative to the input color.
    // [Example: The following manipulates the fill color from having RGB value RRGGBB = (00, FF, 00) to value RRGGBB = (FF, FF, 00)
    //         < a: solidFill >
    //             <a:srgbClr val="00FF00">
    //                 <a:hueMod val="50%" />
    //             </a:srgbClr>
    //         </a: solidFill >

    var hueMod: any = parseInt(getTextByPathList(clrNode, ["a:hueMod", "attrs", "val"])) / 100000;
    //console.log("hueMod: ", hueMod)
    if (!isNaN(hueMod)) {
        color = applyHueMod(color, hueMod, isAlpha);
    }
    //15. "hueOff"(This element specifies a color using the HSL color model):
    // Specifies the actual angular value of the shift.The result of the shift shall be between 0
    // and 360 degrees.Shifts resulting in angular values less than 0 are treated as 0. Shifts
    // resulting in angular values greater than 360 are treated as 360.
    // [Example:
    //     The following increases the hue angular value by 10 degrees.
    //     < a: solidFill >
    //         <a:hslClr hue="0" sat="100%" lum="50%"/>
    //             <a:hueOff val="600000"/>
    //     </a: solidFill >
    //var hueOff = parseInt(getTextByPathList(clrNode, ["a:hueOff", "attrs", "val"])) / 100000;
    // if (!isNaN(hueOff)) {
    //     //console.log("hueOff: ", hueOff, " (TODO)")
    //     //color = applyHueOff(color, hueOff, isAlpha);
    // }

    //16. "inv" (inverse)
    //This element specifies the inverse of its input color.
    //The inverse of red (1, 0, 0) is cyan (0, 1, 1 ).
    // The following represents cyan, the inverse of red:
    // <a:solidFill>
    //     <a:srgbClr val="FF0000">
    //         <a:inv />
    //     </a:srgbClr>
    // </a:solidFill>

    //17. "invGamma" - This element specifies that the output color rendered by the generating application should be the inverse sRGB
    //                  gamma shift of the input color.

    //18. "lum":
    // This element specifies the input color with the specified luminance, but with its hue and saturation unchanged.
    // Typically luminance values fall in the range[0 %, 100 %].
    // The following two solid fills are equivalent:
    // <a:solidFill>
    //     <a:hslClr hue="14400000" sat="100%" lum="50%">
    // </a:solidFill>
    // <a:solidFill>
    //     <a:hslClr hue="14400000" sat="100%" lum="0%">
    //         <a:lum val="50%" />
    //     <a:hslClr />
    // </a:solidFill>
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    // to value RRGGBB = (00, 66, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:lum val="20%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    // end example]
    //19. "lumMod":
    // Specifies the luminance as expressed by a percentage relative to the input color.
    // Increases never increase the luminance beyond 100 %, decreases never decrease the
    // luminance below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (00, 75, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:lumMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    // end example]
    var lumMod: any = parseInt(getTextByPathList(clrNode, ["a:lumMod", "attrs", "val"])) / 100000;
    //console.log("lumMod: ", lumMod)
    if (!isNaN(lumMod)) {
        color = applyLumMod(color, lumMod, isAlpha);
    }
    //var lumMod_color = applyLumMod(color, 0.5);
    //console.log("lumMod_color: ", lumMod_color)
    //20. "lumOff"
    // Specifies the luminance as expressed by a percentage offset increase or decrease to the
    // input color.Increases never increase the luminance beyond 100 %, decreases never
    // decrease the luminance below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (00, 99, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:lumOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    var lumOff: any = parseInt(getTextByPathList(clrNode, ["a:lumOff", "attrs", "val"])) / 100000;
    //console.log("lumOff: ", lumOff)
    if (!isNaN(lumOff)) {
        color = applyLumOff(color, lumOff, isAlpha);
    }


    //21. "red":
    // Specifies the value of the red component.The assigned value is specified as a percentage
    // with 0 % indicating minimal red and 100 % indicating maximum red.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (FF, FF, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:red val="100%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //22. "redMod":
    // Specifies the red component as expressed by a percentage relative to the input color
    // component.Increases never increase the red component beyond 100 %, decreases never
    // decrease the red component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
    //     to value RRGGBB = (80, 00, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="FF0000">
    //             <a:redMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //23. "redOff":
    // Specifies the red component as expressed by a percentage offset increase or decrease to
    // the input color component.Increases never increase the red component beyond 100 %,
    //     decreases never decrease the red component below 0 %.
    //     [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
    //     to value RRGGBB = (CC, 00, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="FF0000">
    //             <a:redOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >

    //23. "sat":
    // This element specifies the input color with the specified saturation, but with its hue and luminance unchanged.
    // Typically saturation values fall in the range[0 %, 100 %].
    // [Example:
    //     The following two solid fills are equivalent:
    //     <a:solidFill>
    //         <a:hslClr hue="14400000" sat="100%" lum="50%">
    //     </a:solidFill>
    //     <a:solidFill>
    //         <a:hslClr hue="14400000" sat="0%" lum="50%">
    //             <a:sat val="100000" />
    //         <a:hslClr />
    //     </a:solidFill>
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (40, C0, 40)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:sat val="50%" />
    //         </a:srgbClr>
    //     <a: solidFill >
    // end example]

    //24. "satMod":
    // Specifies the saturation as expressed by a percentage relative to the input color.
    // Increases never increase the saturation beyond 100 %, decreases never decrease the
    // saturation below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (66, 99, 66)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:satMod val="20%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    var satMod: any = parseInt(getTextByPathList(clrNode, ["a:satMod", "attrs", "val"])) / 100000;
    if (!isNaN(satMod)) {
        color = applySatMod(color, satMod, isAlpha);
    }
    //25. "satOff":
    // Specifies the saturation as expressed by a percentage offset increase or decrease to the
    // input color.Increases never increase the saturation beyond 100 %, decreases never
    // decrease the saturation below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (19, E5, 19)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:satOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    // var satOff = parseInt(getTextByPathList(clrNode, ["a:satOff", "attrs", "val"])) / 100000;
    // if (!isNaN(satOff)) {
    //     console.log("satOff: ", satOff, " (TODO)")
    // }

    //26. "shade":
    // This element specifies a darker version of its input color.A 10 % shade is 10 % of the input color combined with 90 % black.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (00, BC, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:shade val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    // end example]
    var shade: any = parseInt(getTextByPathList(clrNode, ["a:shade", "attrs", "val"])) / 100000;
    if (!isNaN(shade)) {
        color = applyShade(color, shade, isAlpha);
    }
    //27.  "tint":
    // This element specifies a lighter version of its input color.A 10 % tint is 10 % of the input color combined with
    // 90 % white.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (BC, FF, BC)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:tint val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    var tint: any = parseInt(getTextByPathList(clrNode, ["a:tint", "attrs", "val"])) / 100000;
    if (!isNaN(tint)) {
        color = applyTint(color, tint, isAlpha);
    }
    //console.log("color [%cfinal]: ", "color: #" + color, tinycolor(color).toHslString(), color)

    return color;
}

// ported from legacy/js/pptxjs.js:13129
export function getSchemeColorFromTheme(schemeClr?: any, clrMap?: any, phClr?: any, warpObj?: any): any {
    //<p:clrMap ...> in slide master
    // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1" slideLayoutClrOvride
    //console.log("getSchemeColorFromTheme: schemeClr: ", schemeClr, ",clrMap: ", clrMap)
    var slideLayoutClrOvride: any;
    if (clrMap !== undefined) {
        slideLayoutClrOvride = clrMap;//getTextByPathList(clrMap, ["p:sldMaster", "p:clrMap", "attrs"])
    } else {
        var sldClrMapOvr: any = getTextByPathList(warpObj["slideContent"], ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            slideLayoutClrOvride = sldClrMapOvr;
        } else {
            var sldClrMapOvr: any = getTextByPathList(warpObj["slideLayoutContent"], ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                slideLayoutClrOvride = sldClrMapOvr;
            } else {
                slideLayoutClrOvride = getTextByPathList(warpObj["slideMasterContent"], ["p:sldMaster", "p:clrMap", "attrs"]);
            }

        }
    }
    //console.log("getSchemeColorFromTheme slideLayoutClrOvride: ", slideLayoutClrOvride);
    var schmClrName: any = schemeClr.substr(2);
    if (schmClrName == "phClr" && phClr !== undefined) {
        color = phClr;
    } else {
        if (slideLayoutClrOvride !== undefined) {
            switch (schmClrName) {
                case "tx1":
                case "tx2":
                case "bg1":
                case "bg2":
                    schemeClr = "a:" + slideLayoutClrOvride[schmClrName];
                    break;
            }
        } else {
            switch (schmClrName) {
                case "tx1":
                    schemeClr = "a:dk1";
                    break;
                case "tx2":
                    schemeClr = "a:dk2";
                    break;
                case "bg1":
                    schemeClr = "a:lt1";
                    break;
                case "bg2":
                    schemeClr = "a:lt2";
                    break;
            }
        }
        //console.log("getSchemeColorFromTheme:  schemeClr: ", schemeClr);
        var refNode: any = getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
        var color: any = getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
        //console.log("themeContent: color", color);
        if (color === undefined && refNode !== undefined) {
            color = getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
        }
    }
    //console.log(color)
    return color;
}

// ported from legacy/js/pptxjs.js:13092
export function toHex(n?: any): any {
    var hex: any = n.toString(16);
    while (hex.length < 2) { hex = "0" + hex; }
    return hex;
}

// ported from legacy/js/pptxjs.js:13097
export function hslToRgb(hue?: any, sat?: any, light?: any): any {
    var t1: any, t2: any, r: any, g: any, b: any;
    hue = hue / 60;
    if (light <= 0.5) {
        t2 = light * (sat + 1);
    } else {
        t2 = light + sat - (light * sat);
    }
    t1 = light * 2 - t2;
    r = hueToRgb(t1, t2, hue + 2) * 255;
    g = hueToRgb(t1, t2, hue) * 255;
    b = hueToRgb(t1, t2, hue - 2) * 255;
    return { r: r, g: g, b: b };
}

// ported from legacy/js/pptxjs.js:13111
export function hueToRgb(t1?: any, t2?: any, hue?: any): any {
    if (hue < 0) hue += 6;
    if (hue >= 6) hue -= 6;
    if (hue < 1) return (t2 - t1) * hue + t1;
    else if (hue < 3) return t2;
    else if (hue < 4) return (t2 - t1) * (4 - hue) + t1;
    else return t1;
}

// ported from legacy/js/pptxjs.js:13119
export function getColorName2Hex(name?: any): any {
    var hex: any;
    var colorName: any = ['white', 'AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk', 'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenRod', 'DarkGray', 'DarkGrey', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray', 'DarkSlateGrey', 'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DimGrey', 'DodgerBlue', 'FireBrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'GoldenRod', 'Gray', 'Grey', 'Green', 'GreenYellow', 'HoneyDew', 'HotPink', 'IndianRed', 'Indigo', 'Ivory', 'Khaki', 'Lavender', 'LavenderBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenRodYellow', 'LightGray', 'LightGrey', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray', 'LightSlateGrey', 'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquaMarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue', 'MintCream', 'MistyRose', 'Moccasin', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenRod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue', 'Purple', 'RebeccaPurple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'SlateGrey', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Turquoise', 'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen'];
    var colorHex: any = ['ffffff', 'f0f8ff', 'faebd7', '00ffff', '7fffd4', 'f0ffff', 'f5f5dc', 'ffe4c4', '000000', 'ffebcd', '0000ff', '8a2be2', 'a52a2a', 'deb887', '5f9ea0', '7fff00', 'd2691e', 'ff7f50', '6495ed', 'fff8dc', 'dc143c', '00ffff', '00008b', '008b8b', 'b8860b', 'a9a9a9', 'a9a9a9', '006400', 'bdb76b', '8b008b', '556b2f', 'ff8c00', '9932cc', '8b0000', 'e9967a', '8fbc8f', '483d8b', '2f4f4f', '2f4f4f', '00ced1', '9400d3', 'ff1493', '00bfff', '696969', '696969', '1e90ff', 'b22222', 'fffaf0', '228b22', 'ff00ff', 'dcdcdc', 'f8f8ff', 'ffd700', 'daa520', '808080', '808080', '008000', 'adff2f', 'f0fff0', 'ff69b4', 'cd5c5c', '4b0082', 'fffff0', 'f0e68c', 'e6e6fa', 'fff0f5', '7cfc00', 'fffacd', 'add8e6', 'f08080', 'e0ffff', 'fafad2', 'd3d3d3', 'd3d3d3', '90ee90', 'ffb6c1', 'ffa07a', '20b2aa', '87cefa', '778899', '778899', 'b0c4de', 'ffffe0', '00ff00', '32cd32', 'faf0e6', 'ff00ff', '800000', '66cdaa', '0000cd', 'ba55d3', '9370db', '3cb371', '7b68ee', '00fa9a', '48d1cc', 'c71585', '191970', 'f5fffa', 'ffe4e1', 'ffe4b5', 'ffdead', '000080', 'fdf5e6', '808000', '6b8e23', 'ffa500', 'ff4500', 'da70d6', 'eee8aa', '98fb98', 'afeeee', 'db7093', 'ffefd5', 'ffdab9', 'cd853f', 'ffc0cb', 'dda0dd', 'b0e0e6', '800080', '663399', 'ff0000', 'bc8f8f', '4169e1', '8b4513', 'fa8072', 'f4a460', '2e8b57', 'fff5ee', 'a0522d', 'c0c0c0', '87ceeb', '6a5acd', '708090', '708090', 'fffafa', '00ff7f', '4682b4', 'd2b48c', '008080', 'd8bfd8', 'ff6347', '40e0d0', 'ee82ee', 'f5deb3', 'ffffff', 'f5f5f5', 'ffff00', '9acd32'];
    var findIndx: any = colorName.indexOf(name);
    if (findIndx != -1) {
        hex = colorHex[findIndx];
    }
    return hex;
}

// ported from legacy/js/pptxjs.js:13347
/**
 * applyShade
 * @param {string} rgbStr
 * @param {number} shadeValue
 */
export function applyShade(rgbStr?: any, shadeValue?: any, isAlpha?: any): any {
    var color: any = tinycolor(rgbStr).toHsl();
    //console.log("applyShade  color: ", color, ", shadeValue: ", shadeValue)
    if (shadeValue >= 1) {
        shadeValue = 1;
    }
    var cacl_l: any = Math.min(color.l * shadeValue, 1);//;color.l * shadeValue + (1 - shadeValue);
    // if (isAlpha)
    //     return color.lighten(tintValue).toHex8();
    // return color.lighten(tintValue).toHex();
    if (isAlpha)
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

// ported from legacy/js/pptxjs.js:13367
/**
 * applyTint
 * @param {string} rgbStr
 * @param {number} tintValue
 */
export function applyTint(rgbStr?: any, tintValue?: any, isAlpha?: any): any {
    var color: any = tinycolor(rgbStr).toHsl();
    //console.log("applyTint  color: ", color, ", tintValue: ", tintValue)
    if (tintValue >= 1) {
        tintValue = 1;
    }
    var cacl_l: any = color.l * tintValue + (1 - tintValue);
    // if (isAlpha)
    //     return color.lighten(tintValue).toHex8();
    // return color.lighten(tintValue).toHex();
    if (isAlpha)
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

// ported from legacy/js/pptxjs.js:13387
/**
 * applyLumOff
 * @param {string} rgbStr
 * @param {number} offset
 */
export function applyLumOff(rgbStr?: any, offset?: any, isAlpha?: any): any {
    var color: any = tinycolor(rgbStr).toHsl();
    //console.log("applyLumOff  color.l: ", color.l, ", offset: ", offset, ", color.l + offset : ", color.l + offset)
    var lum: any = offset + color.l;
    if (lum >= 1) {
        if (isAlpha)
            return tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex8();
        return tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex();
    }
    if (isAlpha)
        return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex();
}

// ported from legacy/js/pptxjs.js:13406
/**
 * applyLumMod
 * @param {string} rgbStr
 * @param {number} multiplier
 */
export function applyLumMod(rgbStr?: any, multiplier?: any, isAlpha?: any): any {
    var color: any = tinycolor(rgbStr).toHsl();
    //console.log("applyLumMod  color.l: ", color.l, ", multiplier: ", multiplier, ", color.l * multiplier : ", color.l * multiplier)
    var cacl_l: any = color.l * multiplier;
    if (cacl_l >= 1) {
        cacl_l = 1;
    }
    if (isAlpha)
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

// ported from legacy/js/pptxjs.js:13424
export function applyHueMod(rgbStr?: any, multiplier?: any, isAlpha?: any): any {
    var color: any = tinycolor(rgbStr).toHsl();
    //console.log("applyLumMod  color.h: ", color.h, ", multiplier: ", multiplier, ", color.h * multiplier : ", color.h * multiplier)

    var cacl_h: any = color.h * multiplier;
    if (cacl_h >= 360) {
        cacl_h = cacl_h - 360;
    }
    if (isAlpha)
        return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
    return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
}

// ported from legacy/js/pptxjs.js:13460
export function applySatMod(rgbStr?: any, multiplier?: any, isAlpha?: any): any {
    var color: any = tinycolor(rgbStr).toHsl();
    //console.log("applySatMod  color.s: ", color.s, ", multiplier: ", multiplier, ", color.s * multiplier : ", color.s * multiplier)
    var cacl_s: any = color.s * multiplier;
    if (cacl_s >= 1) {
        cacl_s = 1;
    }
    //return;
    // if (isAlpha)
    //     return tinycolor(rgbStr).saturate(multiplier * 100).toHex8();
    // return tinycolor(rgbStr).saturate(multiplier * 100).toHex();
    if (isAlpha)
        return tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex();
}

// ported from legacy/js/pptxjs.js:13480
/**
 * rgba2hex
 * @param {string} rgbaStr
 */
export function rgba2hex(rgbaStr?: any): any {
    var a: any,
        rgb: any = rgbaStr.replace(/\s/g, '').match(/^rgba?\((\d+),(\d+),(\d+),?([^,\s)]+)?/i),
        alpha: any = (rgb && rgb[4] || "").trim(),
        hex: any = rgb ?
            (rgb[1] | 1 << 8).toString(16).slice(1) +
            (rgb[2] | 1 << 8).toString(16).slice(1) +
            (rgb[3] | 1 << 8).toString(16).slice(1) : rgbaStr;

    if (alpha !== "") {
        a = alpha;
    } else {
        a = 1;
    }
    // multiply before convert to HEX
    a = ((a * 255) | 1 << 8).toString(16).slice(1)
    hex = hex + a;

    return hex;
}
