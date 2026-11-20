#target photoshop
/*
 * Add Border.jsx
 *
 * Adds a solid black or white border AROUND the active photo:
 *   1. One dialog asks for units, final photo size (W x H), border width (typed in or
 *      chosen from presets in "Add Border Presets.txt"), border shape (uniform or
 *      keep aspect ratio), and border color (black, white or a custom color).
 *   2. Reads the photo's current pixel size.
 *   3. Computes pixels-per-unit for width and height. If they disagree by more than
 *      MAX_RELATIVE_ERROR, asks whether to use width, height, or quit.
 *   4. Converts the border width to pixels. With "Keep aspect ratio", that width is
 *      used on the photo's long sides (top/bottom of a landscape photo) and the border
 *      on the short sides is widened so the bordered image has the same aspect ratio
 *      as the photo. (For printers that crop uploads to the ordered aspect ratio.)
 *   5. Expands the canvas by the border on every side as blank, transparent canvas
 *      (no fill, Content-Aware or Generative Expand; photo pixels untouched).
 *   6. Saves the original photo area as a selection named "No Border Crop", so the
 *      original crop can be restored later (Select > Load Selection, then Image > Crop).
 *   7. Adds a new top layer "Border" and fills only the new outer ring.
 *
 * No message is shown when it finishes; dialogs appear only for the width/height
 * question and for errors.
 *
 * The photo's pixels, the document resolution and all other layers are left unchanged.
 * A locked Background layer is turned into a normal layer (same pixels, same name),
 * because a Background layer can't have transparent canvas around it.
 * All changes are one history step, so a single Undo removes them.
 */

// Allowed relative difference between the width and height conversion rates.
// 0.002 = 0.2%. Change this if it proves too strict or too loose.
var MAX_RELATIVE_ERROR = 0.002;

// Name given to the new layer.
var BORDER_LAYER_NAME = "Border";

// Name of the saved selection (alpha channel) that marks the original photo area.
var CROP_SELECTION_NAME = "No Border Crop";

// Border width presets are read from this file, in the same folder as this script.
// Each line is "Name = width in centimeters [uniform|aspect]"; lines starting with #
// are comments. "aspect" means keep the aspect ratio; without a keyword it's uniform.
var PRESETS_FILE_NAME = "Add Border Presets.txt";

// Used to pass values into suspendHistory (which evaluates in global scope).
var gBorderJob = null;

function main() {
    if (app.documents.length === 0) {
        alert("Open a photo before running Add Border.");
        return;
    }
    var doc = app.activeDocument;

    if (doc.mode == DocumentMode.BITMAP || doc.mode == DocumentMode.INDEXEDCOLOR ||
        doc.mode == DocumentMode.DUOTONE || doc.mode == DocumentMode.MULTICHANNEL) {
        alert("Add Border needs an RGB, CMYK, Grayscale or Lab image.\n" +
              "This document's color mode doesn't support adding layers this way.");
        return;
    }

    var savedRulerUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    try {
        run(doc);
    } catch (e) {
        alert("Add Border stopped because of an error:\n" + e);
    } finally {
        app.preferences.rulerUnits = savedRulerUnits;
    }
}

function run(doc) {
    // ---- 1. Ask for settings ----------------------------------------------
    var s = showSettingsDialog();
    if (!s) return; // cancelled

    // ---- 2. Read the current photo -----------------------------------------
    var wPx = Math.round(doc.width.as("px"));
    var hPx = Math.round(doc.height.as("px"));

    // ---- 3. Conversion rates (pixels per chosen unit) ----------------------
    var rateW = wPx / s.width;
    var rateH = hPx / s.height;
    var relErr = Math.abs(rateW - rateH) / ((rateW + rateH) / 2);

    var basis = "both";
    if (relErr > MAX_RELATIVE_ERROR) {
        basis = askMismatch(s, wPx, hPx, rateW, rateH, relErr);
        if (!basis) return; // quit
    }

    var rate;
    if (basis === "width") rate = rateW;
    else if (basis === "height") rate = rateH;
    else rate = (rateW + rateH) / 2;

    // ---- 4. Border size in pixels ------------------------------------------
    var borderExact = s.border * rate;
    var borderPx = Math.round(borderExact);
    if (borderPx < 1) {
        alert("A border of " + s.border + " " + s.unitShort +
              " works out to less than 1 pixel at this size. Nothing was changed.");
        return;
    }

    // Left/right border (borderX) and top/bottom border (borderY).
    // Keep aspect ratio: the entered width is the smallest border; the border added
    // along the longer dimension is scaled so (W + 2*bx) / (H + 2*by) = W / H.
    var borderX = borderPx;
    var borderY = borderPx;
    if (s.keepAspect) {
        if (wPx > hPx) borderX = Math.round(borderExact * wPx / hPx);
        else if (hPx > wPx) borderY = Math.round(borderExact * hPx / wPx);
    }

    // ---- 5 & 6. Expand canvas, add top layer, fill the ring ----------------
    gBorderJob = {
        doc: doc,
        wPx: wPx,
        hPx: hPx,
        borderX: borderX,
        borderY: borderY,
        color: s.color
    };
    doc.suspendHistory("Add Border", "applyBorder()");
}

// Runs inside one history state.
function applyBorder() {
    var j = gBorderJob;
    var doc = j.doc;
    var bx = j.borderX;
    var by = j.borderY;
    var newW = j.wPx + 2 * bx;
    var newH = j.hPx + 2 * by;

    // A locked Background layer can't hold transparency, so Photoshop would fill
    // the added canvas on it with a color. Turn it into a normal layer first
    // (same pixels, same name) so the added canvas is blank and transparent.
    unlockBackgroundLayer(doc);

    // Grow the canvas equally on all sides. This is a plain Canvas Size: the new
    // area is left empty (transparent) - no fill, Content-Aware or Generative
    // Expand. Existing pixels are not touched.
    doc.resizeCanvas(UnitValue(newW, "px"), UnitValue(newH, "px"), AnchorPosition.MIDDLECENTER);

    // Save the original photo area as a named selection so the original crop
    // can be restored later: Select > Load Selection, then Image > Crop.
    var inner = [[bx, by], [bx + j.wPx, by], [bx + j.wPx, by + j.hPx], [bx, by + j.hPx]];
    doc.selection.select(inner, SelectionType.REPLACE, 0, false);
    saveSelectionAsChannel(CROP_SELECTION_NAME);

    // New empty layer, placed at the very top of the layer stack.
    doc.activeLayer = doc.layers[0];
    var layer = doc.artLayers.add();
    layer.name = BORDER_LAYER_NAME;
    try {
        layer.move(doc, ElementPlacement.PLACEATBEGINNING);
    } catch (e1) {
        try {
            if (doc.layers[0].name !== BORDER_LAYER_NAME) {
                layer.move(doc.layers[0], ElementPlacement.PLACEBEFORE);
            }
        } catch (e2) { /* already on top */ }
    }
    doc.activeLayer = layer;

    // Select the outer ring only (whole canvas minus the original photo area).
    var outer = [[0, 0], [newW, 0], [newW, newH], [0, newH]];
    doc.selection.select(outer, SelectionType.REPLACE, 0, false);
    doc.selection.select(inner, SelectionType.DIMINISH, 0, false);
    doc.selection.fill(j.color, ColorBlendMode.NORMAL, 100, false);
    doc.selection.deselect();
}

// If the bottom layer is a Background layer, make it a normal layer with the
// same name so canvas added around it stays transparent.
function unlockBackgroundLayer(doc) {
    var bottom = doc.layers[doc.layers.length - 1];
    if (bottom.typename === "ArtLayer" && bottom.isBackgroundLayer) {
        var name = bottom.name;
        bottom.isBackgroundLayer = false;
        bottom.name = name;
    }
}

// Same as Select > Save Selection > New Channel with the given name.
// Doesn't change which channels are visible or active.
function saveSelectionAsChannel(name) {
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
    desc.putReference(charIDToTypeID("null"), ref);
    desc.putString(charIDToTypeID("Nm  "), name);
    executeAction(charIDToTypeID("Dplc"), desc, DialogModes.NO);
}

// ---------------------------------------------------------------------------
// Dialogs
// ---------------------------------------------------------------------------

function showSettingsDialog() {
    var presetInfo = loadBorderPresets();
    var presets = presetInfo.presets;
    if (presetInfo.errors.length) {
        alert("Some lines in the presets file couldn't be read and were skipped:\n\n" +
              presetInfo.errors.join("\n") + "\n\nFile: " + presetInfo.path +
              "\n\nEach line should look like:  Name = 5  (width in centimeters)");
    }

    var d = new Window("dialog", "Add Border");
    d.orientation = "column";
    d.alignChildren = "fill";

    var pSize = d.add("panel", undefined, "Final photo size (border is added outside this)");
    pSize.alignChildren = "left";

    var gUnits = pSize.add("group");
    gUnits.add("statictext", undefined, "Units:");
    var ddUnits = gUnits.add("dropdownlist", undefined, ["Inches", "Centimeters"]);
    ddUnits.selection = 0;

    var gWH = pSize.add("group");
    gWH.add("statictext", undefined, "Width:");
    var etW = gWH.add("edittext", undefined, "");
    etW.characters = 7;
    gWH.add("statictext", undefined, "x  Height:");
    var etH = gWH.add("edittext", undefined, "");
    etH.characters = 7;

    var pBorder = d.add("panel", undefined, "Border");
    pBorder.alignChildren = "left";

    // Border width: either typed in (Custom) or chosen from the presets file.
    // The two radio buttons are in different rows, so they're kept exclusive by hand.
    var gB = pBorder.add("group");
    gB.add("statictext", undefined, "Border width:");
    var rbWidthCustom = gB.add("radiobutton", undefined, "Custom:");
    var etB = gB.add("edittext", undefined, "");
    etB.characters = 7;
    var unitLabel = gB.add("statictext", undefined, "in    ");

    var gP = pBorder.add("group");
    gP.margins = [82, 0, 0, 0];
    var rbWidthPreset = gP.add("radiobutton", undefined, "Preset:");
    var ddPreset = gP.add("dropdownlist", undefined, []);
    ddPreset.preferredSize.width = 280;

    function fillPresetList() {
        var keep = ddPreset.selection ? ddPreset.selection.index : 0;
        var inches = ddUnits.selection.index === 0;
        ddPreset.removeAll();
        if (!presets.length) {
            ddPreset.add("item", presetInfo.found ? "No presets in file" : "Presets file not found");
            ddPreset.selection = 0;
            return;
        }
        for (var i = 0; i < presets.length; i++) {
            var v = inches ? presets[i].cm / 2.54 : presets[i].cm;
            ddPreset.add("item", presets[i].name + "  (" + shortNum(v, inches ? 3 : 2) +
                         (inches ? " in" : " cm") +
                         (presets[i].aspect ? ", keep aspect ratio)" : ", uniform)"));
        }
        ddPreset.selection = Math.min(keep, presets.length - 1);
    }

    // Border shape: uniform, or keep the photo's aspect ratio. A preset sets this
    // itself, so the choice is disabled while Preset is selected.
    var gS = pBorder.add("group");
    gS.add("statictext", undefined, "Border shape:");
    var rbUniform = gS.add("radiobutton", undefined, "Uniform");
    var rbAspect = gS.add("radiobutton", undefined, "Keep aspect ratio");
    rbUniform.value = true;
    rbUniform.helpTip = "The same border width on all four sides.";
    rbAspect.helpTip = "The border width is used on the long sides of the photo; the border on the " +
                       "short sides is widened so the bordered image keeps the photo's aspect ratio.";

    function applyPresetShape() {
        if (!rbWidthPreset.value || !presets.length || !ddPreset.selection) return;
        var aspect = presets[ddPreset.selection.index].aspect;
        rbAspect.value = aspect;
        rbUniform.value = !aspect;
    }

    function updateWidthEnabled() {
        etB.enabled = rbWidthCustom.value;
        ddPreset.enabled = rbWidthPreset.value;
        rbUniform.enabled = rbWidthCustom.value;
        rbAspect.enabled = rbWidthCustom.value;
        applyPresetShape();
    }
    rbWidthCustom.onClick = function () { rbWidthPreset.value = false; updateWidthEnabled(); };
    rbWidthPreset.onClick = function () { rbWidthCustom.value = false; updateWidthEnabled(); };
    ddPreset.onChange = applyPresetShape;

    rbWidthCustom.value = true;
    if (!presets.length) {
        rbWidthPreset.enabled = false;
        rbWidthPreset.helpTip = "Add presets to " + presetInfo.path;
    }
    fillPresetList();
    updateWidthEnabled();

    var gC = pBorder.add("group");
    gC.add("statictext", undefined, "Color:");
    var rbBlack = gC.add("radiobutton", undefined, "Black");
    var rbWhite = gC.add("radiobutton", undefined, "White");
    var rbCustom = gC.add("radiobutton", undefined, "Custom:");
    rbBlack.value = true;

    // Custom color: a hex field, a swatch, and a button that opens Photoshop's
    // Color Picker. It starts out as the current foreground color. These
    // controls are only enabled while the Custom radio button is selected.
    var gCustom = pBorder.add("group");
    gCustom.margins = [52, 0, 0, 0];
    var swatch = gCustom.add("group");
    swatch.preferredSize = [36, 20];
    var etHex = gCustom.add("edittext", undefined, colorToHex(app.foregroundColor));
    etHex.characters = 8;
    var btnPick = gCustom.add("button", undefined, "Choose...");

    var lastSwatchRgb = null;
    function updateSwatch() {
        var rgb = hexToRgb(etHex.text) || lastSwatchRgb;
        if (!rgb) return;
        lastSwatchRgb = rgb;
        var r = rgb[0] / 255, g = rgb[1] / 255, b = rgb[2] / 255;
        if (!rbCustom.value) {
            // Show the swatch faded toward the dialog gray while disabled.
            r = (r + 0.5) / 2; g = (g + 0.5) / 2; b = (b + 0.5) / 2;
        }
        try {
            swatch.graphics.backgroundColor = swatch.graphics.newBrush(
                swatch.graphics.BrushType.SOLID_COLOR, [r, g, b, 1]);
        } catch (e) { /* swatch is cosmetic only */ }
    }

    function updateCustomEnabled() {
        var on = rbCustom.value;
        swatch.enabled = on;
        etHex.enabled = on;
        btnPick.enabled = on;
        updateSwatch();
    }
    rbBlack.onClick = updateCustomEnabled;
    rbWhite.onClick = updateCustomEnabled;
    rbCustom.onClick = updateCustomEnabled;
    updateCustomEnabled();

    etHex.onChanging = function () { updateSwatch(); };
    btnPick.onClick = function () {
        var picked = pickColor(hexToRgb(etHex.text));
        if (picked) {
            etHex.text = colorToHex(picked);
            updateSwatch();
        }
    };

    ddUnits.onChange = function () {
        unitLabel.text = (ddUnits.selection.index === 0) ? "in" : "cm";
        fillPresetList();
    };

    var gBtn = d.add("group");
    gBtn.alignment = "right";
    var btnCancel = gBtn.add("button", undefined, "Cancel", { name: "cancel" });
    var btnOK = gBtn.add("button", undefined, "OK", { name: "ok" });

    var result = null;
    btnOK.onClick = function () {
        var w = parseNum(etW.text);
        var h = parseNum(etH.text);
        var b;
        if (rbWidthPreset.value) {
            // Presets are stored in centimeters; convert when working in inches.
            var cm = presets[ddPreset.selection.index].cm;
            b = (ddUnits.selection.index === 0) ? cm / 2.54 : cm;
        } else {
            b = parseNum(etB.text);
        }
        if (!(w > 0) || !(h > 0)) {
            alert("Enter a width and height greater than zero.");
            return;
        }
        if (!(b > 0)) {
            alert("Enter a border width greater than zero.");
            return;
        }
        var inches = ddUnits.selection.index === 0;
        var rgb, colorName;
        if (rbCustom.value) {
            rgb = hexToRgb(etHex.text);
            if (!rgb) {
                alert("Enter the custom color as a hex value such as #1A2B3C, or click Choose.");
                return;
            }
            colorName = colorToHexFromRgb(rgb);
        } else if (rbWhite.value) {
            rgb = [255, 255, 255];
            colorName = "white";
        } else {
            rgb = [0, 0, 0];
            colorName = "black";
        }
        var c = new SolidColor();
        c.rgb.red = rgb[0];
        c.rgb.green = rgb[1];
        c.rgb.blue = rgb[2];
        result = {
            width: w,
            height: h,
            border: b,
            inches: inches,
            unitShort: inches ? "in" : "cm",
            keepAspect: rbAspect.value,
            color: c,
            colorName: colorName
        };
        d.close(1);
    };
    btnCancel.onClick = function () { d.close(2); };

    etW.active = true;
    d.show();
    return result;
}

function askMismatch(s, wPx, hPx, rateW, rateH, relErr) {
    var u = s.unitShort;
    var d = new Window("dialog", "Width and height don't match");
    d.orientation = "column";
    d.alignChildren = "left";

    var lines = [
        "The photo's proportions don't match " + fmt(s.width, 3) + " x " + fmt(s.height, 3) + " " + u + ".",
        "",
        "Photo: " + wPx + " x " + hPx + " px",
        "Width rate:  " + fmt(rateW, 3) + " px/" + u,
        "Height rate: " + fmt(rateH, 3) + " px/" + u,
        "Difference: " + pct(relErr) + "  (allowed: " + pct(MAX_RELATIVE_ERROR) + ")",
        "",
        "If width decides, the height will be " + fmt(hPx / rateW, 3) + " " + u +
            " (" + signed(hPx / rateW - s.height, 3) + ").",
        "If height decides, the width will be " + fmt(wPx / rateH, 3) + " " + u +
            " (" + signed(wPx / rateH - s.width, 3) + ").",
        "",
        "Which size should determine the border?"
    ];
    for (var i = 0; i < lines.length; i++) d.add("statictext", undefined, lines[i] || " ");

    var g = d.add("group");
    g.alignment = "right";
    var bW = g.add("button", undefined, "Use width");
    var bH = g.add("button", undefined, "Use height");
    var bQ = g.add("button", undefined, "Quit", { name: "cancel" });

    var choice = null;
    bW.onClick = function () { choice = "width"; d.close(1); };
    bH.onClick = function () { choice = "height"; d.close(1); };
    bQ.onClick = function () { choice = null; d.close(2); };

    d.show();
    return choice;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function parseNum(t) {
    t = String(t).replace(/^\s+|\s+$/g, "").replace(",", ".");
    if (!/^\d*\.?\d+$/.test(t) && !/^\d+\.$/.test(t)) return NaN;
    return parseFloat(t);
}

// Opens Photoshop's Color Picker starting at the given [r, g, b] (or the current
// foreground color). Returns the chosen SolidColor, or null if cancelled.
// The foreground color is put back afterwards.
function pickColor(startRgb) {
    var savedForeground = app.foregroundColor;
    try {
        if (startRgb) {
            var start = new SolidColor();
            start.rgb.red = startRgb[0];
            start.rgb.green = startRgb[1];
            start.rgb.blue = startRgb[2];
            app.foregroundColor = start;
        }
        if (!app.showColorPicker()) return null;
        return app.foregroundColor;
    } finally {
        app.foregroundColor = savedForeground;
    }
}

// "#1A2B3C", "1a2b3c" or "#abc" -> [26, 43, 60]; null if not valid.
function hexToRgb(text) {
    var t = String(text).replace(/^\s+|\s+$/g, "").replace(/^#/, "");
    if (/^[0-9a-fA-F]{3}$/.test(t)) {
        t = t.charAt(0) + t.charAt(0) + t.charAt(1) + t.charAt(1) + t.charAt(2) + t.charAt(2);
    }
    if (!/^[0-9a-fA-F]{6}$/.test(t)) return null;
    return [parseInt(t.substr(0, 2), 16), parseInt(t.substr(2, 2), 16), parseInt(t.substr(4, 2), 16)];
}

function colorToHex(solidColor) {
    return colorToHexFromRgb([solidColor.rgb.red, solidColor.rgb.green, solidColor.rgb.blue]);
}

function colorToHexFromRgb(rgb) {
    var out = "#";
    for (var i = 0; i < 3; i++) {
        var h = Math.round(rgb[i]).toString(16).toUpperCase();
        out += (h.length < 2 ? "0" : "") + h;
    }
    return out;
}

// Reads border width presets from PRESETS_FILE_NAME next to this script.
// Returns { presets: [{name, cm}], found, path, errors: [text] }.
function loadBorderPresets() {
    var info = { presets: [], found: false, path: PRESETS_FILE_NAME, errors: [] };
    var f;
    try {
        f = new File(File($.fileName).parent.fsName + "/" + PRESETS_FILE_NAME);
    } catch (e) {
        return info;
    }
    info.path = f.fsName;
    if (!f.exists) return info;
    info.found = true;

    f.encoding = "UTF-8";
    if (!f.open("r")) {
        info.errors.push("The file couldn't be opened.");
        return info;
    }
    var lineNo = 0;
    while (!f.eof) {
        var raw = f.readln();
        lineNo++;
        var line = raw.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        if (line === "" || line.charAt(0) === "#") continue;

        var eq = line.lastIndexOf("=");
        var name = eq > 0 ? line.substr(0, eq).replace(/\s+$/, "") : "";
        var value = eq > 0 ? parsePresetValue(line.substr(eq + 1)) : null;
        if (!name || !value) {
            info.errors.push("Line " + lineNo + ": " + line);
            continue;
        }
        info.presets.push({ name: name, cm: value.cm, aspect: value.aspect });
    }
    f.close();
    return info;
}

// The part of a preset line after "=": a width in cm, optionally followed by
// "uniform" (the default) or "aspect". Returns { cm, aspect } or null.
function parsePresetValue(text) {
    var m = String(text).replace(/^\s+|\s+$/g, "").match(/^(\S+)(?:\s+(uniform|aspect))?$/i);
    if (!m) return null;
    var cm = parseNum(m[1]);
    if (!(cm > 0)) return null;
    return { cm: cm, aspect: !!m[2] && m[2].toLowerCase() === "aspect" };
}

// Number with up to d decimals and no trailing zeros, e.g. 1.969, 5, 2.5.
function shortNum(n, d) { return String(Number(Number(n).toFixed(d))); }

function fmt(n, d) { return Number(n).toFixed(d); }

function signed(n, d) {
    var r = Number(n).toFixed(d);
    if (Number(r) === 0) return Number(0).toFixed(d);
    return (n > 0 ? "+" : "") + r;
}

function pct(x) { return (x * 100).toFixed(3) + "%"; }


main();
