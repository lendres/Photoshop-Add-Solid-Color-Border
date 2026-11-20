#target photoshop
/*
 * Edit Border Presets.jsx
 *
 * Companion to Add Border.jsx. Opens a window for managing the border width
 * presets in "Add Border Presets.txt" (in the same folder as this script):
 *   - Add, Edit, Delete, Move Up / Move Down
 *   - Widths can be entered in centimeters or inches; they are always saved in
 *     centimeters.
 *   - Each preset is either "Uniform" (same border on all sides) or "Keep aspect
 *     ratio" (saved with the keyword "aspect" after the width).
 *   - Nothing is written until you click Save.
 *
 * Comments in the file are kept: the comment block at the top stays at the top,
 * and comments just above a preset stay with that preset (and are removed if the
 * preset is deleted). Lines that can't be read are kept as comments.
 */

// Must match the file name used in Add Border.jsx.
var PRESETS_FILE_NAME = "Add Border Presets.txt";

// Written at the top of a new presets file.
var DEFAULT_HEADER = [
    "# Border width presets for Add Border.jsx",
    "#",
    "# Keep this file in the same folder as Add Border.jsx.",
    "# One preset per line:   Name = width [uniform|aspect]",
    "#   - The width is the border on EACH side, in CENTIMETERS (use a period or a comma for decimals).",
    "#   - The dialog converts it to inches when the units are set to inches.",
    "#   - uniform: the same border on all four sides (the default if neither word is given).",
    "#   - aspect:  keep the photo's aspect ratio. The width is the smallest border; the border on",
    "#              the short sides of the photo is widened so the bordered image has the same",
    "#              aspect ratio as the photo.",
    "# Lines starting with # and blank lines are ignored. Presets appear in the order listed.",
    ""
];

// Units last used in the Add/Edit dialog during this run.
var gEntryUnit = "cm";

function main() {
    var f;
    try {
        f = new File(File($.fileName).parent.fsName + "/" + PRESETS_FILE_NAME);
    } catch (e) {
        alert("Couldn't work out where this script is, so the presets file can't be found.\n" + e);
        return;
    }

    var data = readPresets(f);
    if (data.error) {
        alert(data.error);
        return;
    }
    if (data.bad.length) {
        alert("Some lines in the presets file couldn't be read:\n\n" + data.bad.join("\n") +
              "\n\nThey'll be kept in the file as comments (starting with #) when you save.");
    }
    showEditor(f, data);
}

// ---------------------------------------------------------------------------
// Main window
// ---------------------------------------------------------------------------

function showEditor(f, data) {
    var presets = data.presets;
    var dirty = false;

    var w = new Window("dialog", "Edit Border Presets");
    w.orientation = "column";
    w.alignChildren = "fill";

    w.add("statictext", undefined, "File: " + f.fsName + (f.exists ? "" : "  (new - will be created on Save)"));

    var gMain = w.add("group");
    gMain.alignChildren = "top";

    var list = gMain.add("listbox", undefined, [], {
        numberOfColumns: 4,
        showHeaders: true,
        columnTitles: ["Name", "Centimeters", "Inches", "Border shape"],
        columnWidths: [280, 85, 70, 120]
    });
    list.preferredSize = [580, 260];

    var gSide = gMain.add("group");
    gSide.orientation = "column";
    gSide.alignChildren = "fill";
    var btnAdd = gSide.add("button", undefined, "Add...");
    var btnEdit = gSide.add("button", undefined, "Edit...");
    var btnDelete = gSide.add("button", undefined, "Delete");
    gSide.add("panel").preferredSize.height = 2;
    var btnUp = gSide.add("button", undefined, "Move Up");
    var btnDown = gSide.add("button", undefined, "Move Down");

    w.add("statictext", undefined,
          "Widths are the border on each side. They're saved in centimeters.");

    var gBtn = w.add("group");
    gBtn.alignment = "right";
    var btnCancel = gBtn.add("button", undefined, "Cancel", { name: "cancel" });
    var btnSave = gBtn.add("button", undefined, "Save", { name: "ok" });

    function refresh(selectIndex) {
        list.removeAll();
        for (var i = 0; i < presets.length; i++) {
            var item = list.add("item", presets[i].name);
            item.subItems[0].text = shortNum(presets[i].cm, 4);
            item.subItems[1].text = shortNum(presets[i].cm / 2.54, 3);
            item.subItems[2].text = presets[i].aspect ? "Keep aspect ratio" : "Uniform";
        }
        if (presets.length) {
            if (selectIndex === undefined || selectIndex === null) selectIndex = 0;
            list.selection = Math.max(0, Math.min(selectIndex, presets.length - 1));
        }
        updateButtons();
    }

    function selectedIndex() {
        return list.selection ? list.selection.index : -1;
    }

    function updateButtons() {
        var i = selectedIndex();
        btnEdit.enabled = i >= 0;
        btnDelete.enabled = i >= 0;
        btnUp.enabled = i > 0;
        btnDown.enabled = i >= 0 && i < presets.length - 1;
    }

    list.onChange = updateButtons;
    list.onDoubleClick = function () { btnEdit.onClick(); };

    btnAdd.onClick = function () {
        var p = presetDialog("Add Preset", null);
        if (!p) return;
        var at = selectedIndex() + 1; // insert after the selection (or at the end)
        if (at <= 0) at = presets.length;
        p.before = [];
        presets.splice(at, 0, p);
        dirty = true;
        refresh(at);
    };

    btnEdit.onClick = function () {
        var i = selectedIndex();
        if (i < 0) return;
        var p = presetDialog("Edit Preset", presets[i]);
        if (!p) return;
        presets[i].name = p.name;
        presets[i].cm = p.cm;
        presets[i].aspect = p.aspect;
        dirty = true;
        refresh(i);
    };

    btnDelete.onClick = function () {
        var i = selectedIndex();
        if (i < 0) return;
        presets.splice(i, 1);
        dirty = true;
        refresh(i);
    };

    btnUp.onClick = function () { move(-1); };
    btnDown.onClick = function () { move(1); };

    function move(step) {
        var i = selectedIndex();
        var j = i + step;
        if (i < 0 || j < 0 || j >= presets.length) return;
        var t = presets[i];
        presets[i] = presets[j];
        presets[j] = t;
        dirty = true;
        refresh(j);
    }

    btnSave.onClick = function () {
        if (savePresets(f, data)) w.close(1);
    };

    btnCancel.onClick = function () {
        if (dirty && !confirm("Discard your changes to the presets?", true, "Edit Border Presets")) return;
        w.close(2);
    };

    refresh(0);
    w.show();
}

// Add/Edit dialog. Returns { name, cm, aspect } or null if cancelled.
function presetDialog(title, preset) {
    var d = new Window("dialog", title);
    d.orientation = "column";
    d.alignChildren = "left";

    var gName = d.add("group");
    gName.add("statictext", undefined, "Name:").preferredSize.width = 45;
    var etName = gName.add("edittext", undefined, preset ? preset.name : "");
    etName.characters = 36;

    var gWidth = d.add("group");
    gWidth.add("statictext", undefined, "Width:").preferredSize.width = 45;
    var etWidth = gWidth.add("edittext", undefined, "");
    etWidth.characters = 8;
    var ddUnit = gWidth.add("dropdownlist", undefined, ["cm", "in"]);
    ddUnit.selection = (gEntryUnit === "in") ? 1 : 0;
    gWidth.add("statictext", undefined, "(border on each side)");

    var gShape = d.add("group");
    gShape.add("statictext", undefined, "Shape:").preferredSize.width = 45;
    var ddShape = gShape.add("dropdownlist", undefined, ["Uniform", "Keep aspect ratio"]);
    ddShape.selection = (preset && preset.aspect) ? 1 : 0;
    ddShape.helpTip = "Keep aspect ratio: the width is the smallest border; the border on the " +
                      "photo's short sides is widened so the bordered image keeps the photo's aspect ratio.";

    if (preset) {
        var inches = ddUnit.selection.index === 1;
        etWidth.text = shortNum(inches ? preset.cm / 2.54 : preset.cm, 4);
    }

    // Convert the typed value when the units change.
    var lastUnit = ddUnit.selection.index;
    ddUnit.onChange = function () {
        var v = parseNum(etWidth.text);
        if (v > 0 && ddUnit.selection.index !== lastUnit) {
            v = (ddUnit.selection.index === 1) ? v / 2.54 : v * 2.54;
            etWidth.text = shortNum(v, 4);
        }
        lastUnit = ddUnit.selection.index;
    };

    var gBtn = d.add("group");
    gBtn.alignment = "right";
    var btnCancel = gBtn.add("button", undefined, "Cancel", { name: "cancel" });
    var btnOK = gBtn.add("button", undefined, "OK", { name: "ok" });

    var result = null;
    btnOK.onClick = function () {
        var name = etName.text.replace(/[\r\n]+/g, " ").replace(/^\s+|\s+$/g, "");
        if (!name) {
            alert("Enter a name for the preset.");
            return;
        }
        if (name.charAt(0) === "#") {
            alert("A preset name can't start with #, because lines starting with # are comments.");
            return;
        }
        var v = parseNum(etWidth.text);
        if (!(v > 0)) {
            alert("Enter a width greater than zero.");
            return;
        }
        var inches = ddUnit.selection.index === 1;
        gEntryUnit = inches ? "in" : "cm";
        result = { name: name, cm: inches ? v * 2.54 : v, aspect: ddShape.selection.index === 1 };
        d.close(1);
    };
    btnCancel.onClick = function () { d.close(2); };

    etName.active = true;
    d.show();
    return result;
}

// ---------------------------------------------------------------------------
// Reading and writing the presets file
// ---------------------------------------------------------------------------

// Returns { header: [lines], presets: [{name, cm, before: [lines]}], trailer: [lines],
//           bad: [descriptions], error: text or null }.
function readPresets(f) {
    var data = { header: [], presets: [], trailer: [], bad: [], error: null };
    if (!f.exists) {
        data.header = DEFAULT_HEADER.slice(0);
        return data;
    }

    f.encoding = "UTF-8";
    if (!f.open("r")) {
        data.error = "The presets file couldn't be opened:\n" + f.fsName;
        return data;
    }

    var pending = [];
    var lineNo = 0;
    while (!f.eof) {
        var raw = f.readln().replace(/\r$/, "");
        lineNo++;
        if (lineNo === 1) raw = raw.replace(/^﻿/, "");
        var line = raw.replace(/^\s+|\s+$/g, "");

        if (line === "" || line.charAt(0) === "#") {
            pending.push(raw);
            continue;
        }

        var eq = line.lastIndexOf("=");
        var name = eq > 0 ? line.substr(0, eq).replace(/\s+$/, "") : "";
        var value = eq > 0 ? parsePresetValue(line.substr(eq + 1)) : null;
        if (!name || !value) {
            data.bad.push("Line " + lineNo + ": " + line);
            pending.push("# " + raw);
            continue;
        }

        var p = { name: name, cm: value.cm, aspect: value.aspect, before: [] };
        if (data.presets.length === 0) data.header = pending;
        else p.before = pending;
        pending = [];
        data.presets.push(p);
    }
    f.close();

    if (data.presets.length === 0) data.header = pending;
    else data.trailer = pending;
    return data;
}

function buildLines(data) {
    var lines = data.header.slice(0);
    for (var i = 0; i < data.presets.length; i++) {
        var p = data.presets[i];
        lines = lines.concat(p.before || []);
        lines.push(p.name + " = " + shortNum(p.cm, 4) + (p.aspect ? " aspect" : " uniform"));
    }
    return lines.concat(data.trailer);
}

function writeLines(f, lines) {
    f.encoding = "UTF-8";
    f.lineFeed = ($.os.indexOf("Windows") >= 0) ? "Windows" : "Unix";
    if (!f.open("w")) return false;
    var ok = true;
    for (var i = 0; i < lines.length; i++) {
        if (!f.writeln(lines[i])) { ok = false; break; }
    }
    f.close();
    return ok;
}

// Saves to the presets file. If that fails (for example, the Scripts folder
// needs administrator rights), offers to save a copy somewhere else.
function savePresets(f, data) {
    var lines = buildLines(data);
    if (writeLines(f, lines)) return true;

    var msg = "The presets couldn't be saved to:\n" + f.fsName + "\n\n" +
              "If this is Photoshop's Presets\\Scripts folder, saving there needs administrator " +
              "rights. Save a copy somewhere else now, then copy it into that folder " +
              "(replacing the old file).\n\nSave a copy somewhere else?";
    if (!confirm(msg, false, "Edit Border Presets")) return false;

    var copy = File.saveDialog("Save a copy of the presets", "Text files:*.txt");
    if (!copy) return false;
    if (writeLines(copy, lines)) {
        alert("Saved a copy to:\n" + copy.fsName + "\n\nCopy it to:\n" + f.fsName +
              "\nnamed \"" + PRESETS_FILE_NAME + "\".");
        return true;
    }
    alert("The copy couldn't be saved either:\n" + copy.fsName);
    return false;
}

// ---------------------------------------------------------------------------
// Helpers (same as in Add Border.jsx)
// ---------------------------------------------------------------------------

function parseNum(t) {
    t = String(t).replace(/^\s+|\s+$/g, "").replace(",", ".");
    if (!/^\d*\.?\d+$/.test(t) && !/^\d+\.$/.test(t)) return NaN;
    return parseFloat(t);
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

main();
