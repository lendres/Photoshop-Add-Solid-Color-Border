#target photoshop
/*
 * Remove Border.jsx
 *
 * Companion to Add Border.jsx. Restores the photo's original crop:
 *   1. Finds the saved selection "No Border Crop" made by Add Border.
 *   2. If the top layer is the "Border" layer made by Add Border, asks whether
 *      to delete it.
 *   3. Loads the saved selection and crops the canvas to it, with "Delete Cropped
 *      Pixels" turned off: pixels outside the new canvas are hidden, not deleted,
 *      and Image > Reveal All brings them back.
 *
 * No message is shown when it finishes; dialogs appear only for the delete
 * question and for errors. All changes are one history step, so a single Undo
 * removes them. The saved selection itself is kept.
 */

// These must match the names used in Add Border.jsx.
var BORDER_LAYER_NAME = "Border";
var CROP_SELECTION_NAME = "No Border Crop";

// Used to pass values into suspendHistory (which evaluates in global scope).
var gRestoreJob = null;

function main() {
    if (app.documents.length === 0) {
        alert("Open a photo before running Remove Border.");
        return;
    }
    var doc = app.activeDocument;

    var savedRulerUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    try {
        run(doc);
    } catch (e) {
        alert("Remove Border stopped because of an error:\n" + e);
    } finally {
        app.preferences.rulerUnits = savedRulerUnits;
    }
}

function run(doc) {
    // ---- 1. Find the saved selection ---------------------------------------
    var channel = findCropChannel(doc);
    if (!channel) {
        alert("This document has no saved selection named \"" + CROP_SELECTION_NAME + "\".\n\n" +
              "It's created by Add Border. If the file was saved as JPEG, or the channel " +
              "was deleted or renamed, the original crop can't be restored automatically. " +
              "Nothing was changed.");
        return;
    }

    // ---- 2. Ask about the Border layer -------------------------------------
    var top = doc.layers[0];
    var deleteBorder = false;
    if (top.typename === "ArtLayer" && top.name === BORDER_LAYER_NAME) {
        deleteBorder = confirm("The top layer is \"" + BORDER_LAYER_NAME + "\".\n\n" +
                               "Delete this layer?", false, "Remove Border");
    }

    // ---- 3. Load the selection and crop ------------------------------------
    gRestoreJob = {
        doc: doc,
        channel: channel,
        borderLayer: deleteBorder ? top : null
    };
    doc.suspendHistory("Remove Border", "restoreCrop()");
}

// Runs inside one history state.
function restoreCrop() {
    var j = gRestoreJob;
    var doc = j.doc;

    doc.selection.load(j.channel, SelectionType.REPLACE, false);

    var bounds;
    try {
        bounds = doc.selection.bounds;
    } catch (e) {
        throw new Error("The \"" + CROP_SELECTION_NAME + "\" selection is empty.");
    }

    doc.selection.deselect();
    cropKeepingPixels(bounds[0].as("px"), bounds[1].as("px"),
                      bounds[2].as("px"), bounds[3].as("px"));

    if (j.borderLayer) {
        j.borderLayer.remove();
    }
}

// Crops the canvas to the given rectangle (in pixels) like the Crop tool with
// "Delete Cropped Pixels" turned off, so pixels outside it are kept but hidden.
function cropKeepingPixels(left, top, right, bottom) {
    var c = function (s) { return charIDToTypeID(s); };
    var rect = new ActionDescriptor();
    rect.putUnitDouble(c("Top "), c("#Pxl"), top);
    rect.putUnitDouble(c("Left"), c("#Pxl"), left);
    rect.putUnitDouble(c("Btom"), c("#Pxl"), bottom);
    rect.putUnitDouble(c("Rght"), c("#Pxl"), right);

    var desc = new ActionDescriptor();
    desc.putObject(c("T   "), c("Rctn"), rect);
    desc.putUnitDouble(c("Angl"), c("#Ang"), 0);
    desc.putBoolean(c("Dlt "), false); // Delete Cropped Pixels: off
    executeAction(c("Crop"), desc, DialogModes.NO);
}

// Returns the most recently added channel with the saved-selection name,
// or null if there isn't one.
function findCropChannel(doc) {
    var found = null;
    for (var i = 0; i < doc.channels.length; i++) {
        if (doc.channels[i].name === CROP_SELECTION_NAME) {
            found = doc.channels[i];
        }
    }
    return found;
}

main();
