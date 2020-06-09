//DESCRIPTION:Clear style overrides from Text, Tables or Objects
//SCRIPTMENU:Clear overrides
//This script is licensed under the GNU General Public License v3.0 https://www.gnu.org/licenses/gpl-3.0.txt
//Project: https://github.com/grefel/clearOverrides
//Contact: Gregor Fellenz - http://www.publishingx.de

start();

function start() {
    var px = {
        projectName: "Clear style overrides",
        version: "2019-12-22-v1.1",

        // Default GUI configuration
        processText: true,
        clearParagraphStyle: "[All]",
        clearCharacterStyle: "[All]",
        processTables: false,
        clearTableStyle: "[All]",
        processObjects: false,
        clearObjectStyle: "[All]",
    }

    if (app.documents.length == 0) {
        alert("Please run with an active document");
        return;
    }
    if (app.layoutWindows.length == 0) {
        alert("Please run with an active document");
        return;
    }

    var dok = app.documents[0];

    if (!getConfig(dok, px)) {
        // User cancelled
        return;
    }

    var ial = app.scriptPreferences.userInteractionLevel;
    var redraw = app.scriptPreferences.enableRedraw;
    var scriptPrefVersion = app.scriptPreferences.version;
    app.scriptPreferences.version = parseInt(app.version);
    try {
        processDok([dok, px]);
    } catch (error) {
        alert("Error: " + error + " Line: " + error.line);
    }
    finally {
        app.scriptPreferences.userInteractionLevel = ial;
        app.scriptPreferences.enableRedraw = redraw;
        app.scriptPreferences.version = scriptPrefVersion;
    }
}

/**
 * Process the active document with the configured values
 * @param {Array of Objects} args Document and processing information
 */
function processDok(args) {
    dok = args[0];
    px = args[1];

    if (px.processObjects) {
        var results = findOrChangeObject(dok, {}, null, true);
        for (var i = 0; i < results.length; i++) {
            var object = results[i];
            if (px.clearObjectStyle === "[All]" || object.appliedObjectStyle.id === px.clearObjectStyle.id) {
                object.clearObjectStyleOverrides()
            }
        }
    }
    if (px.processTables) {
        var findProperties = {
            findWhat: "<0016>"
        }
        var results = findOrChangeText(dok, findProperties, null, true);
        for (var i = 0; i < results.length; i++) {
            var table = results[i].tables[0];
            if (px.clearTableStyle === "[All]" || table.appliedTableStyle.id === px.clearTableStyle.id) {
                table.clearTableStyleOverrides();
                table.cells.everyItem().clearCellStyleOverrides();
            }
        }
    }
    if (px.processText) {
        var findProperties = {
            findWhat: "(?s).+"
        }
        if (px.clearParagraphStyle !== "[All]") {
            findProperties.appliedParagraphStyle = px.clearParagraphStyle;
        }
        if (px.clearCharacterStyle !== "[All]") {
            findProperties.appliedCharacterStyle = px.clearCharacterStyle;
        }
        var results = findOrChangeGrep(dok, findProperties, null, true);
        for (var i = 0; i < results.length; i++) {
            var text = results[i];
            text.clearOverrides();
        }
        // Need to check for single tables in a textframe oder cell 
        var findProperties = {
            findWhat: "<0016>"
        }
        if (px.clearParagraphStyle !== "[All]") {
            findProperties.appliedParagraphStyle = px.clearParagraphStyle;
        }
        if (px.clearCharacterStyle !== "[All]") {
            findProperties.appliedCharacterStyle = px.clearCharacterStyle;
        }
        var results = findOrChangeText(dok, findProperties, null, true);
        for (var i = 0; i < results.length; i++) {
            var text = results[i];
            text.clearOverrides();
        }
    }
}

/**
 * GUI for configuration
 * @param {Document} dok 
 * @param {Object} px 
 */
function getConfig(dok, px) {
    var mainWindow = new Window("dialog", px.projectName + " – " + px.version, undefined, { resizeable: true });

    textPanel = mainWindow.add("Panel", undefined, "Process Text");
    textPanel.alignment = ["fill", "fill"];
    textPanel.margins = [10, 20, 10, 10];
    textPanel.alignChildren = ["left", "top"];
    var processTextCheckBox = textPanel.add("Checkbox", undefined, "Clear text style overrides");
    processTextCheckBox.value = px.processText;

    textParagraphStyleGroup = textPanel.add("group");
    textParagraphStyleGroup.orientation = "row";
    textParagraphStyleGroup.enabled = processTextCheckBox.value
    var paragraphStyleText = textParagraphStyleGroup.add("statictext", undefined, "Process paragraph style");
    paragraphStyleText.preferredSize.width = "150";
    var paragraphStyleDropDown = textParagraphStyleGroup.add("dropdownlist");
    paragraphStyleDropDown.preferredSize.width = "300";
    paragraphStyleDropDown.alignment = ["fill", "fill"];
    var item = paragraphStyleDropDown.add('item', "[All]");
    item.paragraphStyle = "[All]";
    item.selected = true;
    for (var i = 0; i < dok.allParagraphStyles.length; i++) {
        item = paragraphStyleDropDown.add('item', getStyleString(dok.allParagraphStyles[i]));
        if (dok.allParagraphStyles[i].name == px.clearParagraphStyle) {
            item.selected = true;
        }
        item.paragraphStyle = dok.allParagraphStyles[i];
    }


    textCharacterStyleGroup = textPanel.add("group");
    textCharacterStyleGroup.orientation = "row";
    textCharacterStyleGroup.enabled = processTextCheckBox.value
    var characterStyleText = textCharacterStyleGroup.add("statictext", undefined, "Process character style");
    characterStyleText.preferredSize.width = "150";
    var characterStyleDropDown = textCharacterStyleGroup.add("dropdownlist");
    characterStyleDropDown.preferredSize.width = "300";
    characterStyleDropDown.alignment = ["fill", "fill"];
    var item = characterStyleDropDown.add('item', "[All]");
    item.characterStyle = "[All]";
    item.selected = true;
    for (var i = 0; i < dok.allCharacterStyles.length; i++) {
        item = characterStyleDropDown.add('item', getStyleString(dok.allCharacterStyles[i]));
        if (dok.allCharacterStyles[i].name == px.clearCharacterStyle) {
            item.selected = true;
        }
        item.characterStyle = dok.allCharacterStyles[i];
    }



    tablePanel = mainWindow.add("Panel", undefined, "Process Tables");
    tablePanel.alignment = ["fill", "fill"];
    tablePanel.margins = [10, 20, 10, 10];
    tablePanel.alignChildren = ["left", "top"];
    var processTableCheckBox = tablePanel.add("Checkbox", undefined, "Clear table style overrides");
    processTableCheckBox.value = px.processTable;

    tableStyleGroup = tablePanel.add("group");
    tableStyleGroup.orientation = "row";
    tableStyleGroup.enabled = processTableCheckBox.value
    var tableStyleTable = tableStyleGroup.add("statictext", undefined, "Process table style");
    tableStyleTable.preferredSize.width = "150";
    var tableStyleDropDown = tableStyleGroup.add("dropdownlist");
    tableStyleDropDown.preferredSize.width = "300";
    tableStyleDropDown.alignment = ["fill", "fill"];
    var item = tableStyleDropDown.add('item', "[All]");
    item.tableStyle = "[All]";
    item.selected = true;
    for (var i = 0; i < dok.allTableStyles.length; i++) {
        item = tableStyleDropDown.add('item', getStyleString(dok.allTableStyles[i]));
        if (dok.allTableStyles[i].name == px.clearTableStyle) {
            item.selected = true;
        }
        item.tableStyle = dok.allTableStyles[i];
    }

    objectPanel = mainWindow.add("Panel", undefined, "Process Objects");
    objectPanel.alignment = ["fill", "fill"];
    objectPanel.margins = [10, 20, 10, 10];
    objectPanel.alignChildren = ["left", "top"];
    var processObjectCheckBox = objectPanel.add("Checkbox", undefined, "Clear object style overrides");
    processObjectCheckBox.value = px.processObject;

    objectStyleGroup = objectPanel.add("group");
    objectStyleGroup.orientation = "row";
    objectStyleGroup.enabled = processObjectCheckBox.value
    var objectStyleObject = objectStyleGroup.add("statictext", undefined, "Process object style");
    objectStyleObject.preferredSize.width = "150";
    var objectStyleDropDown = objectStyleGroup.add("dropdownlist");
    objectStyleDropDown.preferredSize.width = "300";
    objectStyleDropDown.alignment = ["fill", "fill"];
    var item = objectStyleDropDown.add('item', "[All]");
    item.selected = true;
    item.objectStyle = "[All]";
    for (var i = 0; i < dok.allObjectStyles.length; i++) {
        item = objectStyleDropDown.add('item', getStyleString(dok.allObjectStyles[i]));
        if (dok.allObjectStyles[i].name == px.clearObjectStyle) {
            item.selected = true;
        }
        item.objectStyle = dok.allObjectStyles[i];
    }



    var controlGroup = mainWindow.add("group");
    controlGroup.orientation = "row";
    controlGroup.alignment = "fill";
    controlGroup.alignChildren = ["right", "center"];


    controlGroup.orientation = "row";
    controlGroup.alignment = "fill";
    with (controlGroup) {
        var goToWebsite = controlGroup.add("button", undefined, "by grefel");
        var divider = controlGroup.add("panel");
        divider.alignment = "fill";
        var cancelButton = controlGroup.add("button", undefined, "Cancel");
        var okButton = controlGroup.add("button", undefined, "Clear overrides", { name: "ok" });
    }

    processTextCheckBox.onClick = function () {
        textParagraphStyleGroup.enabled = processTextCheckBox.value;
        textCharacterStyleGroup.enabled = processTextCheckBox.value;
    }
    processTableCheckBox.onClick = function () {
        tableStyleGroup.enabled = processTableCheckBox.value;
    }
    processObjectCheckBox.onClick = function () {
        objectStyleGroup.enabled = processObjectCheckBox.value;
    }


    goToWebsite.onClick = function () {
        openURL('https://www.publishingX.de');
    }
    cancelButton.onClick = function () {
        mainWindow.close(2);
    }
    okButton.onClick = function () {
        px.processText = processTextCheckBox.value;
        px.clearParagraphStyle = paragraphStyleDropDown.selection.paragraphStyle;
        px.clearCharacterStyle = characterStyleDropDown.selection.characterStyle;
        px.processTables = processTableCheckBox.value;
        px.clearTableStyle = tableStyleDropDown.selection.tableStyle;
        px.processObjects = processObjectCheckBox.value;
        px.clearObjectStyle = objectStyleDropDown.selection.objectStyle;
        mainWindow.close(1);
    }

    mainWindow.center()
    var res = mainWindow.show();

    if (res === 1) {
        return true;
    }
    else {
        return false;
    }
}

/**
 * Find or change with GREP
 * @param {Object} where An InDesign Object to search within (Document, Story, Text, Table, TextFrame) 
 * @param {Object|String} find String or findGrepPreferences.properties to search for
 * @param {Object|String|null} change String or changeGrepPreferences.properties to search for. If null, will only search in object *where* Note: Resulting Array is reversed
 * @param {Boolean} includeMaster Defaults to false
 */
function findOrChangeGrep(where, find, change, includeMaster) {
    if (change == undefined) {
        change = null;
    }
    if (includeMaster == undefined) {
        includeMaster = false;
    }

    // Save Options
    var saveFindGrepOptions = {};
    saveFindGrepOptions.includeFootnotes = app.findChangeGrepOptions.includeFootnotes;
    saveFindGrepOptions.includeHiddenLayers = app.findChangeGrepOptions.includeHiddenLayers;
    saveFindGrepOptions.includeLockedLayersForFind = app.findChangeGrepOptions.includeLockedLayersForFind;
    saveFindGrepOptions.includeLockedStoriesForFind = app.findChangeGrepOptions.includeLockedStoriesForFind;
    saveFindGrepOptions.includeMasterPages = app.findChangeGrepOptions.includeMasterPages;
    if (app.findChangeGrepOptions.hasOwnProperty("searchBackwards")) saveFindGrepOptions.searchBackwards = app.findChangeGrepOptions.searchBackwards;

    // Set Options
    app.findChangeGrepOptions.includeFootnotes = true;
    app.findChangeGrepOptions.includeHiddenLayers = true;
    app.findChangeGrepOptions.includeLockedLayersForFind = false;
    app.findChangeGrepOptions.includeLockedStoriesForFind = false;
    app.findChangeGrepOptions.includeMasterPages = includeMaster;
    if (app.findChangeGrepOptions.hasOwnProperty("searchBackwards")) app.findChangeGrepOptions.searchBackwards = false;

    // Reset Dialog
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    try {
        // Find Change operation
        if (find.constructor.name == "String") {
            app.findGrepPreferences.findWhat = find;
        }
        else {
            app.findGrepPreferences.properties = find;
        }
        if (change != null && change.constructor.name == "String") {
            app.changeGrepPreferences.changeTo = change;
        }
        else if (change != null) {
            app.changeGrepPreferences.properties = change;
        }
        var results = null;
        if (change == null) {
            results = where.findGrep(true);
        }
        else {
            results = where.changeGrep();
        }
    }
    catch (e) {
        throw e;
    }
    finally {
        // Reset Dialog
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;

        // Reset Options
        app.findChangeGrepOptions.includeFootnotes = saveFindGrepOptions.includeFootnotes;
        app.findChangeGrepOptions.includeHiddenLayers = saveFindGrepOptions.includeHiddenLayers;
        app.findChangeGrepOptions.includeLockedLayersForFind = saveFindGrepOptions.includeLockedLayersForFind;
        app.findChangeGrepOptions.includeLockedStoriesForFind = saveFindGrepOptions.includeLockedStoriesForFind;
        app.findChangeGrepOptions.includeMasterPages = saveFindGrepOptions.includeMasterPages;
        if (app.findChangeGrepOptions.hasOwnProperty("searchBackwards")) app.findChangeGrepOptions.searchBackwards = saveFindGrepOptions.searchBackwards;
    }

    return results;
}

/**
 * Find or change with TEXT
 * @param {Object} where An InDesign Object to search within (Document, Story, Text, Table, TextFrame) 
 * @param {Object|String} find String or findTextPreferences.properties to search for
 * @param {Object|String|null} change String or changeTextPreferences.properties to search for. If null, will only search in object *where* Note: Resulting Array is reversed
 * @param {Boolean} includeMaster Defaults to false
 */
function findOrChangeText(where, find, change, includeMaster) {
    if (change == undefined) {
        change = null;
    }
    if (includeMaster == undefined) {
        includeMaster = false;
    }

    // Save Options
    var saveFindTextOptions = {};
    saveFindTextOptions.includeFootnotes = app.findChangeTextOptions.includeFootnotes;
    saveFindTextOptions.includeHiddenLayers = app.findChangeTextOptions.includeHiddenLayers;
    saveFindTextOptions.includeLockedLayersForFind = app.findChangeTextOptions.includeLockedLayersForFind;
    saveFindTextOptions.includeLockedStoriesForFind = app.findChangeTextOptions.includeLockedStoriesForFind;
    saveFindTextOptions.includeMasterPages = app.findChangeTextOptions.includeMasterPages;
    if (app.findChangeTextOptions.hasOwnProperty("searchBackwards")) saveFindTextOptions.searchBackwards = app.findChangeTextOptions.searchBackwards;

    // Set Options
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = true;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = includeMaster;
    if (app.findChangeTextOptions.hasOwnProperty("searchBackwards")) app.findChangeTextOptions.searchBackwards = false;

    // Reset Dialog
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    try {
        // Find Change operation
        if (find.constructor.name == "String") {
            app.findTextPreferences.findWhat = find;
        }
        else {
            app.findTextPreferences.properties = find;
        }
        if (change != null && change.constructor.name == "String") {
            app.changeTextPreferences.changeTo = change;
        }
        else if (change != null) {
            app.changeTextPreferences.properties = change;
        }
        var results = null;
        if (change == null) {
            results = where.findText(true);
        }
        else {
            results = where.changeText();
        }
    }
    catch (e) {
        throw e;
    }
    finally {
        // Reset Dialog
        app.findTextPreferences = NothingEnum.nothing;
        app.changeTextPreferences = NothingEnum.nothing;

        // Reset Options
        app.findChangeTextOptions.includeFootnotes = saveFindTextOptions.includeFootnotes;
        app.findChangeTextOptions.includeHiddenLayers = saveFindTextOptions.includeHiddenLayers;
        app.findChangeTextOptions.includeLockedLayersForFind = saveFindTextOptions.includeLockedLayersForFind;
        app.findChangeTextOptions.includeLockedStoriesForFind = saveFindTextOptions.includeLockedStoriesForFind;
        app.findChangeTextOptions.includeMasterPages = saveFindTextOptions.includeMasterPages;
        if (app.findChangeTextOptions.hasOwnProperty("searchBackwards")) app.findChangeTextOptions.searchBackwards = saveFindTextOptions.searchBackwards;
    }

    return results;
}

/**
 * Find or change with OBJECT
 * @param {Object} where An InDesign Object to search within (Document, Story, Object, Table, ObjectFrame) 
 * @param {Object} find findObjectPreferences.properties to search for
 * @param {Object|null} change changeObjectPreferences.properties to search for. If null, will only search in object *where* Note: Resulting Array is reversed
 * @param {Boolean} includeMaster Defaults to false
 */
function findOrChangeObject(where, find, change, includeMaster) {
    if (change == undefined) {
        change = null;
    }
    if (includeMaster == undefined) {
        includeMaster = false;
    }

    // Save Options
    var saveFindObjectOptions = {};
    saveFindObjectOptions.includeFootnotes = app.findChangeObjectOptions.includeFootnotes;
    saveFindObjectOptions.includeHiddenLayers = app.findChangeObjectOptions.includeHiddenLayers;
    saveFindObjectOptions.includeLockedLayersForFind = app.findChangeObjectOptions.includeLockedLayersForFind;
    saveFindObjectOptions.includeLockedStoriesForFind = app.findChangeObjectOptions.includeLockedStoriesForFind;
    saveFindObjectOptions.includeMasterPages = app.findChangeObjectOptions.includeMasterPages;
    if (app.findChangeObjectOptions.hasOwnProperty("searchBackwards")) saveFindObjectOptions.searchBackwards = app.findChangeObjectOptions.searchBackwards;

    // Set Options
    app.findChangeObjectOptions.includeFootnotes = true;
    app.findChangeObjectOptions.includeHiddenLayers = true;
    app.findChangeObjectOptions.includeLockedLayersForFind = false;
    app.findChangeObjectOptions.includeLockedStoriesForFind = false;
    app.findChangeObjectOptions.includeMasterPages = includeMaster;
    if (app.findChangeObjectOptions.hasOwnProperty("searchBackwards")) app.findChangeObjectOptions.searchBackwards = false;
    app.findChangeObjectOptions.objectType = ObjectTypes.ALL_FRAMES_TYPE;
    // 	ObjectTypes.GRAPHIC_FRAMES_TYPE
    // 	ObjectTypes.TEXT_FRAMES_TYPE
    // 	ObjectTypes.UNASSIGNED_FRAMES_TYPE;

    // Reset Dialog
    app.findObjectPreferences = NothingEnum.nothing;
    app.changeObjectPreferences = NothingEnum.nothing;

    try {
        // Find Change operation
        app.findObjectPreferences.properties = find;
        if (change != null) {
            app.changeObjectPreferences.properties = change;
        }
        var results = null;
        if (change == null) {
            results = where.findObject(true);
        }
        else {
            results = where.changeObject();
        }
    }
    catch (e) {
        throw e;
    }
    finally {
        // Reset Dialog
        app.findObjectPreferences = NothingEnum.nothing;
        app.changeObjectPreferences = NothingEnum.nothing;

        // Reset Options
        app.findChangeObjectOptions.includeFootnotes = saveFindObjectOptions.includeFootnotes;
        app.findChangeObjectOptions.includeHiddenLayers = saveFindObjectOptions.includeHiddenLayers;
        app.findChangeObjectOptions.includeLockedLayersForFind = saveFindObjectOptions.includeLockedLayersForFind;
        app.findChangeObjectOptions.includeLockedStoriesForFind = saveFindObjectOptions.includeLockedStoriesForFind;
        app.findChangeObjectOptions.includeMasterPages = saveFindObjectOptions.includeMasterPages;
        if (app.findChangeObjectOptions.hasOwnProperty("searchBackwards")) app.findChangeObjectOptions.searchBackwards = saveFindObjectOptions.searchBackwards;
    }

    return results;
}

/**
 * Opens an URL in the default system Browser
 * Credits: Trevor http://creative-scripts.com  
 * @param {String} url 
 */
function openURL(url) {
    if ($.os[0] === 'M') { // Mac  
        app.doScript('open location "' + url + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
    } else { // Windows  
        app.doScript('CreateObject("WScript.Shell").Run("' + url + '")', ScriptLanguage.VISUAL_BASIC);
    }
}

/**
 * Returns the style name including its parent stylegroups, separated by : 
 * @param {CharacterStyle|ParagraphStyle|TableStyle|CellStyle|ObjectStyle} style The style to process
 */
function getStyleString(style) {
    var styleName = style.name;
    while (style.parent.constructor.name.match(/Group$/)) {
        style = style.parent;
        styleName = style.name + ":" + styleName;
    }
    return styleName;
}