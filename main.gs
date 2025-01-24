function filterDieSizes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var interfaceSheet = ss.getSheetByName("Interface");

  // Clear previous results
  interfaceSheet.getRange(8, 2, 10, 4).clearContent();
  interfaceSheet.getRange(8, 7, 10, 4).clearContent();
  interfaceSheet.getRange(18, 2, 2, 4).clearContent();
  interfaceSheet.getRange(18, 7, 2, 4).clearContent();

  // Get user input from named ranges
  var lengthInput = ss.getRangeByName("length").getValue();
  var widthInput = ss.getRangeByName("width").getValue();
  var depthInput = ss.getRangeByName("depth").getValue();
  var cartonShape = ss.getRangeByName("carton_shape").getValue();
  
  // Convert cartonShape to numerical ID
  var shapeMapping = { "Auto Lock": 10, "Straight Tuck End": 11 };
  var shapeID = shapeMapping[cartonShape];
  if (!shapeID) {
    Logger.log("Invalid carton shape selection: %s", cartonShape);
    return;
  }
  
  Logger.log("User Input - Length: %s, Width: %s, Depth: %s, Carton Shape ID: %s", lengthInput, widthInput, depthInput, shapeID);

  // Get full data range
  var dataRange = ss.getRangeByName("data_range").getValues();
  
  // Column indexes
  var ID_INDEX = 0;
  var LENGTH_INDEX = 1;
  var WIDTH_INDEX = 2;
  var DEPTH_INDEX = 3;
  var SHAPE_INDEX = 4;
  
  var maxResults = 10;
  var maxSmallerResults = 2;
  
  function filterResults(data, lengthCriteria, widthCriteria, depthCriteria) {
    return data.filter(row =>
      lengthCriteria(row[LENGTH_INDEX]) &&
      widthCriteria(row[WIDTH_INDEX]) &&
      depthCriteria(row[DEPTH_INDEX])
    );
  }
  
  function findNextSize(value, increments) {
    return Math.round((value + increments * 0.125) * 1000) / 1000;
  }
  
  function getFilteredResults(data, expandDepth, reverse = false, previousResults = []) {
    var results = [];

    // Create a Set of unique identifiers for previous results to enable fast lookups
    const previousIds = new Set(previousResults.map(row => row[ID_INDEX]));

    var filterSequence = [
      [0, 0],
      [0.125, 0], [0.125, 0.125],
      [0.25, 0.125], [0.25, 0.25],
      [0.375, 0.25], [0.375, 0.375]
    ];
    
    filterSequence.forEach(step => {
      let nextLength = findNextSize(lengthInput, reverse ? -step[0] / 0.125 : step[0] / 0.125);
      let nextWidth = findNextSize(widthInput, reverse ? -step[1] / 0.125 : step[1] / 0.125);
      
      results = results.concat(filterResults(data, 
        l => l === nextLength,
        w => w === nextWidth,
        d => reverse ? d === depthInput : d >= depthInput && d <= expandDepth));
    });

    // Filter out any results that appear in previousResults for reverse mode
    if (reverse) {
      results = results.filter(row => !previousIds.has(row[ID_INDEX]));
    }

    // Sort results based on the filter sequence order
    results.sort((a, b) => {
      // For smaller sizes (reverse), we want the closest smaller sizes first
      const aDistance = Math.abs(a[LENGTH_INDEX] - lengthInput) + Math.abs(a[WIDTH_INDEX] - widthInput);
      const bDistance = Math.abs(b[LENGTH_INDEX] - lengthInput) + Math.abs(b[WIDTH_INDEX] - widthInput);
      
      if (aDistance !== bDistance) {
        return aDistance - bDistance;
      }
      // If distances are equal, maintain consistent ordering
      if (a[LENGTH_INDEX] !== b[LENGTH_INDEX]) {
        return reverse ? b[LENGTH_INDEX] - a[LENGTH_INDEX] : a[LENGTH_INDEX] - b[LENGTH_INDEX];
      }
      if (a[WIDTH_INDEX] !== b[WIDTH_INDEX]) {
        return reverse ? b[WIDTH_INDEX] - a[WIDTH_INDEX] : a[WIDTH_INDEX] - b[WIDTH_INDEX];
      }
      return a[DEPTH_INDEX] - b[DEPTH_INDEX];
    });
    
    return results.slice(0, reverse ? maxSmallerResults : maxResults);
  }
  
  // Split data by shape type
  var straightTuckData = dataRange.filter(row => row[SHAPE_INDEX] === 11);
  var autoLockData = dataRange.filter(row => row[SHAPE_INDEX] === 10);
  
  // Get results based on selected shape
  var primaryData = shapeID === 11 ? straightTuckData : autoLockData;
  var alternateData = shapeID === 11 ? autoLockData : straightTuckData;
  
  var selectedResults = getFilteredResults(primaryData, depthInput + 2);
  var smallerSelectedResults = getFilteredResults(primaryData, depthInput, true, selectedResults);
  var alternateResults = getFilteredResults(alternateData, depthInput + 2);
  var smallerAlternateResults = getFilteredResults(alternateData, depthInput, true, alternateResults);
  
  if (selectedResults.length > 0) {
    interfaceSheet.getRange(8, 2, selectedResults.length, 4).setValues(
      selectedResults.map(row => [row[LENGTH_INDEX], row[WIDTH_INDEX], row[DEPTH_INDEX], row[ID_INDEX]])
    );
  }

  if (alternateResults.length > 0) {
    interfaceSheet.getRange(8, 7, alternateResults.length, 4).setValues(
      alternateResults.map(row => [row[LENGTH_INDEX], row[WIDTH_INDEX], row[DEPTH_INDEX], row[ID_INDEX]])
    );
  }

  if (smallerSelectedResults.length > 0) {
    interfaceSheet.getRange(18, 2, smallerSelectedResults.length, 4).setValues(
      smallerSelectedResults.map(row => [row[LENGTH_INDEX], row[WIDTH_INDEX], row[DEPTH_INDEX], row[ID_INDEX]])
    );
  }

  if (smallerAlternateResults.length > 0) {
    interfaceSheet.getRange(18, 7, smallerAlternateResults.length, 4).setValues(
      smallerAlternateResults.map(row => [row[LENGTH_INDEX], row[WIDTH_INDEX], row[DEPTH_INDEX], row[ID_INDEX]])
    );
  }

  interfaceSheet.getRange(6, 2).setValue(cartonShape);
  interfaceSheet.getRange(6, 7).setValue(cartonShape === "Auto Lock" ? "Straight Tuck End" : "Auto Lock");

  Logger.log("Results written to sheet.");
}
