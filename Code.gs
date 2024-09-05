function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
      .addItem('Calculate AP Score', 'getRangeValues')
      .addToUi();
}

function getRangeValues() {
  var inputCol = Browser.inputBox('Enter column with scores: ', 'Input 1', Browser.Buttons.OK_CANCEL);
  var start = Browser.inputBox('Enter start range:', 'Input 2', Browser.Buttons.OK_CANCEL);
  var end = Browser.inputBox('Enter end value:', 'Input 3', Browser.Buttons.OK_CANCEL);
  var outputCol = Browser.inputBox('Enter column to paste scores: ', 'Input 4', Browser.Buttons.OK_CANCEL);
  
  // Check if user canceled
  if (inputCol === 'cancel' || start === 'cancel' || end === 'cancel' || outputCol === 'cancel') {
    Logger.log('User canceled the operation');
    return;
  }
  
  var result = calculateAPScores(inputCol, start, end, outputCol);
}

function calculateAPScores(inputCol, start, end, outputCol) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const range = sheet.getRange(inputCol+start+":"+inputCol+end);
  const values = range.getValues();
  
  const scores = values.map(row => row[0]).filter(score => typeof score === "number");
  const mean = scores.reduce((sum, score) => sum + score, 0) / scores.length;
  const stdDev = Math.sqrt(scores.map(score => Math.pow(score - mean, 2)).reduce((sum, sq) => sum + sq, 0) / (scores.length - 1));

  // Percentages for AP scores 1-5
  const distributionPercentages = [0.1, 0.20, 0.40, 0.20, 0.1];

  // Z-scores for the given percentages
  const zScores = [-1.6448536269514729, -1.0364333894937898, 0.2533471031357997, 0.8416212335729143, 1.2815515655446004]
  
  const scoreRanges = zScores.map(z => Math.floor(mean + z * stdDev));
  
  const apScores = scores.map(score => {
    if (score < scoreRanges[0]) return 1;
    if (score < scoreRanges[1]) return 2;
    if (score < scoreRanges[2]) return 3;
    if (score < scoreRanges[3]) return 4;
    return 5;
  });
  
  // Set the AP scores in the output column
  const apScoresRange = sheet.getRange(outputCol+start+":"+outputCol+end);
  apScoresRange.setValues(apScores.map(score => [score]));
  
  return apScores;
}

