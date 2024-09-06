function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
      .addItem('Calculate AP Score', 'calculateCurvedAPScores')
      .addItem('Calculate Normal Score', 'caclulateRegularAPScore')
      .addToUi();
}

function caclulateRegularAPScore() {
  // Prompts user to submit the score columns, and start/end ranges
  const inputCol = Browser.inputBox('Enter column with scores: ', 'Column Letter', Browser.Buttons.OK_CANCEL);
  const start = Browser.inputBox('Enter start range:', 'Start Row', Browser.Buttons.OK_CANCEL);
  const end = Browser.inputBox('Enter end value:', 'End Row', Browser.Buttons.OK_CANCEL);
  const outputCol = Browser.inputBox('Enter column to paste scores: ', 'Column Letter', Browser.Buttons.OK_CANCEL);
  
  // Check if user canceled
  if (inputCol === 'cancel' || start === 'cancel' || end === 'cancel' || outputCol === 'cancel') {
    Logger.log('User canceled the operation');
    return;
  }

  // Get the Google Sheet object, and select the values from the designated column
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(inputCol+start+":"+inputCol+end);
  const values = range.getValues();
  const scores = values.map(row => row[0]).filter(score => typeof score === "number");

  /**
   * Syllabus Grading Scale
   * 5 (A) --> 100% - 88%
   * 4 (B) --> 87% - 75%
   * 3 (C) --> 74% - 60%
   * 2 (D) --> 59% - 30%
   * 1 (F) --> 29% - 0%
   */
  const thresholds = [30, 60, 75, 88];
  const apScores = scores.map(score => {
    if (score < thresholds[0]) return 1;
    if (score < thresholds[1]) return 2;
    if (score < thresholds[2]) return 3;
    if (score < thresholds[3]) return 4;
    return 5;
  });

  // Set the AP scores in the output column
  const apScoresRange = sheet.getRange(outputCol+start+":"+outputCol+end);
  apScoresRange.setValues(apScores.map(score => [score]));
}

function calculateCurvedAPScores() {
  // Prompts user to submit the score columns, and start/end ranges
  const inputCol = Browser.inputBox('Enter column with scores: ', 'Column Letter', Browser.Buttons.OK_CANCEL);
  const start = Browser.inputBox('Enter start range:', 'Start Row', Browser.Buttons.OK_CANCEL);
  const end = Browser.inputBox('Enter end value:', 'End Row', Browser.Buttons.OK_CANCEL);
  const outputCol = Browser.inputBox('Enter column to paste scores: ', 'Column Letter', Browser.Buttons.OK_CANCEL);
  
  // Check if user canceled
  if (inputCol === 'cancel' || start === 'cancel' || end === 'cancel' || outputCol === 'cancel') {
    Logger.log('User canceled the operation');
    return;
  }
  
  // Get the Google Sheet object, and select the values from the designated column
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(inputCol+start+":"+inputCol+end);
  const values = range.getValues();
  
  // Verify the scores are numbers, then calculate the average and standard deviation
  const scores = values.map(row => row[0]).filter(score => typeof score === "number");
  const mean = scores.reduce((sum, score) => sum + score, 0) / scores.length;
  const stdDev = Math.sqrt(scores.map(score => Math.pow(score - mean, 2)).reduce((sum, sq) => sum + sq, 0) / (scores.length - 1));

  // Percentages for AP scores 1-5
  const distributionPercentages = [0.1, 0.20, 0.40, 0.20, 0.1];

  // Z-scores for the given percentages
  const zScores = [-1.6448536269514729, -1.0364333894937898, 0.2533471031357997, 0.8416212335729143, 1.2815515655446004]
  
  // Find threshold for 5, 4, 3, 2, or 1 scores
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
}
