function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Search Similar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('searchBar')
      .setTitle('Search Similar')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function findSimilar(searchQuery, range, threshold) {
  var output = [];
  for (var i = 0; i < range.length; i++) {
    var distance = levenshteinDistance(searchQuery.toString().toLowerCase(), range[i][0].toString().toLowerCase());
    if (distance <= threshold) {
      output.push([range[i][0]]);
    }
  }
  return output;
}

function levenshteinDistance(a, b) {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  
  var matrix = [];
  
  for (var i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  
  for (var j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }
  
  for (var i = 1; i <= b.length; i++) {
    for (var j = 1; j <= a.length; j++) {
      if (b.charAt(i-1) === a.charAt(j-1)) {
        matrix[i][j] = matrix[i-1][j-1];
      } else {
        matrix[i][j] = Math.min(matrix[i-1][j-1] + 1,
                                Math.min(matrix[i][j-1] + 1,
                                         matrix[i-1][j] + 1));
      }
    }
  }
  
  return matrix[b.length][a.length];
}
