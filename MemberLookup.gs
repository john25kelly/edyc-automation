/**
 * Member Name → Email Lookup
 *
 * Type a name (partial, case-insensitive) into cell P2 and press Tab or Enter.
 * The first matching name from column D will have its email written into Q2.
 */

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();

  // Only react to edits on cell P2
  if (range.getRow() !== 2 || range.getColumn() !== 16) return;

  var searchText = String(range.getValue()).trim().toLowerCase();
  var emailCell  = sheet.getRange('Q2');

  // Clear Q2 if P2 is empty
  if (!searchText) {
    emailCell.setValue('');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Read all name + email pairs from columns D and E (skipping header row 1)
  var data = sheet.getRange(2, 4, lastRow - 1, 2).getValues();
  // data[i][0] = column D (name), data[i][1] = column E (email)

  var exactMatch      = null;
  var startsWithMatch = null;
  var containsMatch   = null;

  for (var i = 0; i < data.length; i++) {
    var name  = String(data[i][0]).trim();
    var email = String(data[i][1]).trim();
    var lower = name.toLowerCase();

    if (!name) continue;

    if (lower === searchText) {
      exactMatch = email;
      break; // can't do better than exact
    }
    if (!startsWithMatch && lower.indexOf(searchText) === 0) {
      startsWithMatch = email;
    }
    if (!containsMatch && lower.indexOf(searchText) !== -1) {
      containsMatch = email;
    }
  }

  emailCell.setValue(exactMatch || startsWithMatch || containsMatch || '');
}
