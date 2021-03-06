// Make requests traceable in case of any issues
var HEADERS = {headers: {
  'X-User-Agent': 'Google Sheets'
}};

var ENDPOINT = 'https://api.bionomia.net/users/search';

/**
 * Searches for a name in Bionomia and returns a Bionomia URL.
 *
 * @param {string} search The search name (required).
 * @param {string} options Key:value separated by commas such as "family_collected:Asilidae, family_identified:Pisauridae" (optional).
 * @return {string} The Bionomia URL.
 * @customfunction
 */
function BIONOMIA(search, options) {
  'use strict';
  return fetchBionomia_(search, options, '@id');
}

/**
 * Searches for a name in Bionoimia and returns a Wikidata or an ORCID URI.
 *
 * @param {string} search The search name (required).
 * @param {string} options Key:value separated by commas such as "family_collected:Asilidae, family_identified:Pisauridae" (optional).
 * @return {string} The wikidata or ORCID entity URI.
 * @customfunction
 */
function BIONOMIAURI(search, options) {
  'use strict';
  return fetchBionomia_(search, options, 'sameAs');
}

/**
 * Searches for a name in Bionomia and returns a formatted name.
 *
 * @param {string} search The search name (required).
 * @param {string} options Key:value separated by commas such as "family_collected:Asilidae, family_identified:Pisauridae" (optional).
 * @return {string} The formatted name.
 * @customfunction
 */
function BIONOMIANAME(search, options) {
  'use strict';
  return fetchBionomia_(search, options, 'name');
}

/**
 * Executed on add-on install.
 */
function onInstall() {
  'use strict';
  onOpen();
}

/**
 * Executed on add-on open.
 */
function onOpen() {
  'use strict';
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Show documentation', 'showDocumentation_')
      .addToUi();
}

/**
 * Shows a sidebar with help.
 */
function showDocumentation_() {
  'use strict';
  var html = HtmlService.createHtmlOutputFromFile('Documentation')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Documentation')
      .setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

function fetchBionomia_(search, options, response_key) {
  'use strict';
  if (!search) {
    return '';
  }

  var result = '', collected = '', identified = '', date = '', strict = '';
  var opts = typeof options !== 'undefined' ? options.split(",") : [];
  var json_key = typeof response_key !== 'undefined' ? response_key : "@id";

  opts.forEach(function(item){
    var key_values = item.split(":");
    switch(key_values[0].trim()) {
      case "family_collected":
        collected = key_values[1].trim();
        break;
      case "family_identified":
        identified = key_values[1].trim();
        break;
      case "date":
        date = key_values[1].trim();
        break;
      case "strict":
        strict = true;
        break;
      default:
    }
  });

  try {
    var url = ENDPOINT +
        '?q=' + encodeURIComponent(search) +
        '&families_collected=' + encodeURIComponent(collected) +
        '&families_identified=' + encodeURIComponent(identified) +
        '&date=' + encodeURIComponent(date) +
        '&strict=' + strict;
    var json = JSON.parse(UrlFetchApp.fetch(url, HEADERS).getContentText());
    // TODO: deal with multiple hits, cutoff for scores?
    result = json.dataFeedElement[0].item[json_key];
  } catch (e) {
    //no op
  }
  return result.length > 0 ? result : '';
}
