var path = require( 'path' )
  , XLSX = require( 'xlsx' )
  ;

function extractText( filePath, options, cb ) {
  var CSVs, wb, result, error;

  result = '';
  try {
    wb = XLSX.readFile( filePath );
    Object.entries(wb.Props).forEach( function( entry ) {
      if (entry[0] !== 'Language') {
        result += `${entry[1]},`;
      }
    });
    Object.entries(wb.Sheets).forEach( function( entry ) {
      result += `${entry[0]},+${XLSX.utils.sheet_to_csv(entry[1])},`;
    });
  } catch ( err ) {
    error = new Error( 'Could not extract ' + path.basename( filePath ) + ', ' + err );
    cb( error, null );
    return;
  }

  cb( null, result );
}

module.exports = {
  types: ['application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
    'application/vnd.ms-excel.sheet.macroEnabled.12',
    'application/vnd.oasis.opendocument.spreadsheet',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
    'application/vnd.oasis.opendocument.spreadsheet-template'
  ],
  extract: extractText
};
