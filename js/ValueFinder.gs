/*=====================================================================================================================================
 *
 * ValueFinder finds ranges of VALUES, FORMULAS & DATAVALIDATIONS in cells
 * 
 * @method object	find(string|regEx _search, string _search_for, Sheet|[Sheet] )	find _search value in _search_for type
 * 
 * @method [A1]	getA1Notation()	get found ranges as array of A1 notations	E.G: of returned array ['A1:A5','C1:D5']
 * @method [Range]	getRanges()	get found ranges as array of Range objects	https://developers.google.com/apps-script/reference/spreadsheet/range
 *
 * @method void	testRanges([A1|Range] _search_for)	colorize found ranges use current found ranges if _search_for is undefined
 *
 * @return {SheetName:[A1Notations|Ranges]}
 * 
 * ------------------------------------------------------------------------------------------------------------------------
 * 
 * @example of _search_for static variable
 * 
 *		(new ValueFinder()).find(/foo/,	['VALUE'])	search VALUE,   in cells values,	VALUES is default, not necessary to be defined
 *		(new ValueFinder()).find(/=SUM/gi,	['FORMULA'])	search FORMULA, in cells formulas,	not necessary to be defined BUT REGEX MUST START WITH EQUAL SIGN '=' E.G: '=FORMULA'
 *     
 *		(new ValueFinder()).find( null, 'RULE|RULE_LIST|RULE_RANGE')	search DATA VALIDATIONS of 'List of items|List from range', both will be searched if 'RULE'
 *
 *		(new ValueFinder()).find(/item(A|B)/,	'RULE_LIST')	search ITEMS of 'List of items'
 *		(new ValueFinder()).find('A1:A5',	'RULE_RANGE')	search RANGE of 'List from range'
 *
 *		(new ValueFinder()).find(/item(A|B)/,	'RULE_VALUE')	search selected VALUE of 'List of items' AND 'List from range'
 *		(new ValueFinder()).find(/item(A|B)/,	'RULE_LIST_VALUE')	search selected VALUE of 'List of items'	
 *		(new ValueFinder()).find(/item(A|B)/,	'RULE_RANGE_VALUE')	search selected VALUE of 'List from range'
 *		
 *		(new ValueFinder()).find( /Validation\s\d+/, 'RULE_HELP')	search in all data validations for 'Help text'
 *
 *		(new ValueFinder(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet())).find().getA1Notation() // get A1notations of all filled cells in active sheet
 *		(new ValueFinder()).find().getRanges() 	// get Ranges of all filled cells
 *		(new ValueFinder()).find().testRanges()	// colrize foud ranges
 *
=====================================================================================================================================*/

var ValueFinder = function()
{ 
	var spread	= SpreadsheetApp.getActiveSpreadsheet();
	var _SheetValueFinders	= {};
	/** Get new instance of _Template class
	 * 'Shadow' for function _new();
	 * Allows easily call this class from library
	 *
	 * Execution _Template._new() works for STANDALONE and also for LIBRARY
	 */
	this._new = function()
	{
		return this;
	};
	/** find
	 * @param string|regEx	_search 
	 * @param string	_search_for 
	 * @param Sheet|[Sheet]	Sheet object or array of Sheets where to search, use all Sheets if undefined
	 * 
	 * @return this ValueFinder
	 */
	this.find = function( _search, _search_for, sheets )
	{
		if(typeof sheets !== 'undefined' && !Array.isArray(sheets)){
			sheets = [sheets];
		}else if( typeof sheets === 'undefined' )
			sheets = spread.getSheets();
			
		for(var s=0; s<sheets.length;s++) {
			var sheet = sheets[s];
			_SheetValueFinders [sheet.getSheetName()] = (new SheetValueFinder()).find( _search, _search_for, sheet );
		}
		Logger.log( "_SheetValueFinders "+JSON.stringify( _SheetValueFinders)  );		
		return this;

	};		
	/** getA1Notation of found ranges
	 *
	 * @return {SheetName:[A1Notation]} _SheetValueFinders
	 */
	this.getA1Notation = function()
	{
		//sheet_name_prefixed = typeof sheet_name_prefixed !== 'undefined' && sheet_name_prefixed===true;
		
		for(var sheet_name in _SheetValueFinders){if (_SheetValueFinders.hasOwnProperty(sheet_name)){
			//sheet_name_prefix = sheet_name_prefixed ? sheet_name : false;
			_SheetValueFinders[sheet_name] = _SheetValueFinders[sheet_name].getA1Notation(sheet_name);
		}}
		Logger.log( "_SheetValueFinders "+JSON.stringify( _SheetValueFinders)  );		
		return _SheetValueFinders;
	};
	/** convert number notation to Range objects
	 *
	 * @return {SheetName:[Range]} _SheetValueFinders
	 */
	this.getRanges = function()
	{
		return call('getRanges');
	};
	/** colorize found ranges
	 * 
	 * @return this ValueFinder
	 */
	this.testRanges = function() {
		call('testRanges');
		return this;
	};
	
	
	/** call one of public function in SheetValueFinder
	 *
	 * @return {SheetName:[]} _SheetValueFinders
	 */
	var call = function(fn_name)
	{
		for(var sheet_name in _SheetValueFinders){if (_SheetValueFinders.hasOwnProperty(sheet_name)){
			_SheetValueFinders[sheet_name] = _SheetValueFinders[sheet_name][fn_name](sheet_name);
		}}
		Logger.log( "_SheetValueFinders "+JSON.stringify( _SheetValueFinders)  );		
		return _SheetValueFinders;
	};
	
	return this;
//	return typeof ValueFinder !=='object' ? project : this;
}; // end of ValueFinder 

