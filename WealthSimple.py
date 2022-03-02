import re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

def sheet_by_name( name ):
	desktop = XSCRIPTCONTEXT.getDesktop( )
	model = desktop.getCurrentComponent( )
	return model.Sheets.getByName( name )

def configuration( ):
	KEY_COLUMN = 0
	VALUE_COLUMN = 1
	
	sheet = sheet_by_name( "Configuration" )
	cursor = sheet.createCursor( )
	cursor.gotoEndOfUsedArea( False )
	rows = cursor.RangeAddress.EndRow + 1
	
	config = { }
	for row in range( 1, rows ):
		key = sheet.getCellByPosition( KEY_COLUMN, row ).getString( )
		value = sheet.getCellByPosition( VALUE_COLUMN, row ).getString( )
		config[key] = value
	return config

def open_driver_session( config ):
	option = webdriver.ChromeOptions( )
	option.binary_location = config["browser"]
	option.add_argument( "--incognito" )
	# option.add_argument( "--headless" )

	webpage = webdriver.Chrome( executable_path = config["driver"], chrome_options = option )
	webpage.get( config["url"] )

	return webpage

def get_webpage_value( webpage, key ):
	text = webpage.find_element_by_id( key ).text
	return float( re.sub( "[^\d\.]", "", text ) )

def set_webpage_value( webpage, key, value ):
	WebDriverWait( webpage, 10 ).until( EC.element_to_be_clickable( ( By.ID, key ) ) ).send_keys( Keys.CONTROL + "a" )
	WebDriverWait( webpage, 10 ).until( EC.element_to_be_clickable( ( By.ID, key ) ) ).send_keys( Keys.DELETE )
	WebDriverWait( webpage, 10 ).until( EC.element_to_be_clickable( ( By.ID, key ) ) ).send_keys( value )

def get_named_range( sheet, name ):
	try:
		return sheet.getCellRangeByName( name )
	except:
		return None

def get_spreadsheet_value( range, name, row ):
	return str( range[name].getCellByPosition( 0, row ).getValue( ) if range[name] else 0.0 )

def set_spreadsheet_value( range, name, row, value ):
	if range[name]:
		cell = range[name].getCellByPosition( 0, row )
		cell.setValue( value )

def CalculateTaxes( *args ):
	config = configuration( )
	webpage = open_driver_session( config )

	sheet = sheet_by_name( "Model" )
	input_range = {
		"employmentIncome": get_named_range( sheet, "IN_employmentIncome" ),
		"selfEmploymentIncome": get_named_range( sheet, "IN_selfEmploymentIncome" ),
		"rrspDeduction": get_named_range( sheet, "IN_rrspDeduction" ),
		"capitalGains": get_named_range( sheet, "IN_capitalGains" ),
		"eligibleDividends": get_named_range( sheet, "IN_eligibleDividends" ),
		"ineligibleDividends": get_named_range( sheet, "IN_ineligibleDividends" ),
		"otherIncome": get_named_range( sheet, "IN_otherIncome" ),
		"incomeTaxesPaid": get_named_range( sheet, "IN_incomeTaxesPaid" ), }
	output_range = {
		"totalIncome": get_named_range( sheet, "OUT_totalIncome" ),
		"totalTax": get_named_range( sheet, "OUT_totalTax" ),
		"afterTaxIncome": get_named_range( sheet, "OUT_afterTaxIncome" ),
		"averageTaxRate": get_named_range( sheet, "OUT_averageTaxRate" ),
		"marginalTaxRate": get_named_range( sheet, "OUT_marginalTaxRate" ), }

	for row in range( input_range["otherIncome"].Rows.Count ):
		for column_name in input_range:
			income_data = get_spreadsheet_value( input_range, column_name, row )
			set_webpage_value( webpage, column_name, income_data )
		for column_name in output_range:
			tax_data = get_webpage_value( webpage, column_name )
			set_spreadsheet_value( output_range, column_name, row, tax_data )

	webpage.close( )

	return None

g_exportedScripts = CalculateTaxes,
