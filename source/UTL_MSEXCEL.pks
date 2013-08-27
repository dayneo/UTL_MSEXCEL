CREATE OR REPLACE package utl_msexcel as

	--
	-- Workbook methods
	--
	function create_workbook return number;
	function open_workbook(p_file in blob) return number;
	function get_workbook(p_hnd in number) return blob;
	procedure free_workbook(p_hnd in number);

	--
	-- Worksheet methods
	--
	function sheet_count(p_hnd in number) return number;
	function add_sheet(p_hnd in number, p_before in number default null, p_after in number default null) return number;
	-- procedure remove_sheet(p_hnd in number, p_sheet in number);
	function clone_sheet(p_hnd in number, p_sheet in number, p_after in number default null) return number;
	procedure set_sheet_name(p_hnd in number, p_sheet in number, p_name in varchar2);
	
	--
	-- WRITE_VALUE
	-- Write a value to a specified cell
	--
	procedure write_value(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_value in varchar2);
	procedure write_value(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_value in number);
	procedure write_value(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_value in date);
	
	-- Note dates must be formatted as 'yyyy-mm-dd"T"hh24:mi:ss'
	-- This can be achieved by either using the NLS parameters or via explicit to_char on the 
	-- date column.
	procedure write_value(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_cursor in sys_refcursor);
	procedure write_value(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_cursor in sys_refcursor, p_row_count out pls_integer);

end utl_msexcel;
/

SHOW ERRORS