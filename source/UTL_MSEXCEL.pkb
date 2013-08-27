CREATE OR REPLACE package body utl_msexcel as

	function create_workbook return number
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.NewWorkbook() return int';

	function open_workbook(p_file in blob) return number
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.OpenWorkbook(oracle.sql.BLOB) return int';

	function get_workbook(p_hnd in number) return blob
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.SaveWorkbook(int) return oracle.sql.BLOB';

	procedure free_workbook(p_hnd in number)
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.FreeWorkbook(int)';

	function sheet_count(p_hnd in number) return number
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.SheetCount(int) return int';

	function j$clone_sheet(p_hnd in number, p_sheet in number) return number
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.CloneSheet(int, int) return int';

	function j$add_sheet(p_hnd in number) return number
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.AddSheet(int) return int';

	procedure set_sheet_name(p_hnd in number, p_sheet in number, p_name in varchar2)
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.SetSheetName(int, int, java.lang.String)';

	procedure write_value(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_value in varchar2)
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.WriteString(int, int, int, int, java.lang.String)';
	procedure write_value(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_value in number)
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.WriteDouble(int, int, int, int, double)';
	procedure write_value(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_value in date)
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.WriteAnything(int, int, int, int, Date)';
	procedure write_dataset(p_hnd in number, p_sheet in number, p_row in number, p_column in number, p_dataset in clob)
	as language java name 'org.pgwc.oracle.msexcel.MSExcel.writeXmlDataset(int, int, int, int, oracle.sql.CLOB)';

	--
	-- CAST_DATASET
	-- Transform the cursor into an XML dataset for use with WRITE_DATASET
	--
	function cast_dataset(p_cursor in sys_refcursor, p_row_count out pls_integer) return clob is

		l_hnd     dbms_xmlgen.ctxhandle;
		l_ds      clob;

	begin

		l_hnd := dbms_xmlgen.newcontext(p_cursor);
		dbms_xmlgen.setConvertSpecialChars(l_hnd, true);
		dbms_xmlgen.setNullHandling(l_hnd, dbms_xmlgen.EMPTY_TAG);
		l_ds := dbms_xmlgen.getxml(l_hnd);
		if l_ds is null then

			l_ds := '<?xml version="1.0"?><ROWSET />';

		else

			--
			-- This hack changes the XML encoding from the default UTF-8 to the ISO-8859-5.
			-- This is to allow for a wider set of characters since the database allows for 
			-- more. 
			--
			l_ds := replace(l_ds, '<?xml version="1.0"?>', '<?xml version="1.0" encoding="ISO-8859-1"?>');

		end if;
		
		p_row_count := dbms_xmlgen.getNumRowsProcessed(l_hnd);
		
		dbms_xmlgen.closecontext(l_hnd);

		return l_ds;

	exception
		when OTHERS then
			dbms_xmlgen.closecontext(l_hnd);
			raise;

	end cast_dataset;

	function add_sheet(p_hnd in number, p_before in number default null, p_after in number default null) return number is

		l_sheet  number;

	begin

		if p_before is null and p_after is null then

			l_sheet := j$add_sheet(p_hnd);

		elsif p_after is null then

			l_sheet := -1;

		else

			l_sheet := -1;

		end if;

		return l_sheet;

	end add_sheet;

	function clone_sheet(p_hnd in number, p_sheet in number, p_after in number default null) return number is

		l_sheet  number;

	begin

		if p_after is null then

			l_sheet := j$clone_sheet(p_hnd, p_sheet);

		else

			raise_application_error(-20000, 'Method not implemented');

		end if;

		return l_sheet;

	end clone_sheet;

	procedure write_value
	(
		p_hnd       in number,
		p_sheet     in number,
		p_row       in number,
		p_column    in number,
		p_cursor    in sys_refcursor
	) is
	
		l_row_count pls_integer;
	
	begin
	
		write_value(p_hnd, p_sheet, p_row, p_column, p_cursor, l_row_count);
	
	end write_value;
	
	procedure write_value
	(
		p_hnd       in number,
		p_sheet     in number,
		p_row       in number,
		p_column    in number,
		p_cursor    in sys_refcursor,
		p_row_count out pls_integer
	) is

		l_ds        clob;
		
	begin

		l_ds := cast_dataset(p_cursor, p_row_count);
		write_dataset(p_hnd, p_sheet, p_row, p_column, l_ds);

	end write_value;

end utl_msexcel;
/