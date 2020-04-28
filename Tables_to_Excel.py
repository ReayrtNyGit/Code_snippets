class Table_out:
    '''Allows mulitple dataframe outputs to Excel tables
    Requires classes to be wrapped between a loaded workbook template,
    and a saved workbook command from openpyxl. Classes must also
    use the the writeout() function to export the dataframe
    
    Example use below fills two tables from a workbook template starting at
    row 2, column 1 in Sheet1 and Sheet2 with dataframes df1 and df2

    wb = load_workbook(path_in_template)
    Table_out(wb,'Sheet1',2,1,df1).writeout()
    Table_out(wb,'Sheet2',2,1,df2).writeout()
    wb.save(path_out_filled_workbook)
    
    Attributes:

    wb an opened workbook from openpyxl
    sheet Excel sheet to write into
    start_row row to start writing into
    start_col column to start writing into
    df dataframe to write out
    '''

    def __init__(self,wb,sheet,start_row,start_col,df):
        ''' Return an object with in and output info'''
        self.wb=wb
        self.sheet=sheet
        self.df=df
        self.start_row=start_row
        self.start_col=start_col
        
    def writeout(self):
        '''Outputs the dataframe into a workbook'''
        ws = self.wb.get_sheet_by_name(self.sheet)
        rows = dataframe_to_rows(self.df,index=False, header=False)
        
        for r_idx, row in enumerate(rows, 1):  
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=(r_idx+(self.start_row-1)), column=(c_idx+(self.start_col-1)), value=value)
