using System.Collections.Generic;
using System.Data;
using Aspose.Cells;
using Aspose.Pdf;

class AsposHelper {
    public void convertPDFExcel(string PDFPath, string outputPath){
        var lic = new Aspose.Pdf.License();
        lic.SetLicense("Aspose.Total.lic");
        Aspose.Pdf.Document doc = new Aspose.Pdf.Document(PDFPath);
        Aspose.Pdf.ExcelSaveOptions excelsave = new ExcelSaveOptions();
        doc.Save(outputPath, excelsave);
    }
    public int getWorksheetsCount(string workBookpath ) {
        var lic = new Aspose.Cells.License();
        lic.SetLicense("Aspose.Total.lic");
        int worksheetCount;
        Workbook workbook =new Workbook(workBookpath);
        WorksheetCollection worksheets = workbook.Worksheets;
        worksheetCount = worksheets.Count ;
        return worksheetCount;
   }
    public bool searchStringWorksheet(string workBookpath,int sheetIndex , string searchString ) {
        if(!string.IsNullOrEmpty(searchString))
        {
            var lic = new Aspose.Cells.License();
            lic.SetLicense("Aspose.Total.lic");
            Workbook workbook =new Workbook(workBookpath);
            Worksheet sheet = workbook.Worksheets[sheetIndex];
            Aspose.Cells.Cells cells = sheet.Cells;
            FindOptions findOptions = new FindOptions();
            findOptions.CaseSensitive = false;
            findOptions.LookInType = LookInType.Values;
            Aspose.Cells.Cell foundCell = cells.Find(searchString, null, findOptions);
            if (foundCell != null)
            {
                return true;    
            }else{
                return false;
            }
        }else{
            return false;
        }
   }
    public int getRowFromString(string workBookpath,int sheetIndex , string searchString ) {
        var lic = new Aspose.Cells.License();
        lic.SetLicense("Aspose.Total.lic");
        int rowIndex;
        Workbook workbook =new Workbook(workBookpath);
        Worksheet sheet = workbook.Worksheets[sheetIndex];
        Aspose.Cells.Cells cells = sheet.Cells;
        FindOptions findOptions = new FindOptions();
        findOptions.CaseSensitive = false;
        findOptions.LookInType = LookInType.Values;
        Aspose.Cells.Cell foundCell = cells.Find(searchString, null, findOptions);
        if (foundCell != null)
        {
            rowIndex = foundCell.Row;   
        }else{
            rowIndex = 0;
        }
        return rowIndex;
   }
    public System.Tuple<int, int> getRowsColumns(string workBookpath,int sheetIndex){
        var lic = new Aspose.Cells.License();
        lic.SetLicense("Aspose.Total.lic");
        Workbook workbook =new Workbook(workBookpath);
        Worksheet sheet = workbook.Worksheets[sheetIndex];
        Aspose.Cells.Cell cell =sheet.Cells.LastCell;    
        return System.Tuple.Create(cell.Row + 1,cell.Column + 1);
    }
    public string getCellValue(string workBookpath,int sheetIndex,int rowIndex,int columnIndex){
        var lic = new Aspose.Cells.License();
        lic.SetLicense("Aspose.Total.lic");
        string cellvalue ;
        Workbook workbook =new Workbook(workBookpath);
        Worksheet sheet = workbook.Worksheets[sheetIndex];
        Aspose.Cells.Cell cell = sheet.Cells[row: rowIndex, column: columnIndex];
        if(cell!=null)
        {
            if(cell.Value!=null)
            {
            cellvalue = cell.Value.ToString();
            }else{
                cellvalue = null;
            }
        }else{
            cellvalue = null;
        }
        return cellvalue;
    }
    public DataTable loadExcelDataTable(string workBookpath , int sheetIndex){
        var lic = new Aspose.Cells.License();
        lic.SetLicense("Aspose.Total.lic");
        DataTable dt = new DataTable();
        Workbook wb = new Workbook(workBookpath);
        dt = wb.Worksheets[sheetIndex].Cells.ExportDataTable(0, 0, wb.Worksheets[0].Cells.MaxDataRow + 1, wb.Worksheets[sheetIndex].Cells.MaxDataColumn + 1);
        return dt;
    }
    public bool isColumnEmpty(string workBookPath , int sheetIndex,int startRow , int maxRow, int columnIndex)
    {
        List<string> columnList = new List<string>();
        for(int i=startRow ; i <maxRow ;i++){
            var columnValue = getCellValue(workBookPath, sheetIndex,i,columnIndex);    
            if(!string.IsNullOrEmpty(columnValue))
                {
                    columnList.Add(columnValue);
                }
         }
        if(columnList.Count > 0)
        {
            return false;
        }else{
            return true;
        }
    }
    public int getColumnIndexString(string workBookpath,int sheetIndex , string searchString ) {
        var lic = new Aspose.Cells.License();
        lic.SetLicense("Aspose.Total.lic");
        int columnIndex;
        Workbook workbook =new Workbook(workBookpath);
        Worksheet sheet = workbook.Worksheets[sheetIndex];
        Aspose.Cells.Cells cells = sheet.Cells;
        FindOptions findOptions = new FindOptions();
        findOptions.CaseSensitive = false;
        findOptions.LookInType = LookInType.Values;
        Aspose.Cells.Cell foundCell = cells.Find(searchString, null, findOptions);
        
        if (foundCell != null)
        {
            columnIndex = foundCell.Column;   
        }else{
            columnIndex = 0;
        }
        return columnIndex;
   }

}