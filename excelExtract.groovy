
import java.nio.charset.*
import java.text.SimpleDateFormat
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory


def flowFile = session.get()
if(!flowFile) return

def date = new Date()
 
flowFile = session.write(flowFile, {inputStream, outputStream ->
    try {
        Workbook wb = WorkbookFactory.create(inputStream,);
        Sheet mySheet = wb.getSheetAt(1);

        def record = ''
        def celda = ''
        def indiceP = 0
        def indiceM = 0
        def tstamp = new SimpleDateFormat("YYYY-MM-dd HH:mm:ss.SSS").format(date)
        def ultmes = "Noviembre"

        /*  Find last row  */
        int rowEnd 
        int rowIndexe=0
        for(Row row : mySheet){
            for(Cell cell : row){
                if(cell.toString()=="Total general"){
                    rowEnd = rowIndexe
                }
            }
            rowIndexe++
        }
        rowEnd = rowEnd + 2 // Le sumo 2 por las dos filas en blanco
        


        /*  Find first row  */
        int rowStart 
        int rowIndexs=0
        for(Row row : mySheet){
            for(Cell cell : row){
                if(cell.toString()=="empresa"){
                    rowStart = rowIndexs
                }
            }
            rowIndexs++
        }
        rowStart = rowStart + 3 // Le sumo 3 por las dos filas en blanco y la cabecera
        

        /*  Find last month    */
        int ultmesCol =0
        for(Cell cell : mySheet.getRow(rowStart-1)){
            if(cell.toString()==ultmes){
                ultmesCol = cell.getColumnIndex()
            }
        }
        


        /*  Extract info  */    
        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row r = mySheet.getRow(rowNum);
            if (r == null) {
                record = record + "  "
                continue;
            }
            int lastColumn = Math.max(r.getLastCellNum(), 0);
            for (int cn = 0; cn < lastColumn; cn++) {
                Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
                if(cn==ultmesCol || cn == 0){
                     if (cn == null) {
                    record = record + ", ,"
                    } else {
                        record = record + c + ","
                    }
                }
                
            }
            record = record + "\n"
        }

        outputStream.write(record.getBytes(StandardCharsets.UTF_8))
        record = ''

    }catch(e) {
     log.error("Error!: ", e.printStackTrace())
    } finally {
        if (wb != null) {
            try {
               wb.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

} as StreamCallback)

def filename = flowFile.getAttribute('filename').split('/.')[0] + '_' + new SimpleDateFormat("YYYYMMdd-HHmmss").format(date)+'.csv'
flowFile = session.putAttribute(flowFile, 'filename', filename) 

session.transfer(flowFile, REL_SUCCESS)

