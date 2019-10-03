/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelread;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author hdurmaz
 */
public class ExcelRead {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
        try{
            FileInputStream file = new FileInputStream(new File("C:\\test2.xls"));// Excel dosya yolu
            HSSFWorkbook workbook = new HSSFWorkbook(file); // Excel Dosyamizi Temsil Eden Workbook Nesnesi oluşturulur.(çalışma dosyamızı belirttik)
            HSSFSheet sheet = workbook.getSheetAt(0); // Excel Dosyasının Hangi Sayfası İle Çalışacağımızı Seçtik
            Iterator<Row> rowIterator = sheet.iterator(); //Belirlediğimiz sayfa içerisinde tüm satırları tek tek dolaşacak iteratör nesnesi
            //okunacak satır olduğu sürece
             while(rowIterator.hasNext()) {
                Row row = rowIterator.next(); // Excel içerisindeki satiri temsil eden nesne
                Iterator<Cell> cellIterator = row.cellIterator();  // Her bir satir icin tum hucreleri dolasacak iterator nesnesi
                while(cellIterator.hasNext()) {
                    Cell cell = cellIterator.next(); // Excel icerisindeki hucreyi temsil eden nesne
                    
                    // Hucrede bulunan deger turunu kontrol et
                     switch(cell.getCellType()) {

                        case Cell.CELL_TYPE_BOOLEAN:                         
                            System.out.print(cell.getBooleanCellValue() + "\t\t");
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t\t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t\t");
                            break;
                      
            }

        }
        System.out.println(""); 
        }
        file.close();
        FileOutputStream out = new FileOutputStream(new File("C:\\test2.xls"));
        workbook.write(out);
        out.close();
        
    } catch (FileNotFoundException e) {

    e.printStackTrace();

} catch (IOException e) {

    e.printStackTrace();

}
    
}
}

