package Models;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.FileOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.io.IOException;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell ;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;


public class ModeloExcel {
	Workbook wb;
    DefaultTableModel modelo;
    String rutaGuardar; // Variable para almacenar la ruta de guardado de los PDF


	public String importar(File archivo,JTable tabla) {
		String respuesta = "No Se puede importar";
		DefaultTableModel modelo = new DefaultTableModel();
		tabla.setModel(modelo);
		
		try {
			
			wb = WorkbookFactory.create(new FileInputStream(archivo));
            Sheet sheet = wb.getSheetAt(0); // Obtener la hoja de trabajo (worksheet)
            Iterator filaIterator =sheet.rowIterator();
            int indiceFila =-1;
            
            while(filaIterator.hasNext()) {
            	indiceFila++;
            	
            	Row fila = (Row)filaIterator.next();
            	Iterator columnaIterator = fila.cellIterator();
            	Object[] listaColumna = new Object[5];
            	
            	int indiceColumna =-1;
            	while(columnaIterator.hasNext()) {
            		indiceColumna++;
            		Cell celda = (Cell)columnaIterator.next();
            		if (indiceFila == 0) {
						modelo.addColumn(celda.getStringCellValue());
					}else {
						if (celda!=null) {
							
							switch(celda.getCellType()){
                            case NUMERIC:
                                listaColumna[indiceColumna]= (int)Math.round(celda.getNumericCellValue());
                                break;
                            case STRING:
                                listaColumna[indiceColumna]= celda.getStringCellValue();
                                break;
                            case BOOLEAN:
                                listaColumna[indiceColumna]= celda.getBooleanCellValue();
                                break;
                            default:
                                listaColumna[indiceColumna]=celda.getDateCellValue();
                                break;
                        }
                            System.out.println("col"+indiceColumna+" valor: true - "+celda+".");

							
						}
						
						
					}
            		
            		
            		
            	}
            	if(indiceFila!=0)modelo.addRow(listaColumna);
            	
            	
            }
            		
            
            
            respuesta = "Importación exitosa";
        } catch (Exception e) {
            System.err.println(e.getMessage());
        }
		
		return respuesta;
	}
	
	
}
