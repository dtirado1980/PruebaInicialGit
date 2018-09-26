
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.GroupLayout.Alignment;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.converter.WordToTextConverter;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.util.Region;


@SuppressWarnings("deprecation")
public class ReporteExcel {

	public static void main(String[] args) {
		FileOutputStream archivo = null;
		
		HSSFWorkbook wbook = null;
		HSSFSheet sheet = null;
		HSSFRow row = null;
		HSSFCell cell = null;
		HSSFCellStyle cs_tablatitulo = null;
		HSSFFont font = null;
		
		int contadorFila= 0;
		int contadorColumna= 0;
		
		String valor1="aaaaaaaaaaa";
		String valor2="bbbbbbaaaaaaaaaaa";
		String valor3="ccccccccccccccaaaaaaaaaaa";
		
		String direccion = null;
		String nombreArchivo = null;
		
		HSSFPatriarch patriarch = null;
		HSSFClientAnchor anchor = null;
		
		@SuppressWarnings("unused")
		HSSFPicture picture = null;
		String rutaImagen = null;
		try {
			
			//Definir ruta del archivo
			direccion=new String("C:\\WOG documentacion\\ejemplos poi e itext\\");
			nombreArchivo = new String("ReporteDAT.xls");
			archivo = new FileOutputStream(direccion + nombreArchivo);
			
			rutaImagen = "C:\\WOG documentacion\\logoIbroker.png";
			
			//crear un libro
			wbook = new HSSFWorkbook( );
			// crear una nueva hoja
			sheet = wbook.createSheet("Prueba Excel");
			
			// instanciamos una fila
			row = null;
			// instanciamos una celda
			cell = null;
						
			font = null;
			font = wbook.createFont();
			font.setFontHeightInPoints((short)10);
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			
			//Definimos un estilo de celda
			cs_tablatitulo = wbook.createCellStyle();
			cs_tablatitulo.setFont(font);
			cs_tablatitulo.setAlignment(HSSFCellStyle.ALIGN_JUSTIFY);
			
			contadorColumna=0;
			contadorFila=0;
			
			sheet.setColumnWidth((short) (contadorColumna), (short) ( (15 * 8) / ((double) 1 / 25)));
			contadorColumna++;
			sheet.setColumnWidth((short) (contadorColumna), (short) ( (15* 8) / ((double) 1 / 25)));
			contadorColumna++;
			sheet.setColumnWidth((short) (contadorColumna), (short) ( (20 * 8) / ((double) 1 / 25)));
			
			contadorFila+=3;
			contadorColumna= 0;
			
			row = sheet.createRow(contadorFila);
			row.setHeight((short) ( (15 * 8) / ((double) 1 / 25)));
			
			cell = row.createCell(contadorColumna);
			cell.setCellStyle(cs_tablatitulo);
		 	cell.setCellValue("FECHA DEL PROCESO");
		 	contadorColumna++;
			cell = row.createCell(contadorColumna);
			cell.setCellType((short)HSSFCell.CELL_TYPE_STRING);
			cell.setCellStyle(cs_tablatitulo);
		 	cell.setCellValue("MONEDA");
		 	contadorColumna++;
			cell = row.createCell(contadorColumna);
			cell.setCellType((short)HSSFCell.CELL_TYPE_STRING);
			cell.setCellStyle(cs_tablatitulo);
		 	cell.setCellValue("NUMERO DE EMISION");
		 	contadorColumna++;
		 	
		 	contadorFila++;
			row = sheet.createRow(contadorFila);
			contadorColumna=0;
			
			cell = row.createCell(contadorColumna);
			cell.setCellType((short)HSSFCell.CELL_TYPE_STRING);
		 	cell.setCellValue(valor1);
		 	contadorColumna++;
			cell = row.createCell(contadorColumna);
			cell.setCellType((short)HSSFCell.CELL_TYPE_STRING);
		 	cell.setCellValue(valor2);
		 	contadorColumna++;
			cell = row.createCell(contadorColumna);
			cell.setCellType((short)HSSFCell.CELL_TYPE_STRING);
		 	cell.setCellValue(valor3);
			
		 	contadorFila++;
			row = sheet.createRow(contadorFila);
			contadorColumna=0;
		 	
		 	sheet.addMergedRegion(new Region(contadorFila,(short)contadorColumna, contadorFila, (short)(contadorColumna+3)));
		 	cell = row.createCell((short)contadorColumna);
			cell.setCellType((short)HSSFCell.CELL_TYPE_STRING);
		 	cell.setCellValue("PRUEBA");
		 	
		 	
		 	patriarch = sheet.createDrawingPatriarch();
	        //int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2
		 	//cord x y cell 1 cord x y cell 2 
	        anchor = new HSSFClientAnchor(10,10,0,0,(short)1,1,(short)3,3);
	        anchor.setAnchorType( ClientAnchor.MOVE_AND_RESIZE);
	        picture = patriarch.createPicture(anchor, loadPicture(new File(rutaImagen), wbook));
		 	
		 	wbook.write(archivo);
		 	archivo.flush();
			archivo.close();
			
			System.out.println("pruebita");
			
		} catch (Exception excepcion) {
			excepcion.printStackTrace();
		}

	}
	
	public static int loadPicture( File path, HSSFWorkbook wb ) throws IOException  {  
		int indiceImagen = 0;
		FileInputStream archivoEntrada = null;
		ByteArrayOutputStream arregloDatosSalida = null;
		try{  
			// read in the image file  
			archivoEntrada = new FileInputStream(path);  
			arregloDatosSalida = new ByteArrayOutputStream( );
			int c;  
			// copy the image bytes into the ByteArrayOutputStream
			while ( (c = archivoEntrada.read()) != -1)
				arregloDatosSalida.write( c );  
				// add the image bytes to the workbook  
				indiceImagen = wb.addPicture(arregloDatosSalida.toByteArray(), HSSFWorkbook.PICTURE_TYPE_PNG );  
		} catch (Exception excepcion){
			excepcion.printStackTrace();
		}  
		finally  {  
			if (archivoEntrada != null)  
				archivoEntrada.close();  
			if (arregloDatosSalida != null)  
				arregloDatosSalida.close();  
		}  
		return indiceImagen;  
	}
	
	
	public static String ToCapitalFormat(String pParametro){
		String respuesta = null;
		try {
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		return respuesta;
	}

}





//sheet.setColumnWidth((short) (j), (short) ( (tam * 8) / ((double) 1 / 25)));
