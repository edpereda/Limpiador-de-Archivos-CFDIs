
package limpiadorexcel;

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.Dimension;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JToggleButton;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.*;
import jxl.CellFeatures;
import jxl.*;
import static limpiadorexcel.ElementosEstaticos.Matriz;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;


public class GestorExcel {
    public static void llenarArreglos (File ArchivoExcel) throws FileNotFoundException, IOException{
        InputStream excelStream = null;
        
            excelStream = new FileInputStream(ArchivoExcel);
            
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(excelStream);

            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
            XSSFRow xssfRowTitulos = xssfSheet.getRow(0);
            XSSFRow xssfRow = xssfSheet.getRow(0);
            XSSFRow xssfRow2 = xssfSheet.getRow(0);
            XSSFCell cell;     
            
            int rows = xssfSheet.getLastRowNum()+2;
            int cols = 0;            
            String ValorTitulo;
            int contadorrenglones=1; 
            ElementosEstaticos.ListaBotones.clear();
            VentanaPrincipal.PanelBotones.removeAll();
            //-----------------------------vamos a leer los titulos de cada columna en el primer renglon
            //JOptionPane.showMessageDialog(null, rows);
            try{
                for (int columna=0;columna<=xssfRowTitulos.getLastCellNum();columna++){
                ValorTitulo="";
                    if (xssfRowTitulos.getCell(columna)==null){
                        break;
                    }else{
                        ValorTitulo=xssfRowTitulos.getCell(columna).toString();
                        System.out.println("ValorTitulo: "+ValorTitulo);
                        ElementosEstaticos.ListaBotones.add(new JToggleButton(ValorTitulo));            /*En esta parte añade al arreglo un boton y lo pone en la ventana principal*/
                        //JOptionPane.showMessageDialog(null, ElementosEstaticos.ListaBotones.get(columna).getText());
                        
                        VentanaPrincipal.PanelBotones.add(ElementosEstaticos.ListaBotones.get(columna));
                        ElementosEstaticos.ListaBotones.get(columna).setVisible(true);
                    }
                }
            }catch(Exception e){
                System.out.println("Error:"+e);
            }
            
    }
    
    /*public static void main (String args[]) throws IOException{
        llenarArreglos(new File("C:\\Users\\eddyp\\Desktop\\año 2017.eduardoxlsx.xlsx"));
    }*/

    static void LlenarColumna(String NombreBoton,File ArchivoExcel)throws FileNotFoundException, IOException{
        //JOptionPane.showMessageDialog(null, "Nombre que llego: "+NombreBoton);
        InputStream excelStream = null;

        excelStream = new FileInputStream(ArchivoExcel);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(excelStream);

        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
        XSSFRow xssfRowTitulos = xssfSheet.getRow(0);
        XSSFRow xssfRow = xssfSheet.getRow(0);
        XSSFRow xssfRow2 = xssfSheet.getRow(0);
        XSSFCell cell;     
            
        int rows = xssfSheet.getLastRowNum();
        int cols = 0;            
        String ValorTitulo;
        String ValorCelda="";
        int contadorrenglones=1; 
            
            //-----------------------------vamos a leer los titulos de cada columna en el primer renglon
            //JOptionPane.showMessageDialog(null, rows);
            try{
                for (int columna=0;columna<=xssfRowTitulos.getLastCellNum();columna++){
                    if (xssfRowTitulos.getCell(columna)==null){
                        break;
                    }else{
                        if (xssfRowTitulos.getCell(columna).toString().equals(NombreBoton)){
                            for (int renglon=0;renglon<(rows);renglon++){
                                if (xssfSheet.getRow(renglon)==null){
                                    ElementosEstaticos.ListaCadenas.add("");
                                }else{
                                    xssfRow= xssfSheet.getRow(renglon);
                                    ValorCelda= xssfRow.getCell(columna).toString();
                                    System.out.println("Valor Celda"+ValorCelda);
                                    ElementosEstaticos.ListaCadenas.add(ValorCelda); 
                                    ValorCelda="";
                                }
                            }
                        }
                        
                    }
                }
            excelStream.close();
            xssfWorkbook.close();
            }catch(Throwable e){
                
            }
            //JOptionPane.showMessageDialog(null, "Tamaño de Lista: "+ElementosEstaticos.ListaCadenas.size());
            ElementosEstaticos.Matriz.add(new ArrayList<>(ElementosEstaticos.ListaCadenas));
            //JOptionPane.showMessageDialog(null, "Tamaño de Matriz: "+ElementosEstaticos.Matriz.size());
            ElementosEstaticos.ListaCadenas.clear();
    }
    
    static void ImprimirArreglo (){
        for (int i=0;i<ElementosEstaticos.ListaCadenas.size();i++){
            System.out.println("****"+ElementosEstaticos.ListaCadenas.get(i).toString());
        }
    }
    
    public static void imprimirMatriz (){
        JOptionPane.showMessageDialog(null, "TAM MATRIZ: "+ElementosEstaticos.Matriz.size()+" TAM LISTA: "+ElementosEstaticos.Matriz.get(0).size());
        for (int i=0;i<Matriz.size();i++){
            for (int x=0;x < Matriz.get(i).size();x++){
                JOptionPane.showMessageDialog(null, Matriz.get(i).get(x).toString());
                System.out.println("MATRIZ--  "+Matriz.get(i).get(x).toString());
            }
        }
    }
    static void LlenarColumnaExcel(File ruta) throws IOException, WriteException{            //por medio de una ruta dada genera el archivo excel
        
        int index=0;
        double sumatoria=0;
        boolean sumar=false;
        
        WorkbookSettings conf=new WorkbookSettings();
        conf.setEncoding("ISO-8859-1");
        WritableWorkbook workbook = Workbook.createWorkbook(ruta,conf);
        
        WritableSheet sheet= workbook.createSheet("Hoja Limpia de Excel", 0);
        WritableFont fuentetitulos= new WritableFont(WritableFont.ARIAL, 9 , WritableFont.BOLD);
        WritableFont fuentedatos= new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD);
        
        WritableCellFormat formatitulos= new WritableCellFormat(fuentetitulos);
        WritableCellFormat formatdatos= new WritableCellFormat(fuentedatos);
        //JOptionPane.showMessageDialog(null, "conceptos: "+Arreglos.Conceptos);
        
        try {
            do{
                do{
                    sheet.setColumnView(ElementosEstaticos.NumColumna, 25);
                    if (index==0){
                        //JOptionPane.showMessageDialog(null, "Entro a index=0");
                        for(short i=0;i<ElementosEstaticos.palabrasnumericas.length;i++){//aqui hace un for para recorrer el arreglo y si una palabra es igual a la de los arreglos me manda true
                            if ((ElementosEstaticos.Matriz.get(ElementosEstaticos.Index_lista).get(index).toString()).equals(ElementosEstaticos.palabrasnumericas[i])){
                                //JOptionPane.showMessageDialog(null, "Coincide: ");
                                sumar=true;
                                sheet.setColumnView(ElementosEstaticos.NumColumna, 10);
                                break;
                            }
                        }
                    }
                    
                    if (index==0){
                        sheet.addCell(new jxl.write.Label(ElementosEstaticos.NumColumna,index,ElementosEstaticos.Matriz.get(ElementosEstaticos.Index_lista).get(index).toString(),formatitulos));
                    }else{
                        sheet.addCell(new jxl.write.Label(ElementosEstaticos.NumColumna,index,ElementosEstaticos.Matriz.get(ElementosEstaticos.Index_lista).get(index).toString(),formatdatos));
                    }
                    
                    
                    if ((sumar)&&(index!=0)){
                        if (ElementosEstaticos.Matriz.get(ElementosEstaticos.Index_lista).get(index).toString().equals("")){
                            
                        }else{
                            sumatoria=sumatoria+(Double.parseDouble(ElementosEstaticos.Matriz.get(ElementosEstaticos.Index_lista).get(index).toString()));
                            //JOptionPane.showMessageDialog(null, "Sumatoria: "+sumatoria);
                        }
                    }
                    
                    System.out.println("Hoja NUEVA: "+ElementosEstaticos.Matriz.get(ElementosEstaticos.Index_lista).get(index).toString());
                    index++;
                
                    if (sumar){
                        //JOptionPane.showMessageDialog(null, "solo una vez");
                        if (index==ElementosEstaticos.Matriz.get(ElementosEstaticos.Index_lista).size()){
                            //JOptionPane.showMessageDialog(null, "La sumatoria mas perroa: "+sumatoria);
                            sheet.addCell(new jxl.write.Number(ElementosEstaticos.NumColumna,(index+1),sumatoria,formatdatos));
                        }
                    }
                }while(index!=ElementosEstaticos.Matriz.get(ElementosEstaticos.Index_lista).size());
                
                
                ElementosEstaticos.NumColumna+=1;
                ElementosEstaticos.Index_lista+=1;
                index=0;
                sumatoria=0;
                sumar=false;
                
            }while(ElementosEstaticos.Index_lista!=ElementosEstaticos.Matriz.size());
            ElementosEstaticos.NumColumna=0;
            ElementosEstaticos.Index_lista=0;
        
        }catch(Throwable e){
            
        }
        workbook.write();
        workbook.close();
    }
    
    public static File obtenerArchivo (){
        JFileChooser Filechooser = new JFileChooser();
       
        Filechooser.setPreferredSize(new Dimension(1000, 1000));
        String ruta = null;
        File path = null;
        
        Filechooser.setFileFilter(new FileNameExtensionFilter("Excel (*.xlsx)", "xlsx"));   //asi se ponen filtros sobre encontrar tipos de archivos
        if (Filechooser.showDialog(null, "Seleccionar")==JFileChooser.APPROVE_OPTION){
            
                path=Filechooser.getSelectedFile();
                VentanaPrincipal.mostrarBotones=true;
                return path;       
        }else{
            JOptionPane.showMessageDialog(null, "No se selecciono archivo");
            
        }
        VentanaPrincipal.mostrarBotones=false;
        return null;
    }
    
    public static File obtenerRutaGuardado () throws IOException{                     //genera un file chooser el cual le da una direccion para crear un archivo excel
        JFileChooser Filechooser = new JFileChooser();
        File path = null;
        
        Filechooser.setFileFilter(new FileNameExtensionFilter("Excel (*.xls)", "xls"));   //asi se ponen filtros sobre encontrar tipos de archivos
        
        if (Filechooser.showDialog(null, "Guardar")==JFileChooser.APPROVE_OPTION){
            
                path=Filechooser.getSelectedFile();
                File ruta=new File(String.valueOf(path)+".xls");
                path=ruta;
        }
    return path;    
    }
    
    public static void generarArchivoTexto (File direccion){
        FileWriter fw;
        BufferedWriter bw;
        PrintWriter pw;
        try{
                fw= new FileWriter(direccion);
                bw= new BufferedWriter(fw);
                pw = new PrintWriter(direccion);

                for (int index=0;index<ElementosEstaticos.BotonesSelecciones.size();index++){
                    String botonseleccionado = ElementosEstaticos.BotonesSelecciones.get(index)+"-";
                    
                    pw.append(botonseleccionado);  //Imprime el String dentro del archivo txt

                }
                pw.close();
                bw.close();
                fw.close();
            }catch(Throwable e){
                JOptionPane.showMessageDialog(null, "ERROR EN GENERAR ARCHIVO");
            }
    }
    /**-----------------------------------------------------------------arreglar este archivo para las configuraciones*/
    public static void PresionarBotonesConfiguracion (File direccion){
        
        FileReader fr;
        BufferedReader br;
        String cadena;
        String []Usuario = null;
        int n=0;
        
        for (int i =0;i<ElementosEstaticos.ListaBotones.size();i++){
            ElementosEstaticos.ListaBotones.get(i).setSelected(false);
        }
        try{
            fr = new FileReader(direccion);
            br = new BufferedReader(fr);

            while((cadena=br.readLine())!=null){
                Usuario=cadena.split("-");
            }       
            
            for (int i = 0;i<Usuario.length;i++){
                for ( int k = n;k<ElementosEstaticos.ListaBotones.size();k++){
                    //JOptionPane.showMessageDialog(null, Usuario[i].toString()+"  "+ElementosEstaticos.ListaBotones.get(k).getText());
                    if (ElementosEstaticos.ListaBotones.get(k).getText().equals(Usuario[i])){
                        //JOptionPane.showMessageDialog(null, "COINCIDE");
                        ElementosEstaticos.ListaBotones.get(k).setSelected(true);
                        n=k;
                        k=ElementosEstaticos.ListaBotones.size()-1;   
                    }
                }
            }
        br.close();
        fr.close();
        }catch(Throwable e){
            JOptionPane.showMessageDialog(null, "ERROR EN CUADROS VENTANA");
        }
    }
}
