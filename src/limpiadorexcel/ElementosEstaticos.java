/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package limpiadorexcel;

import java.util.ArrayList;
import javax.swing.JOptionPane;


public class ElementosEstaticos {
    static int NumColumna=0;
    static int Index_lista=0;
    static ArrayList <javax.swing.JToggleButton> ListaBotones = new ArrayList<>();
    static ArrayList <String> ListaCadenas = new ArrayList<>();
    static ArrayList <ArrayList<String>> Matriz = new ArrayList<>();
    static ArrayList <String> BotonesSelecciones = new ArrayList<>();
    static String []palabrasnumericas={"SubTotal","IVA 16%","Descuento","Retenido IVA","Retenido ISR","Total",
                                       "Total Retenidos","IEPS"};
}
