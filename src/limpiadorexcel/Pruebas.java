/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package limpiadorexcel;

import java.util.ArrayList;
import javax.swing.JFrame;
import javax.swing.JToggleButton;

/**
 *
 * @author eddyp
 */
public class Pruebas {
    static JFrame ventana = new JFrame("Ventana de prueba");
    static javax.swing.JToggleButton Boton = new JToggleButton("Hola");

    public Pruebas() {
        ventana.setSize(300,300);
        ventana.setVisible(true);
        ventana.add(Boton);
        Boton.setVisible(true);
    }
    
    
    public static void main (String args[]){
        ventana.setVisible(true);
        ventana.add(Boton);
        Boton.setVisible(true);
        Boton.setAlignmentX(30);
        Boton.setAlignmentY(30);
        if (Boton.isValid()){
            System.out.println("Boton activado");
        }
    }
}
