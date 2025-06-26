/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package AsignacionSillas;

import java.io.File;
import java.util.List;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author NOEMI
 */
 class Aula {
    String nombre;
    List<Integer> sillas;
    Map<String, Integer> colegioCount = new HashMap<>();
    Map<String, List<Integer>> apellidoSillas = new HashMap<>();
    Map<String, List<Integer>> ubigeoSillas = new HashMap<>();

    public Aula(String nombre, int cantidadSillas) {
        this.nombre = nombre;
        this.sillas = new ArrayList<>();
        for (int i = 1; i <= cantidadSillas; i++) {
            sillas.add(i);
        }
        Collections.shuffle(sillas); // Mezcla inicial
    }

    public boolean puedeAsignar(String ubigeo, String apellido) {
        int countUbigeo = colegioCount.getOrDefault(ubigeo, 0);
        int countApellido = apellidoSillas.getOrDefault(apellido, new ArrayList<>()).size();

        if (countUbigeo >= 3 || countApellido >= 4) return false;

        for (int silla : sillas) {
            boolean lejanoApellido = apellidoSillas.getOrDefault(apellido, new ArrayList<>())
                .stream().allMatch(s -> Math.abs(s - silla) >= 3);
            boolean lejanoUbigeo = ubigeoSillas.getOrDefault(ubigeo, new ArrayList<>())
                .stream().allMatch(s -> Math.abs(s - silla) >= 3);
            if (lejanoApellido && lejanoUbigeo) return true;
        }
        return false;
    }

    public int asignarSilla(String ubigeo, String apellido) {
        for (int i = 0; i < sillas.size(); i++) {
            int silla = sillas.get(i);
            boolean lejanoApellido = apellidoSillas.getOrDefault(apellido, new ArrayList<>())
                .stream().allMatch(s -> Math.abs(s - silla) >= 3);
            boolean lejanoUbigeo = ubigeoSillas.getOrDefault(ubigeo, new ArrayList<>())
                .stream().allMatch(s -> Math.abs(s - silla) >= 3);

            if (lejanoApellido && lejanoUbigeo) {
                sillas.remove(i);
                colegioCount.put(ubigeo, colegioCount.getOrDefault(ubigeo, 0) + 1);
                apellidoSillas.computeIfAbsent(apellido, k -> new ArrayList<>()).add(silla);
                ubigeoSillas.computeIfAbsent(ubigeo, k -> new ArrayList<>()).add(silla);
                return silla;
            }
        }
        return -1;
    }

    public int asignarLibre() {
        if (!sillas.isEmpty()) return sillas.remove(0);
        return -1;
    }
}

public class asignarSilla extends javax.swing.JFrame {
private List<Aula> aulasCargadas = new ArrayList<>();
private List<String[]> estudiantes = new ArrayList<>();
private DefaultTableModel modelo;
private void cargarAulasDesdeExcel() {
    try {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File archivo = fileChooser.getSelectedFile();
            Workbook workbook;
            if (archivo.getName().endsWith(".xls")) {
                workbook = new HSSFWorkbook(new FileInputStream(archivo));
            } else if (archivo.getName().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(new FileInputStream(archivo));
            } else {
                JOptionPane.showMessageDialog(this, "Archivo no válido.");
                return;
            }

            Sheet hoja = workbook.getSheetAt(0);
            aulasCargadas.clear(); 

            for (int i = 1; i <= hoja.getLastRowNum(); i++) {
                Row fila = hoja.getRow(i);
                if (fila == null) continue;

                Cell celdaNombre = fila.getCell(0);
                Cell celdaCapacidad = fila.getCell(1);

                if (celdaNombre == null || celdaCapacidad == null) continue;

                String nombreAula = celdaNombre.toString().trim();
                int capacidad = 0;
    switch (celdaCapacidad.getCellType()) {
        case NUMERIC:
        capacidad = (int) celdaCapacidad.getNumericCellValue();
        break;
        case STRING:
        String valor = celdaCapacidad.getStringCellValue().trim();
        if (!valor.matches("\\d+")) {
        JOptionPane.showMessageDialog(this, "Error: Capacidad no válida en la fila " + (i + 1) + " → '" + valor + "'");
        continue;
        }
        capacidad = Integer.parseInt(valor);
        break;
        default:
        JOptionPane.showMessageDialog(this, "Tipo de dato no válido para capacidad en la fila " + (i + 1));
        continue;
}
                aulasCargadas.add(new Aula(nombreAula, capacidad));
            }

            workbook.close();
            JOptionPane.showMessageDialog(this, "Aulas cargadas correctamente.");
        }
    } catch (Exception e) {
        JOptionPane.showMessageDialog(this, "Error al cargar aulas: " + e.getMessage());
        e.printStackTrace();
    } 
    
}

private void cargarArchivoExcel() {
    try {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File archivoSeleccionado = fileChooser.getSelectedFile();
            Workbook workbook;
            if (archivoSeleccionado.getName().endsWith(".xls")) {
                workbook = new HSSFWorkbook(new FileInputStream(archivoSeleccionado));
            } else if (archivoSeleccionado.getName().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(new FileInputStream(archivoSeleccionado));
            } else {
                throw new IllegalArgumentException("El archivo no es un formato de Excel válido.");
            }
            Sheet sheet = workbook.getSheetAt(0);
            DefaultTableModel model = new DefaultTableModel();
            model = new DefaultTableModel();
            modelo=model;
            String[] columnasExcel = {"CODIGO", "Apellidos y Nombres", "DNI", "OPCION 1", "NOMBRE COLEGIO", "UBIGEO COLEGIO"};       
            String[] encabezadosTabla = {"Código", "Nombres", "DNI", "Carrera", "Colegio"," ubigeo colegio","Aula", "Silla"};
            model.setColumnIdentifiers(encabezadosTabla);
            Row filaEncabezado = sheet.getRow(0);
            Map<String, Integer> indices = new HashMap<>();
            for (int i = 0; i < filaEncabezado.getLastCellNum(); i++) {
                String nombre = filaEncabezado.getCell(i).toString().trim();
                for (String col : columnasExcel) {
                    if (nombre.equalsIgnoreCase(col)) {
                        indices.put(col, i);
                    }
                }
            }

            for (String col : columnasExcel) {
                if (!indices.containsKey(col)) {
                    throw new Exception("Columna faltante: " + col);
                }
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row fila = sheet.getRow(i);
                if (fila == null) continue;

                Object[] datosFila = new Object[columnasExcel.length];
                for (int j = 0; j < columnasExcel.length; j++) {
                    int colIndex = indices.get(columnasExcel[j]);
                    Cell celda = fila.getCell(colIndex);
                if (celda != null) {
    switch (celda.getCellType()) {
        case STRING:
            datosFila[j] = celda.getStringCellValue();
            break;
        case NUMERIC:
            datosFila[j] = String.valueOf((long) celda.getNumericCellValue());
            break;
        default:
            datosFila[j] = celda.toString();
    }
} else {
    datosFila[j] = "";
}
                }
                model.addRow(datosFila);
             estudiantes.add(new String[]{
             datosFila[1].toString(),
             datosFila[0].toString(),
             datosFila[4].toString() });
            }

            jTable1.setModel(model);
            workbook.close();
            JOptionPane.showMessageDialog(this, "Archivo cargado correctamente.");
        }
    } catch (Exception e) {
        JOptionPane.showMessageDialog(this, "Error al leer el archivo: " + e.getMessage());
        e.printStackTrace();
    }
jTable1.setModel(modelo);

}

    /**
     * Creates new form asignarSilla
     */
    public asignarSilla() {
    initComponents();

    jButton1.setText("CARGAR ARCHIVO");
    jButton1.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            cargarArchivoExcel();
        }
    });

    jButton4.setText("CARGAR AULAS Y SILLAS");
    jButton4.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            cargarAulasDesdeExcel();
        }
    });

    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jPanel2 = new javax.swing.JPanel();
        jButton1 = new javax.swing.JButton();
        jButton4 = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jButton2 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setBackground(new java.awt.Color(153, 255, 153));
        setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 0, Short.MAX_VALUE)
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 579, Short.MAX_VALUE)
        );

        jPanel2.setBackground(new java.awt.Color(255, 255, 204));
        jPanel2.setForeground(new java.awt.Color(255, 255, 153));

        jButton1.setBackground(new java.awt.Color(0, 204, 153));
        jButton1.setForeground(new java.awt.Color(0, 0, 0));
        jButton1.setText("CARGAR ARCHIVO");

        jButton4.setBackground(new java.awt.Color(0, 204, 153));
        jButton4.setForeground(new java.awt.Color(0, 0, 0));
        jButton4.setText("CARGAR AULAS Y SILLAS");

        jButton3.setBackground(new java.awt.Color(0, 204, 153));
        jButton3.setForeground(new java.awt.Color(0, 0, 0));
        jButton3.setText("ASIGNAR");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        jTable1.setBackground(new java.awt.Color(153, 255, 153));
        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "", "", "", ""
            }
        ));
        jScrollPane1.setViewportView(jTable1);

        jButton2.setBackground(new java.awt.Color(0, 204, 153));
        jButton2.setForeground(new java.awt.Color(0, 0, 0));
        jButton2.setText("EXPORTAR");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGap(33, 33, 33)
                        .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 163, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(50, 50, 50)
                        .addComponent(jButton4))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGap(276, 276, 276)
                        .addComponent(jButton3))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 1366, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGap(43, 43, 43)
                        .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 151, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap(348, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(24, 24, 24)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton4, javax.swing.GroupLayout.PREFERRED_SIZE, 36, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(jButton3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 369, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(43, 43, 43)
                .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(316, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(134, 134, 134)
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGap(397, 397, 397))
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents


    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
    modelo = (DefaultTableModel) jTable1.getModel(); // ← ¡Importante!

    int fila = 0;
    for (String[] est : estudiantes) {
    String nombre = est[0];
    String codigo = est[1];
    String ubigeo = est[2];
    String apellido = nombre.split(" ")[0]; // puedes ajustar esto
boolean asignado = false;
Collections.shuffle(estudiantes);
for (Aula aula : aulasCargadas) {
    if (aula.puedeAsignar(ubigeo, apellido)) {
        int silla = aula.asignarSilla(ubigeo, apellido);
        if (silla != -1) {
            modelo.addRow(new Object[]{codigo, nombre, "", "", "", ubigeo, aula.nombre, silla});
            asignado = true;
            break;
        }
    }
}

if (!asignado) {
    for (Aula aula : aulasCargadas) {
        int silla = aula.asignarLibre();
        if (silla != -1) {
            modelo.addRow(new Object[]{codigo, nombre, "", "", "", ubigeo, aula.nombre, silla});
            asignado = true;
            break;
        }
    }
}

if (!asignado) {
    modelo.addRow(new Object[]{codigo, nombre, "", "", "", ubigeo, "SIN AULA", "SIN SILLA"});
}
}

    jTable1.setModel(modelo); // opcional, si el modelo ya está asignado no es necesario

    }//GEN-LAST:event_jButton3ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
    try {
        Map<String, List<Object[]>> datosAgrupados = new TreeMap<>(); // TreeMap para orden alfabético
        int filas = modelo.getRowCount();
        int columnas = modelo.getColumnCount();

        for (int i = 0; i < filas; i++) {
            Object[] fila = new Object[columnas];
            for (int j = 0; j < columnas; j++) {
                fila[j] = modelo.getValueAt(i, j);
            }

            String aula = fila[modelo.findColumn("Aula")].toString();
            datosAgrupados.computeIfAbsent(aula, k -> new ArrayList<>()).add(fila);
        }

        Workbook workbook = new XSSFWorkbook();
        Sheet hoja = workbook.createSheet("Asignación");

        Row filaEncabezado = hoja.createRow(0);
        for (int i = 0; i < columnas; i++) {
            Cell celda = filaEncabezado.createCell(i);
            celda.setCellValue(modelo.getColumnName(i));
        }

        int filaExcel = 1;
        for (Map.Entry<String, List<Object[]>> entrada : datosAgrupados.entrySet()) {
            for (Object[] fila : entrada.getValue()) {
                Row row = hoja.createRow(filaExcel++);
                for (int i = 0; i < fila.length; i++) {
                    row.createCell(i).setCellValue(fila[i].toString());
                }
            }
        }

        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File archivo = fileChooser.getSelectedFile();
            if (!archivo.getName().toLowerCase().endsWith(".xlsx")) {
                archivo = new File(archivo.getAbsolutePath() + ".xlsx");
            }
            FileOutputStream out = new FileOutputStream(archivo);
            workbook.write(out);
            out.close();
            workbook.close();
            JOptionPane.showMessageDialog(this, "Archivo exportado correctamente.");
        }

    } catch (Exception e) {
        JOptionPane.showMessageDialog(this, "Error al exportar: " + e.getMessage());
        e.printStackTrace();
    }
    }//GEN-LAST:event_jButton2ActionPerformed

private void exportarTablaAExcel() {
    JFileChooser fileChooser = new JFileChooser();
    int option = fileChooser.showSaveDialog(this);

    if (option == JFileChooser.APPROVE_OPTION) {
        File archivo = fileChooser.getSelectedFile();
  
        if (!archivo.getName().toLowerCase().endsWith(".xlsx")) {
            archivo = new File(archivo.toString() + ".xlsx");
        }

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Asignación");

            DefaultTableModel model = (DefaultTableModel) jTable1.getModel();
            Row header = sheet.createRow(0);
            for (int col = 0; col < model.getColumnCount(); col++) {
                Cell cell = header.createCell(col);
                cell.setCellValue(model.getColumnName(col));
            }

            for (int row = 0; row < model.getRowCount(); row++) {
                Row fila = sheet.createRow(row + 1);
                for (int col = 0; col < model.getColumnCount(); col++) {
                    Cell cell = fila.createCell(col);
                    Object valor = model.getValueAt(row, col);
                    cell.setCellValue(valor != null ? valor.toString() : "");
                }
            }

            FileOutputStream out = new FileOutputStream(archivo);
            workbook.write(out);
            out.close();

            JOptionPane.showMessageDialog(this, "Datos exportados correctamente a:\n" + archivo.getAbsolutePath());
        } catch (IOException e) {
            JOptionPane.showMessageDialog(this, "Error al exportar: " + e.getMessage());
            e.printStackTrace();
        }
    }
}


    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(asignarSilla.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(asignarSilla.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(asignarSilla.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(asignarSilla.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new asignarSilla().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
}
