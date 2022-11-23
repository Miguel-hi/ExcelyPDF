package Creacion;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import static javax.swing.WindowConstants.EXIT_ON_CLOSE;
/////////////////////////////////////////////////////////////////
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * ************************************************************
 */

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

/**
 *
 * @author migue
 */
/**
 * *****************************************************************
 */
public class Crear extends JFrame {

    private JPanel panel;

    public Crear() {

        this.setSize(550, 350);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setTitle("<<<<EXCEL o PDF>>>>");
        setLocationRelativeTo(null);
        setMinimumSize(new Dimension(200, 200));

        Iniciar();

    }

    private void Iniciar() {

        Panel();
        Etiquetas();
        BotonExcel();
        BotonPDF();
    }

    private void Panel() {

        panel = new JPanel();
        panel.setLayout(null);
        this.getContentPane().add(panel);
    }

    private void Etiquetas() {

        JLabel etiqueta = new JLabel("Escoge EXCEL o PDF");//texto etiqueta
        panel.add(etiqueta);//agregamos un panel a la etiqueta
        etiqueta.setBounds(150, 70, 300, 80);//dimension de la etiqueta x,y,ancho,alto
        etiqueta.setForeground(Color.blue);//color del texto
        etiqueta.setFont(new Font("Baskerville", Font.ITALIC, 22));//tipo de fuente y tamaño de la letra

    }

    private void BotonExcel() {
        JButton boton = new JButton("Excel");
        boton.setBounds(90, 250, 150, 40);
        boton.setForeground(Color.blue);
        boton.setFont(new Font("Baskerville", Font.CENTER_BASELINE, 15));
        boton.setEnabled(true);
        panel.add(boton);

        boton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent arg0) {

                Workbook libro = new XSSFWorkbook();
                final String nombreArchivo = "ejemplo excel.xlsx";
                Sheet hoja = libro.createSheet("Hoja 1");
                Row primeraFila = hoja.createRow(0);
                Cell primeraCelda = primeraFila.createCell(0);
                primeraCelda.setCellValue("Yo voy en la primera celda y primera fila");
                File directorioActual = new File(".");
                String ubicacion = directorioActual.getAbsolutePath();
                String ubicacionArchivoSalida = ubicacion.substring(0, ubicacion.length() - 1) + nombreArchivo;
                FileOutputStream outputStream;
                try {
                    outputStream = new FileOutputStream(ubicacionArchivoSalida);
                    libro.write(outputStream);
                    libro.close();
                    System.out.println("Libro guardado correctamente");
                } catch (FileNotFoundException ex) {
                    System.out.println("Error de filenotfound");
                } catch (IOException ex) {
                    System.out.println("Error de IOException");
                }

                Excel excel = new Excel();
                excel.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                excel.setVisible(true);

                setVisible(false);

            }
        });

    }

    private void BotonPDF() {
        JButton boton = new JButton("PDF");
        boton.setBounds(300, 250, 150, 40);
        boton.setForeground(Color.blue);
        boton.setFont(new Font("Baskerville", Font.CENTER_BASELINE, 15));
        boton.setEnabled(true);
        panel.add(boton);

        boton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent arg0) {

 
                String imagen = "C:\\Users\\migue\\Desktop\\komi.jpg";
                try {
                    PDDocument documento = new PDDocument();
                    PDPage pagina = new PDPage(PDRectangle.LETTER);
                    documento.addPage(pagina);

                    PDImageXObject image = PDImageXObject.createFromFile(imagen, documento);
                    PDPageContentStream contenido = new PDPageContentStream(documento, pagina);

                    contenido.drawImage(image, 20f, 20f);

                    contenido.beginText();
                    contenido.setFont(PDType1Font.TIMES_BOLD, 17);
                    contenido.setLeading(14.5f);
                    contenido.newLineAtOffset(20, pagina.getMediaBox().getHeight() - 52);
                    contenido.showText("Hola, este es mi PDF");
                    
                    contenido.endText();

                    contenido.close();

                    documento.save("C:\\Users\\migue\\OneDrive - Universidad Autónoma del Estado de México\\Documents\\NetBeansProjects\\ProyectoParadigmas\\ejemplo.pdf");

                } catch (Exception x) {
                    System.out.println("Error: " + x.getMessage().toString());
                }

                /////////////////////////////////////////////////////////////////   
                PDF p = new PDF();
                p.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                p.setVisible(true);

                setVisible(false);

            }
        });

    }
}
