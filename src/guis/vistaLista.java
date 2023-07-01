package guis;

import java.awt.EventQueue;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.File;

import java.io.IOException;

import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;

import java.util.Date;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.time.LocalDate;
import java.util.Locale;
import javax.imageio.ImageIO;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;


import Models.ModeloExcel;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import javax.swing.JLabel;
import javax.swing.JTextField;
import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;

public class vistaLista extends JFrame {
	private JButton btnNewButton;
	private JScrollPane scrollPane;
	private JTable tablaDatos;
	private JButton btnNewButton_1;
	public String NombreActividad;
	public Date FechaActividad;
	public String Ubicacion ;
	public String nombreMes ;
	public String diaSemana ;
	public int anio ;
	public int diaNumero;
	
	public String HrasTrabajo;
	private JLabel lblNewLabel;
	private JLabel lblFechaActividad;
	private JLabel lblNewLabel_2;
	private JLabel lblNewLabel_3;
	private JTextField txtNombreActividad;
	private JTextField txtFechaActividad;
	private JTextField txtUbicacion;
	private JTextField txtHrsActividad;
	
	
	

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					vistaLista frame = new vistaLista();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public vistaLista() {
		setTitle("Vista");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 792, 483);
		getContentPane().setLayout(null);
		
		btnNewButton = new JButton("Importar");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				importar(e);
			}
		});
		btnNewButton.setBounds(23, 11, 164, 50);
		getContentPane().add(btnNewButton);
		
		scrollPane = new JScrollPane();
		scrollPane.setBounds(10, 145, 756, 288);
		getContentPane().add(scrollPane);
		
		tablaDatos = new JTable(new DefaultTableModel());
		tablaDatos.setFillsViewportHeight(true);
		scrollPane.setViewportView(tablaDatos); 
		
		btnNewButton_1 = new JButton("GenerarConstancias");
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				generarPDF(e);
			}
		});
		btnNewButton_1.setBounds(602, 84, 164, 50);
		getContentPane().add(btnNewButton_1);
		
		lblNewLabel = new JLabel("Nombre Actibvidad");
		lblNewLabel.setFont(new Font("Bodoni MT Black", Font.PLAIN, 11));
		lblNewLabel.setForeground(new Color(112, 128, 144));
		lblNewLabel.setBounds(197, 11, 110, 20);
		getContentPane().add(lblNewLabel);
		
		lblFechaActividad = new JLabel("Fecha Actividad");
		lblFechaActividad.setBounds(197, 41, 110, 20);
		getContentPane().add(lblFechaActividad);
		
		lblNewLabel_2 = new JLabel("Ubicacion");
		lblNewLabel_2.setBounds(197, 75, 110, 20);
		getContentPane().add(lblNewLabel_2);
		
		lblNewLabel_3 = new JLabel("Hrs de Actividad");
		lblNewLabel_3.setBounds(197, 114, 110, 20);
		getContentPane().add(lblNewLabel_3);
		
		txtNombreActividad = new JTextField();
		txtNombreActividad.setText("REFORESTACION EN CARABAYLLO \"MANOS A LA PLANTA\"");
		txtNombreActividad.setBounds(329, 11, 140, 20);
		getContentPane().add(txtNombreActividad);
		txtNombreActividad.setColumns(10);
		
		txtFechaActividad = new JTextField();
		txtFechaActividad.setText("22/04/2023");
		txtFechaActividad.setColumns(10);
		txtFechaActividad.setBounds(329, 41, 140, 20);
		getContentPane().add(txtFechaActividad);
		
		txtUbicacion = new JTextField();
		txtUbicacion.setText("Carabayllo, Lima");
		txtUbicacion.setColumns(10);
		txtUbicacion.setBounds(329, 75, 140, 20);
		getContentPane().add(txtUbicacion);
		
		txtHrsActividad = new JTextField();
		txtHrsActividad.setText("4");
		txtHrsActividad.setColumns(10);
		txtHrsActividad.setBounds(329, 114, 140, 20);
		getContentPane().add(txtHrsActividad);


	}
	protected void importar(ActionEvent e) {
		
		JFileChooser fileChooser = new JFileChooser();
	    int result = fileChooser.showOpenDialog(this);
	    try {
		    if (result == JFileChooser.APPROVE_OPTION) {
		        File selectedFile = fileChooser.getSelectedFile();
		        
		        ModeloExcel modeloExcel = new ModeloExcel();
		        String respuesta = modeloExcel.importar(selectedFile, tablaDatos);
		        
		        // Puedes mostrar un mensaje de respuesta al usuario según el resultado de la importación
		        JOptionPane.showMessageDialog(this, respuesta);
		    }
		} catch (Exception e2) {
			e2.printStackTrace();
		}

		
		
		
	}
	protected void generarPDF(ActionEvent e) {
		
		BufferedImage imagenFondo = seleccionarImagen();

	    // Verificar si se seleccionó una imagen
	    if (imagenFondo == null) {
	        	JOptionPane.showMessageDialog(null, "No se seleccionó una imagen de fondo.", "Error", JOptionPane.ERROR_MESSAGE);
	        return;
	    }
		
		try {
			
			this.NombreActividad = txtNombreActividad.getText();
			String fechaString = txtFechaActividad.getText();
	        DateTimeFormatter formato = DateTimeFormatter.ofPattern("dd/MM/yyyy", new Locale("es"));

			try {
				LocalDate fecha = LocalDate.parse(fechaString, formato);
	            this.diaNumero = fecha.getDayOfMonth();
	
	            this.nombreMes = fecha.getMonth().getDisplayName(TextStyle.FULL, new Locale("es"));
	            this.diaSemana = fecha.getDayOfWeek().getDisplayName(TextStyle.FULL, new Locale("es"));
	            this.anio = fecha.getYear();
	            
	            
	            
	        } catch (Exception e1) {
	            e1.printStackTrace();
	        }
			
			

			
			this.HrasTrabajo = txtHrsActividad.getText();
			this.Ubicacion = txtUbicacion.getText();
			
			generarPDFsDesdeTabla(tablaDatos,imagenFondo);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		
	}
	
	private BufferedImage seleccionarImagen() {
	    BufferedImage imagen = null;

	    // Abre un cuadro de diálogo de selección de archivo
	    JFileChooser fileChooser = new JFileChooser();
	    int result = fileChooser.showOpenDialog(null);

	    if (result == JFileChooser.APPROVE_OPTION) {
	        try {
	            // Lee la imagen seleccionada y la guarda en un objeto BufferedImage
	            File selectedFile = fileChooser.getSelectedFile();
	            imagen = ImageIO.read(selectedFile);
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }

	    return imagen;
	}

	public void generarPDFsDesdeTabla(JTable tabla,BufferedImage imagen) throws IOException {
	    DefaultTableModel modelo = (DefaultTableModel) tabla.getModel();
	    
	    JFileChooser fileChooser = new JFileChooser();
	    fileChooser.setDialogTitle("Seleccionar carpeta de guardado");
	    fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
	    
	    int seleccion = fileChooser.showOpenDialog(this);
	    if (seleccion == JFileChooser.APPROVE_OPTION) {
	        File carpetaGuardar = fileChooser.getSelectedFile();
	        String outputPath = carpetaGuardar.getAbsolutePath();
	        
	        for (int i = 0; i < modelo.getRowCount(); i++) {

	        	
	        	 Object nombre = modelo.getValueAt(i, 0);
	             /*Object apellido = modelo.getValueAt(i, 1);*/
	             /*Object dni = modelo.getValueAt(i, 2);*/

	             if (nombre != null) {


	              
	    
	                 // Crear documento PDF
                    PDDocument document = new PDDocument();
                    PDPage page = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));
                    document.addPage(page);

	                 
                    PDImageXObject backgroundImage = LosslessFactory.createFromImage(document, imagen);

                    // Obtener tamaño de página y escalar imagen de fondo
                    PDRectangle pageSize = page.getMediaBox();
                    float scaleFactor = Math.min(pageSize.getWidth() / backgroundImage.getWidth(),
                            pageSize.getHeight() / backgroundImage.getHeight());
                    float scaledWidth = backgroundImage.getWidth() * scaleFactor;
                    float scaledHeight = backgroundImage.getHeight() * scaleFactor;


                    // Crear contenido para el PDF
                    try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                        // Colocar imagen de fondo en la página
                        contentStream.drawImage(backgroundImage, 0, 0, scaledWidth, scaledHeight);




                        // Agregar nombre y apellido
                        contentStream.setFont(PDType1Font.HELVETICA_BOLD,30);
                        contentStream.beginText();
                        contentStream.setNonStrokingColor(new Color(3,98,44));
                        /*String nombrCompleto = nombre + " " + apellido;*/
                        String nombrCompleto = nombre + "" ;
                        float tamanioNombre = PDType1Font.HELVETICA_BOLD.getStringWidth(nombrCompleto)/ 1000f * (30);
                        float xN = (page.getMediaBox().getWidth()- tamanioNombre)/2;	
                        
                        contentStream.newLineAtOffset(xN,335);
                        contentStream.showText(nombrCompleto);
                        contentStream.endText();

                        
                        //Imprimir parrafo
                        contentStream.setFont(PDType1Font.HELVETICA, 16);

                        contentStream.beginText();
                        contentStream.setNonStrokingColor(new Color(0, 0, 0));
                        contentStream.newLineAtOffset(90, 285);
                        contentStream.showText("En reconocimiento a su arduo trabajo y valioso aporte personal como voluntario para el");
                        contentStream.endText();

                        contentStream.beginText();
                        contentStream.setNonStrokingColor(new Color(0, 0, 0));
                        contentStream.newLineAtOffset(90, 265);
                        contentStream.showText("éxito de la actividad");
                        contentStream.setStrokingColor(new Color(119, 200, 84));
                        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 16);
                        contentStream.showText(" " + this.NombreActividad);
                        contentStream.setNonStrokingColor(new Color(0, 0, 0));
                        contentStream.showText(",");
                        contentStream.endText();

                        contentStream.beginText();
                        contentStream.setFont(PDType1Font.HELVETICA_BOLD_OBLIQUE, 16);
                        contentStream.setNonStrokingColor(new Color(0, 0, 0));
                        contentStream.newLineAtOffset(90, 245);
                        contentStream.showText("realizada el " + this.diaSemana + " de " + this.nombreMes);
                        contentStream.setFont(PDType1Font.HELVETICA, 16);
                        contentStream.showText(" en " + this.Ubicacion + ". La constancia a continuación se");
                        contentStream.endText();

                        contentStream.beginText();
                        contentStream.setFont(PDType1Font.HELVETICA, 16);
                        contentStream.setNonStrokingColor(new Color(0, 0, 0));
                        contentStream.newLineAtOffset(90, 225);
                        contentStream.showText("emite por un total de " + this.HrasTrabajo + " horas de trabajo voluntario.");
                        contentStream.endText();


						//Agregar Fecha de Actividad
                        
                        contentStream.setFont(PDType1Font.HELVETICA, 16);
                        contentStream.beginText();
                        contentStream.setNonStrokingColor(new Color(37,37,35));

                        contentStream.newLineAtOffset(60, 65);
                        contentStream.showText("Lima, " + this.nombreMes +" de " + this.anio);
                        contentStream.endText();
                    }
                    
                    // Guardar PDF
                    /*String outputFilename = outputPath + File.separator + "PDF" + nombre + apellido + ".pdf";*/
                    String outputFilename = outputPath + File.separator + "PDF" + nombre  + ".pdf";
                    document.save(outputFilename);
                    document.close();
                    
                    
	            }
	        }
	        
	        JOptionPane.showMessageDialog(this, "Los PDFs han sido generados y guardados en la carpeta seleccionada.");
	    }else {
	    	
	    	JOptionPane.showMessageDialog(this, "Error");
	    	
	    }
	}
}
