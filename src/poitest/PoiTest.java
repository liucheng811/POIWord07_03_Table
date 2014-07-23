package poitest;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

/**
 * Created by WangHao on 22/07/2014.
 */

public class PoiTest extends JFrame {
    static JTextField TextField;
    static PoiTest testFrame;

    public static void main(String[] args) {

        testFrame = new PoiTest();
        TextField = new JTextField("                                                                       ");
        JButton button = new JButton("ChooseFile");
        button.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                //choose file
                String filePath = "";
                String fileStyle = "";
                JFileChooser chooser = new JFileChooser("./");
                FileNameExtensionFilter filter = new FileNameExtensionFilter(
                        "doc/docx", "doc", "docx");         //only choose doc or docx
                chooser.setFileFilter(filter);
                int returnVal = chooser.showOpenDialog(testFrame);
                if (returnVal == JFileChooser.APPROVE_OPTION) {
                    TextField.setText(chooser.getSelectedFile().getAbsolutePath());
                    filePath = (String) chooser.getSelectedFile().getAbsolutePath(); //get file path
                    System.out.println(filePath);
                }

                int dot = filePath.lastIndexOf(".");
                if ((dot > -1) && (dot < (filePath.length() - 1))) {
                    fileStyle = filePath.substring(dot + 1);    //The types of file only are doc or docx
                }
                if (fileStyle.equals("docx")) {
                    wordDocx(filePath);         //is word2007
                } else if (fileStyle.equals("doc")) {
                    wordDoc(filePath);    //is word2003
                }

            }
        });
        Container contentPane = testFrame.getContentPane();
        contentPane.setLayout(new FlowLayout());
        contentPane.add(button);
        contentPane.add(TextField);
        testFrame.setSize(300, 100);
        testFrame.setVisible(true);

    }

    //word2003
    public static void wordDoc(String wordPath) {
        try {
            FileInputStream in = new FileInputStream(wordPath);// Input the file
            POIFSFileSystem pfs = new POIFSFileSystem(in);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            Range range = hwpf.getRange();// get the range of file
            TableIterator it = new TableIterator(range);// Iterate the table
            while (it.hasNext()) {
                org.apache.poi.hwpf.usermodel.Table tb = (org.apache.poi.hwpf.usermodel.Table) it.next();
                // Iterate row,the default starts with 0
                for (int i = 0; i < tb.numRows(); i++) {
                    org.apache.poi.hwpf.usermodel.TableRow tr = tb.getRow(i);
                    // Iterate column,the default starts with 0
                    for (int j = 0; j < tr.numCells(); j++) {
                        TableCell td = tr.getCell(j);
                        // Obtain the contents of the cell
                        for (int k = 0; k < td.numParagraphs(); k++) {
                            Paragraph para = td.getParagraph(k);
                            String s = para.text().trim();
                            System.out.print(s + "  ");
                        } // end for
                    } // end for
                    System.out.println();
                } // end for
            } // end while
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //word2007
    public static void wordDocx(String wordPath) {
        try {
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(wordPath));  // Input the file
            Iterator<XWPFTable> it = document.getTablesIterator();//Iterate the table
            while (it.hasNext()) {
                XWPFTable table = (XWPFTable) it.next();
                int count = table.getNumberOfRows();     //get the row
                for (int i = 0; i < count; i++) {
                    XWPFTableRow row = table.getRow(i);         //locate the number of row
                    List<XWPFTableCell> cells = row.getTableCells();  //get the contents
                    for (XWPFTableCell cell : cells) {
                        System.out.print(cell.getText() + "  ");
                    }
                    System.out.println();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}