import javax.swing.*;
import java.awt.*;
import java.io.File;

public class TestFilePicker extends JFrame {
    public TestFilePicker() {
        super("Test using JFilePicker");

        setLayout(new FlowLayout());

        // set up a file picker component
        JFilePicker filePicker = new JFilePicker("Pick a file", "Browse...");
        filePicker.setMode(JFilePicker.MODE_SAVE);
        filePicker.addFileTypeFilter(".jpg", "JPEG Images");
        filePicker.addFileTypeFilter(".mp4", "MPEG-4 Videos");

        // access JFileChooser class directly
        JFileChooser fileChooser = filePicker.getFileChooser();
        fileChooser.setCurrentDirectory(new File("D:/"));

        // add the component to the frame
        add(filePicker);

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(520, 100);
        setLocationRelativeTo(null);    // center on screen
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new TestFilePicker().setVisible(true);
            }
        });
    }
}
