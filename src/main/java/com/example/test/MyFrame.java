package com.example.test;

import javax.swing.*;
import java.awt.*;

public class MyFrame extends JFrame {

    private JTextArea textArea;
    private JScrollPane scroll;

    public MyFrame() {
        super("Excel feldolgozo programocska");

        textArea = new JTextArea(10, 20);
        Font font = new Font("Consolas", Font.PLAIN, 14);
        textArea.setFont(font);

        scroll = new JScrollPane(textArea);
        add(scroll);

        setPreferredSize(new Dimension(1400, 600));
        pack();
        setLocationRelativeTo(null);
        setVisible(true);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }

    public void println(String text) {
        textArea.append(text + "\n");
        textArea.setCaretPosition(textArea.getDocument().getLength());
    }
}
