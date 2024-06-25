package org.rpa;

import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.ArrayList;

public class ModalFolios extends Dialog {

    public ModalFolios(Frame owner, ArrayList<Folio> folios) {
        super(owner, "Folios", true);

        setLayout(new BorderLayout());
        TextArea textArea = new TextArea();

        for (Folio folio : folios) {
            textArea.append(folio.toString() + "\n");
        }

        add(textArea, BorderLayout.CENTER);

        Button closeButton = new Button("Close");
        closeButton.addActionListener(e -> setVisible(false));
        add(closeButton, BorderLayout.SOUTH);

        setSize(400, 300);

        addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                setVisible(false);
            }
        });
    }
}
