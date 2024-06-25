package org.rpa;

import java.awt.*;
import java.io.IOException;

public class Main {
    static Robot bot;

    public static void main(String[] args) throws AWTException {
        // Start the key press listener thread
        KeyPressListener keyPressListener = new KeyPressListener();
        keyPressListener.start();

        // Main application logic
        while (true) {
            // Simulate some work in the main application
            try {
                Thread.sleep(1000);
                System.out.println("Main application running...");

                // Add your main application logic here
                // Ejecución Bot Excel
                Machine bot = new Machine();
                bot.startBot();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }

    // Inner class to handle key press listening
    static class KeyPressListener extends Thread {
        private volatile boolean running = true;

        @Override
        public void run() {
            try {
                while (running) {
                    int key = System.in.read();
                    if (key == 'q' || key == 'Q') {  // Change this to your desired key
                        running = false;
                        System.out.println("Key pressed. Stopping the application.");
                        System.exit(0);  // Exit the application
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        public void stopListening() {
            running = false;
        }
    }
}
