package org.rpa;

import java.io.IOException;

public class KeyPress {
    public static void main(String[] args) {
        // Start the key press listener thread
        KeyPressListener keyPressListener = new KeyPressListener();
        keyPressListener.start();

        // Main application logic
        while (true) {
            // Simulate some work in the main application
            try {
                Thread.sleep(1000);
                System.out.println("Main application running...");
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            // Add your main application logic here
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
