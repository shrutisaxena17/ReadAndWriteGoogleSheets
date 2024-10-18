package org.example;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;

import java.io.FileReader;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class GoogleSheetsExample {

    private static final String APPLICATION_NAME = "Google Sheets Demo";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "token";
    private static final String CREDENTIALS_FILE_PATH = "src/main/resources/credentials.json";
    private static final String SPREADSHEET_ID = "146R8XeveHYsDPRSLPVRd65tGHMogjhT1qRhqG5IGGGA";

    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new FileReader(CREDENTIALS_FILE_PATH));

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, Collections.singletonList(SheetsScopes.SPREADSHEETS))
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();

        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8881).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    // Method to print data from all sheets in the spreadsheet
    public void printAllSheetsData() throws IOException, GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();

        Spreadsheet spreadsheet = service.spreadsheets().get(SPREADSHEET_ID).execute();
        List<Sheet> sheets = spreadsheet.getSheets();

        for (Sheet sheet : sheets) {
            String sheetName = sheet.getProperties().getTitle();
            System.out.println("Data from sheet: " + sheetName);

            ValueRange response = service.spreadsheets().values()
                    .get(SPREADSHEET_ID, sheetName).execute();
            List<List<Object>> values = response.getValues();

            if (values == null || values.isEmpty()) {
                System.out.println("No data found in this sheet.");
            } else {
                for (List<Object> row : values) {
                    System.out.println(row);
                }
            }
            System.out.println();
        }
    }

    public void writeIntoSheet() throws IOException, GeneralSecurityException {
        List<List<Object>> data = Arrays.asList(
                Arrays.asList("Deepti", "Thapar"),
                Arrays.asList("Shalini", "Singh")
        );

        String range="Sheet1!A4:B5";

        ValueRange body = new ValueRange().setValues(data);

        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();

        UpdateValuesResponse result = service.spreadsheets().values()
                .update(SPREADSHEET_ID, range, body)
                .setValueInputOption("RAW")
                .execute();

    }

    // Method to write data using multiple threads repeatedly with a 5-second delay
    public void writeDataWithThreads() throws IOException, GeneralSecurityException, InterruptedException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();

        // Infinite loop to keep running the threads every 5 seconds
        while (true) {
            // Create three threads, each writing different data
            Thread thread1 = new Thread(new WriteDataTask(service, "John", "john@example.com", "1234567890"));
            Thread thread2 = new Thread(new WriteDataTask(service, "Alice", "alice@example.com", "9876543210"));
            Thread thread3 = new Thread(new WriteDataTask(service, "Bob", "bob@example.com", "4567890123"));

            // Start and synchronize the threads to run one after another
            thread1.start();
            thread1.join(); // Ensure thread1 finishes before starting the next
            thread2.start();
            thread2.join(); // Ensure thread2 finishes before starting the next
            thread3.start();
            thread3.join(); // Ensure thread3 finishes before the next cycle

            // Wait for 5 seconds before running the threads again
            Thread.sleep(5000);
        }
    }


    // Runnable task to handle writing data to Google Sheets
    static class WriteDataTask implements Runnable {
        private final Sheets service;
        private final String name;
        private final String email;
        private final String phone;

        public WriteDataTask(Sheets service, String name, String email, String phone) {
            this.service = service;
            this.name = name;
            this.email = email;
            this.phone = phone;
        }

        @Override
        public void run() {
            try {
                writeDataToSheet(service, name, email, phone);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        private void writeDataToSheet(Sheets service, String name, String email, String phone) throws IOException {
            // Determine the next available row by checking how many rows are filled in column A
            String range = "Sheet1!A:A"; // Column A is checked for filled rows
            ValueRange response = service.spreadsheets().values().get(SPREADSHEET_ID, range).execute();
            List<List<Object>> values = response.getValues();
            int nextRow = (values != null ? values.size() + 1 : 1);

            // Define the new range to write data
            String newRange = "Sheet1!A" + nextRow + ":C" + nextRow;
            List<List<Object>> data = Arrays.asList(
                    Arrays.asList(name, email, phone)
            );

            ValueRange body = new ValueRange().setValues(data);
            service.spreadsheets().values()
                    .update(SPREADSHEET_ID, newRange, body)
                    .setValueInputOption("RAW")
                    .execute();

            System.out.println("Data written to row " + nextRow);
        }
    }

    public static void main(String[] args) {
        GoogleSheetsExample example = new GoogleSheetsExample();
        try {
            //example.printAllSheetsData();
            //example.writeIntoSheet();
            example.writeDataWithThreads();
        } catch (IOException | GeneralSecurityException e) {
            e.printStackTrace();
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
    }
}
