
package googlesheetstest;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class Main {

    private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";
    private static Sheets service;

    /**
     * Global instance of the scopes required by this quickstart. If modifying
     * these scopes, delete your previously saved tokens/ folder.
     */
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        
        // PASTE THE ID OF YOUR GOOGLE SHEET HERE ***************************************
        final String spreadsheetId = "1upY5Zq2zbgWb9-BkJSnqi5JukW07cAgipW3OgicM3f0"; 
        // ******************************************************************************   
        
        final String range = "Test!A1";
        service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        List<List<Object>> values = response.getValues();
        if (values == null || values.isEmpty()) {
            System.out.println("No data found.");
        } else {
            System.out.println(values.get(0));
        }
//        UpdateValuesResponse updateResponse = updateValues(spreadsheetId, range);
        AppendValuesResponse appendResponse = appendValues(spreadsheetId, range);
    }
    
    public static UpdateValuesResponse updateValues(String spreadsheetId, String range)
            throws IOException {
        
        List<List<Object>> values = new ArrayList();
        values.add(new ArrayList());
        values.get(0).add("Jackdaws");
        values.get(0).add("love");
        values.get(0).add("my");
        values.get(0).add("big");
        
        ValueRange body = new ValueRange()
                .setValues(values);
        UpdateValuesResponse result =
                service.spreadsheets().values().update(spreadsheetId, range, body)
                        .setValueInputOption("USER_ENTERED")
                        .execute();
        System.out.printf("%d cells updated.", result.getUpdatedCells());

        // [END sheets_update_values]
        return result;
    }

        public static AppendValuesResponse appendValues(String spreadsheetId, String range)
            throws IOException {
        
        List<List<Object>> values = new ArrayList();
        values.add(new ArrayList());
        values.get(0).add("Jackdaws");
        values.get(0).add("love");
        values.get(0).add("my");
        values.get(0).add("big");
        

        
        ValueRange body = new ValueRange()
                .setValues(values);
//        UpdateValuesResponse result =
//                service.spreadsheets().values().update(spreadsheetId, range, body)
//                        .setValueInputOption("USER_ENTERED")
//                        .execute();
        // System.out.printf("%d cells updated.", result.getUpdatedCells());

        AppendValuesResponse result =
                service.spreadsheets().values().append(spreadsheetId, range, body)
                        .setValueInputOption("USER_ENTERED")
                        .execute();
        // [END sheets_update_values]
        return result;
    }
    
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = Main.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        return new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
    }

}
