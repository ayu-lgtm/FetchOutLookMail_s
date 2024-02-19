package com.jbdl.fetchoutlookmail.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.*;
import javax.mail.Folder;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.ZoneId;
import java.util.Date;
import java.util.HashSet;
import java.util.Properties;
import java.io.*;
import java.util.Set;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;


import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimeBodyPart;

@RestController
@EnableScheduling
public class OutlookController {
    private final Set<String> processedDateTimeIds = new HashSet<>();

    @GetMapping("/fetchEmailsFromGroup")
    public String fetchEmailsFromGroup() {
        // Connect to Outlook and fetch emails
        fetchEmails();
        return "Emails fetched from the Outlook group.";
    }

    // Component for scheduled task to fetch emails periodically
    @Component
    public class EmailFetchingTask {
        @Scheduled(fixedRate = 40000) // Run every hour (adjust interval as needed)
        public void fetchEmails() {
            // Call your method to fetch emails from Outlook group
            // This method will be automatically invoked by Spring at the specified interval
            fetchEmailsFromGroup();
        }
    }

    public void fetchEmails() {
        // Outlook email configuration
        String host = "outlook.office365.com";
        String username = "rastogiayush143@outlook.com";
        String password = "Ayush@123";

        // Properties for connecting to Outlook server
        Properties properties = new Properties();
        properties.put("mail.imap.host", host);
        properties.put("mail.imap.port", "993");
        properties.put("mail.imap.starttls.enable", "true");
        properties.put("mail.debug", "true");
        properties.put("mail.imap.ssl.enable", "true");

        // Get session
        Session session = Session.getInstance(properties);

        FileOutputStream fileOut = null;
        Workbook workbook = null;

        try {
            // Connect to the store
            Store store = session.getStore("imap");
            store.connect(host, username, password);
            System.out.println("Connected to the mail server.");

            // Get inbox folder
            Folder inbox = store.getFolder("Inbox");
            if (inbox == null || !inbox.exists()) {
                System.out.println("Inbox folder not found.");
                return;
            }
            inbox.open(Folder.READ_ONLY);
            System.out.println("Opened Inbox folder.");

            // Fetch messages
            Message[] messages = inbox.getMessages();
            System.out.println("Fetched " + messages.length + " messages from Inbox.");

            // Check if file exists
            String desktopPath = System.getProperty("user.home") + "\\Desktop";
            String filePath = Paths.get(desktopPath, "emails.xlsx").toString();
            boolean fileExists = new File(filePath).exists();

            if (fileExists) {
                // If file exists, open it
                FileInputStream fis = new FileInputStream(filePath);
                workbook = new XSSFWorkbook(fis);
            } else {
                // If file doesn't exist, create new workbook
                workbook = new XSSFWorkbook();
                // Create header row
                Sheet sheet = workbook.createSheet("Emails");
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Sender");
                headerRow.createCell(1).setCellValue("Date");
                headerRow.createCell(2).setCellValue("Time");
                headerRow.createCell(3).setCellValue("Receiver");
                headerRow.createCell(4).setCellValue("Subject");
                headerRow.createCell(5).setCellValue("Description");
            }

            // Get the sheet
            Sheet sheet = workbook.getSheet("Emails");
            int lastRowNum = sheet.getLastRowNum();

            // Process the fetched messages
            for (Message message : messages) {

                String dateTimeId = getDateTimeId(message);

                // Check if this message has already been processed
                if (processedDateTimeIds.contains(dateTimeId)) {
                    System.out.println("Skipping duplicate message: " + dateTimeId);
                    continue; // Skip processing this message
                }

                processedDateTimeIds.add(dateTimeId);

                Row row = sheet.createRow(++lastRowNum);

                Address[] fromAddresses = message.getFrom();
                String senderEmail = (fromAddresses != null && fromAddresses.length > 0) ? fromAddresses[0].toString() : "Unknown";
                String receiverEmail = username; // Assuming emails are fetched from the inbox of the configured user
                String subject = message.getSubject().trim();
                String description = getTextFromMessage(message); // Get email body as description
                String description_text = extractTextFromHtml(description);


                // Truncate description if it exceeds the maximum length
                if (description_text.length() > 32767) {
                    description_text = description_text.substring(0, 32767);
                }

                row.createCell(0).setCellValue(senderEmail);
                row.createCell(1).setCellValue(message.getSentDate().toInstant().atZone(ZoneId.systemDefault()).toLocalDate().toString());
                row.createCell(2).setCellValue(message.getSentDate().toInstant().atZone(ZoneId.systemDefault()).toLocalTime().toString());
                row.createCell(3).setCellValue(receiverEmail);
                row.createCell(4).setCellValue(subject);
                row.createCell(5).setCellValue(description_text);
            }

            // Write the workbook to the file
            fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
            System.out.println("Emails saved to emails.xlsx");

        } catch (MessagingException | IOException e) {
            e.printStackTrace();
            System.out.println("Exception: " + e.getMessage());
        } finally {
            try {
                // Close resources
                if (fileOut != null) {
                    fileOut.close();
                }
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private String getTextFromMessage(Message message) throws MessagingException, IOException {
        Object content = message.getContent();
        StringBuilder text = new StringBuilder();

        if (content instanceof String) {
            // If the content is plain text, append it to the StringBuilder
            text.append((String) content);
        }
        if (content instanceof MimeMultipart) {
            // If the content is multipart (HTML email), parse it and extract text
            MimeMultipart multipart = (MimeMultipart) content;
            for (int i = 0; i < multipart.getCount(); i++) {
                MimeBodyPart bodyPart = (MimeBodyPart) multipart.getBodyPart(i);
                if (bodyPart.isMimeType("text/plain")) {
                    // If the part is plain text, append it to the StringBuilder
                    text.append(bodyPart.getContent());
                }
            }
        }

        // Clean the text by removing extra whitespace
        return cleanText(text.toString());
    }

    private String cleanText(String text) {
        // Remove extra whitespace
        return text.replaceAll("\\s+", " ").trim();
    }

    private String extractTextFromHtml(String htmlContent) {
        // Parse HTML content using Jsoup
        Document doc = Jsoup.parse(htmlContent);

        // Extract text from the parsed document
        String textContent = doc.text();

        // Clean the text by removing extra whitespace
        return cleanText(textContent);
    }

    private String getDateTimeId(Message message) throws MessagingException {
        // Get the sent date and time of the message
        Date sentDate = message.getSentDate();
        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        DateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
        String dateString = dateFormat.format(sentDate);
        String timeString = timeFormat.format(sentDate);

        // Concatenate date and time strings with a separator (e.g., underscore)
        return dateString + "_" + timeString;
    }


}
