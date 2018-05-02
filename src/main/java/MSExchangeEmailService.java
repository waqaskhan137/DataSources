import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

import java.net.URI;
import java.util.*;
/**
 * @author Shantanu Sikdar
 */
public class MSExchangeEmailService {
    private static ExchangeService service;
    private static Integer NUMBER_EMAILS_FETCH = 5; // only latest 5 emails/appointments are fetched.
    /**
     * Firstly check, whether "https://webmail.xxxx.com/ews/Services.wsdl" and "https://webmail.xxxx.com/ews/Exchange.asmx"
     * is accessible, if yes that means the Exchange Webservice is enabled on your MS Exchange.
     */
    static {
        try {
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
//service = new ExchangeService(ExchangeVersion.Exchange2007_SP1); //depending on the version of your Exchange. 
            service.setUrl(new URI("https://mail.afiniti.com/ews/Exchange.asmx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    /**
     * Initialize the Exchange Credentials.
     * Don't forget to replace the "USRNAME","PWD","DOMAIN_NAME" variables.
     */
    public MSExchangeEmailService() {
        ExchangeCredentials credentials = new WebCredentials("rana.waqas", "Blackhorse@113", "afiniti.com");
        service.setCredentials(credentials);
    }
    /**
     * Reading one email at a time. Using Item ID of the email.
     * Creating a message data map as a return value.
     */
    public Map readEmailItem(ItemId itemId) {
        Map messageData = new HashMap();
        try {
            Item itm = Item.bind(service, itemId, PropertySet.FirstClassProperties);
            EmailMessage emailMessage = EmailMessage.bind(service, itm.getId());
            messageData.put("emailItemId", emailMessage.getId().toString());
            messageData.put("subject", emailMessage.getSubject().toString());
            messageData.put("fromAddress", emailMessage.getFrom().getAddress().toString());
            messageData.put("senderName", emailMessage.getSender().getName().toString());
            Date dateTimeCreated = emailMessage.getDateTimeCreated();
            messageData.put("SendDate", dateTimeCreated.toString());
            Date dateTimeRecieved = emailMessage.getDateTimeReceived();
            messageData.put("RecievedDate", dateTimeRecieved.toString());
            messageData.put("Size", emailMessage.getSize() + "");
            messageData.put("emailBody", emailMessage.getBody().toString());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return messageData;
    }
/**
 * Number of email we want to read is defined as NUMBER_EMAILS_FETCH, 
 */
    public List readEmails() {
        List msgDataList = new ArrayList ();
        try {
            Folder folder = Folder.bind(service, WellKnownFolderName.Inbox);
            FindItemsResults results = service.findItems(folder.getId(), new ItemView(NUMBER_EMAILS_FETCH));

            int i = 1;
            List mailList = results.getItems();
            for (int k=0; k< mailList.size(); k++){
                Map messageData = new HashMap();
                messageData = readEmailItem((ItemId) mailList.get(k));
                System.out.println("\nEmails #" + (i++) + ":");
                System.out.println("subject : " + messageData.get("subject").toString());
                System.out.println("Sender : " + messageData.get("senderName").toString());
                msgDataList.add(messageData);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return msgDataList;
    }
    /**
     * Reading one appointment at a time. Using Appointment ID of the email.
     * Creating a message data map as a return value.
     */
    public Map readAppointment(Appointment appointment) {
        Map appointmentData = new HashMap();
        try {
            appointmentData.put("appointmentItemId", appointment.getId().toString());
            appointmentData.put("appointmentSubject", appointment.getSubject());
            appointmentData.put("appointmentStartTime", appointment.getStart() + "");
            appointmentData.put("appointmentEndTime", appointment.getEnd() + "");
            //appointmentData.put("appointmentBody", appointment.getBody().toString());
        } catch (ServiceLocalException e) {
            e.printStackTrace();
        }
        return appointmentData;
    }
    /**
     *Number of Appointments we want to read is defined as NUMBER_EMAILS_FETCH,
     *  Here I also considered the start data and end date which is a 30 day span.
     *  We need to set the CalendarView property depending upon the need of ours.   
     */
    public List readAppointments() {
        List apntmtDataList = new ArrayList ();
        Calendar now = Calendar.getInstance();
        Date startDate = Calendar.getInstance().getTime();
        now.add(Calendar.DATE, 30);
        Date endDate = now.getTime();
        try {
            CalendarFolder calendarFolder = CalendarFolder.bind(service, WellKnownFolderName.Calendar, new PropertySet());
            CalendarView cView = new CalendarView(startDate, endDate, 5);
            cView.setPropertySet(new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End));// we can set other properties
            // as well depending upon our need.
            FindItemsResults appointments = calendarFolder.findAppointments(cView);
            int i = 1;
            List appList = appointments.getItems();

            for (int j=0; j<appList.size(); j++){

                System.out.println("\nAPPOINTMENT #" + (i++) + ":");
                Map appointmentData = new HashMap();
                appointmentData = readAppointment((Appointment) appList.get(j));
                System.out.println("subject : " + appointmentData.get("appointmentSubject").toString());
                System.out.println("On : " + appointmentData.get("appointmentStartTime").toString());
                apntmtDataList.add(appointmentData);
            }


        } catch (Exception e) {
            e.printStackTrace();
        }
        return apntmtDataList;
    }
    public void sendEmails(List recipientsList) {
        try {
            StringBuilder strBldr = new StringBuilder();
            strBldr.append("The client submitted the SendAndSaveCopy request at:");
            strBldr.append(Calendar.getInstance().getTime().toString() + " .");
            strBldr.append("Thanks and Regards");
            strBldr.append("Shantanu Sikdar");
            EmailMessage message = new EmailMessage(service);
            message.setSubject("Test sending email");
            message.setBody(new MessageBody(strBldr.toString()));
           for (int i=0; i< recipientsList.size(); i++){
               message.getBccRecipients().add((String) recipientsList.get(i));
           }
            message.sendAndSaveCopy();
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("message sent");
    }
    public static void main(String[] args) {
        MSExchangeEmailService msees = new MSExchangeEmailService();
        msees.readEmails();
        msees.readAppointments();
        List recipientsList = new ArrayList<>();
        recipientsList.add("waqaskhan137@gmail.com");
//        recipientsList.add("email.id2@domain1.com");
//        recipientsList.add("email.id3@domain2.com");
        msees.sendEmails(recipientsList);
    }
}