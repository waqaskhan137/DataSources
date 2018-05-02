import microsoft.exchange.webservices.data.*;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

import java.net.URI;
import java.util.List;

public class Main {
    public static void main(String [] args) throws Exception {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
        ExchangeCredentials credentials = new WebCredentials("rana.waqas@afiniti.com","Blackhorse@113");
        service.setCredentials(credentials);
        service.setUrl(new URI("https://mail.afiniti.com/ews/exchange.asmx"));

        ItemView view = new ItemView (10);
        FindItemsResults findResults = service.findItems(WellKnownFolderName.Inbox, view);
        List itemList = findResults.getItems();

        for (int i=0; i<itemList.size(); i++){
            itemList.get(i).load(new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.MimeContent));
            System.out.println("id==========" + item.getId());
            System.out.println("sub==========" + item.getSubject());
            System.out.println("sub==========" + item.getMimeContent());
        }
        for(Item item : (Item) findResults.getItems()){

        }
    }
}
