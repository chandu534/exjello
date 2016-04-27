
```

/*
 * Sample based on Sun Microsystems's JavaMail Example
 * 
 */

import java.text.SimpleDateFormat;
import java.util.GregorianCalendar;
import java.util.Properties;
import java.util.TimeZone;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.Store;

import org.exjello.mail.ExchangeConstants;
import org.exjello.mail.ExchangeStore;

/**
 * You can use this test program in console mode with parameters or with
 * inner static fields
 * 
 * @param argv
 * @throws Exception
 */
public class TestExJello {

	/*
	 * Static field if you whant to test it without console parameters
	 */
	private static String server = "servername";
	// Username sample: DOMAIN\\account:mailbox@mydomain.com
	private static String username = "";
	private static String password = "";
	private static String folder = "Inbox";

	public static final String FORMAT_ISO_8601 = "yyyy-MM-dd'T'HH:mm:ss'Z'";

	public static void main(String argv[]) throws Exception {
		String lastCheck = "";

		/*
		 * A sample "for": if you whant to test FILTER_LAST_CHECK exJello filter
		 */
		for (int i = 0; i < 10; i++) {
			System.out.println(lastCheck);

			CheckMail(argv, lastCheck);

			Thread.currentThread().sleep(3000);

			//
			SimpleDateFormat dateFormatGmt = new SimpleDateFormat(FORMAT_ISO_8601);
			dateFormatGmt.setTimeZone(TimeZone.getTimeZone("GMT"));

			lastCheck = dateFormatGmt.format(new GregorianCalendar().getTime());

			// break;
		}

	}

	private static void CheckMail(String argv[], String laskCheck) {

		if (argv == null || argv.length == 0) {
			argv = new String[] { server, username, password, folder };
		}

		System.out.println(argv);

		if (argv.length != 4) {
			System.out.println("Usage: TestExJello <host> <user> <password> <mbox>");
			// System.exit(1);
		}

		System.out.println("\nTesting exJello\n");

		try {
			Properties props = System.getProperties();

			/*
			 * Use UNFILTERED_PROPERTY if you whant Readed and Unreaded emails
			 */
			props.setProperty(ExchangeConstants.UNFILTERED_PROPERTY, String.valueOf(true));

			/*
			 * You can use other filter implemented on exJello, they are based
			 * on:
			 * http://blogs.msdn.com/b/andrewdelin/archive/2005/05/11/416312.aspx
			 */

			/*
			 * Use FILTER_FROM_PROPERTY to filter email only form certain
			 * address or domain
			 */
			props.setProperty(ExchangeConstants.FILTER_FROM_PROPERTY, "@mydomain.com");

			/*
			 * Use FILTER_TO_PROPERTY to fiter emails sent to specified address
			 * (usefull with alias of mailbox)
			 */
			props.setProperty(ExchangeConstants.FILTER_TO_PROPERTY, "alias@mydomain.com");

			/*
			 * You can use FILTER_LAST_CHECK to get email after certain date
			 */
			if (laskCheck != null && !laskCheck.isEmpty()) {
				props.setProperty(ExchangeConstants.FILTER_LAST_CHECK, laskCheck);
			}

			// Get a Session object
			Session session = Session.getInstance(props, null);

			Store store = new ExchangeStore(session, null);

			// Connect
			store.connect(argv[0], argv[1], argv[2]);

			Folder folder = store.getFolder(argv[3]);

			if (folder == null || !folder.exists()) {
				System.out.println("Invalid folder");
				// System.exit(1);
			}

			// folder.open(Folder.READ_WRITE);
			folder.open(Folder.READ_ONLY);

			Message[] msgs = folder.getMessages();
			System.out.println("Got " + msgs.length + " new messages");

			for (Message msg : msgs) {
				System.out.println(String.format("From: %s", msg.getSubject()));
			}

		} catch (Exception ex) {
			System.err.println(ex.getMessage());
		}

	}
}
```