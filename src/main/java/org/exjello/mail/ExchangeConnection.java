/*
Copyright (c) 2010 Eric Glass, Mirco Attocchi

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
 */

package org.exjello.mail;

import static org.exjello.mail.ExchangeConstants.CONNECTION_TIMEOUT_PROPERTY;
import static org.exjello.mail.ExchangeConstants.DELETE_PROPERTY;
import static org.exjello.mail.ExchangeConstants.FROM_PROPERTY;
import static org.exjello.mail.ExchangeConstants.LIMIT_PROPERTY;
import static org.exjello.mail.ExchangeConstants.LOCAL_ADDRESS_PROPERTY;
import static org.exjello.mail.ExchangeConstants.MAILBOX_PROPERTY;
import static org.exjello.mail.ExchangeConstants.PORT_PROPERTY;
import static org.exjello.mail.ExchangeConstants.SSL_PROPERTY;
import static org.exjello.mail.ExchangeConstants.TIMEOUT_PROPERTY;
import static org.exjello.mail.ExchangeConstants.UNFILTERED_PROPERTY;

import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.math.BigInteger;
import java.net.InetAddress;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.UnknownHostException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Random;

import javax.mail.Address;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.httpclient.HttpClient;
import org.apache.commons.httpclient.NTCredentials;
import org.apache.commons.httpclient.UsernamePasswordCredentials;
import org.apache.commons.httpclient.auth.AuthScope;
import org.apache.commons.httpclient.methods.ByteArrayRequestEntity;
import org.apache.commons.httpclient.methods.GetMethod;
import org.apache.commons.httpclient.methods.InputStreamRequestEntity;
import org.apache.commons.httpclient.methods.OptionsMethod;
import org.apache.commons.httpclient.methods.PostMethod;
import org.apache.commons.httpclient.methods.PutMethod;
import org.apache.commons.httpclient.methods.RequestEntity;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

class ExchangeConnection {

	private static final Map<String, byte[]> RESOURCES = new HashMap<String, byte[]>();

	private static final Random RANDOM = new Random();

	private static final String GET_UNREAD_MESSAGES_SQL_RESOURCE = "get-unread-messages.sql";

	private static final String GET_ALL_MESSAGES_SQL_RESOURCE = "get-all-messages.sql";

	private static final String GET_FILTERED_MESSAGES_SQL_RESOURCE = "get-filtered-messages.sql";

	private static final String BOOKMARK_FILTER_UNREADED = "{BOOKMARK_FILTER_UNREADED}";

	private static final String BOOKMARK_FILTER_TO = "{BOOKMARK_FILTER_TO}";

	private static final String BOOKMARK_FILTER_FROM = "{BOOKMARK_FILTER_FROM}";

	private static final String BOOKMARK_FILTER_NOT_FROM = "{BOOKMARK_FILTER_NOT_FROM}";

	private static final String BOOKMARK_FILTER_LAST_CHECK = "{BOOKMARK_FILTER_LAST_CHECK}";

	private static final String SIGN_ON_URI = "/exchweb/bin/auth/owaauth.dll";

	private static final String DEBUG_PASSWORD_PROPERTY = "org.exjello.mail.debug.password";

	private static final String HTTPMAIL_NAMESPACE = "urn:schemas:httpmail:";

	private static final String MAILHEADER_NAMESPACE = "urn:schemas:mailheader:";

	private static final String DAV_NAMESPACE = "DAV:";

	private static final String PROPFIND_METHOD = "PROPFIND";

	private static final String SEARCH_METHOD = "SEARCH";

	private static final String BDELETE_METHOD = "BDELETE";

	private static final String BPROPPATCH_METHOD = "BPROPPATCH";

	private static final String PROPPATCH_METHOD = "PROPPATCH";

	private static final String MOVE_METHOD = "MOVE";

	private static final String MESSAGE_CONTENT_TYPE = "message/rfc822";

	private static final String XML_CONTENT_TYPE = "text/xml; charset=\"UTF-8\"";

    private static final String FORM_URLENCODED_CONTENT_TYPE =
            "application/x-www-form-urlencoded; charset=utf-8";

	private static final int HTTP_PORT = 80;

	private static final int HTTPS_PORT = 443;

	private static final boolean[] ALLOWED_CHARS = new boolean[128];

	private static final char[] HEXABET = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F' };

	private static byte[] findInboxEntity;

	private static byte[] unreadInboxEntity;

	private static byte[] allInboxEntity;

	private static byte[] customInboxEntity;

	private final Session session;

	private final String server;

	private final String mailbox;

	private final String username;

	private final String password;

	private final int timeout;

	private final int connectionTimeout;

	private final InetAddress localAddress;

	private final boolean unfiltered;

	/* Mirco */
	private final String filterLastCheck;
	private final String filterFrom;
	private final String filterNotFrom;
	private final String filterTo;

	private final boolean delete;

	private final int limit;

	private HttpClient client;

	private String inbox;

	private String drafts;

	private String submissionUri;

	private String sentitems;

	private String outbox;

	static {
		// a - z
		for (int i = 97; i < 123; i++)
			ALLOWED_CHARS[i] = true;
		// A - Z
		for (int i = 64; i < 91; i++)
			ALLOWED_CHARS[i] = true;
		// 0 - 9
		for (int i = 48; i < 58; i++)
			ALLOWED_CHARS[i] = true;
		ALLOWED_CHARS[(int) '-'] = true;
		ALLOWED_CHARS[(int) '_'] = true;
		ALLOWED_CHARS[(int) '.'] = true;
		ALLOWED_CHARS[(int) '!'] = true;
		ALLOWED_CHARS[(int) '~'] = true;
		ALLOWED_CHARS[(int) '*'] = true;
		ALLOWED_CHARS[(int) '\''] = true;
		ALLOWED_CHARS[(int) '('] = true;
		ALLOWED_CHARS[(int) ')'] = true;
		ALLOWED_CHARS[(int) '%'] = true;
		ALLOWED_CHARS[(int) ':'] = true;
		ALLOWED_CHARS[(int) '@'] = true;
		ALLOWED_CHARS[(int) '&'] = true;
		ALLOWED_CHARS[(int) '='] = true;
		ALLOWED_CHARS[(int) '+'] = true;
		ALLOWED_CHARS[(int) '$'] = true;
		ALLOWED_CHARS[(int) ','] = true;
		ALLOWED_CHARS[(int) ';'] = true;
		ALLOWED_CHARS[(int) '/'] = true;
	}

	public static ExchangeConnection createConnection(String protocol, Session session, String host, int port, String username, String password) throws Exception {
		String prefix = "mail." + protocol.toLowerCase() + ".";
		boolean debugPassword = Boolean.parseBoolean(session.getProperty(DEBUG_PASSWORD_PROPERTY));
		String pwd = (password == null) ? null : debugPassword ? password : "<password>";
		if (host == null || username == null || password == null) {
			if (session.getDebug()) {
				session.getDebugOut().println("Missing parameter; host=\"" + host + "\",username=\"" + username + "\",password=\"" + pwd + "\"");
			}
			throw new IllegalStateException("Host, username, and password must be specified.");
		}
		boolean unfiltered = Boolean.parseBoolean(session.getProperty(UNFILTERED_PROPERTY));
		/* Mirco */
		String filterLastCheck = session.getProperty(ExchangeConstants.FILTER_LAST_CHECK);
		String filterFrom = session.getProperty(ExchangeConstants.FILTER_FROM_PROPERTY);
		String filterNotFrom = session.getProperty(ExchangeConstants.FILTER_NOT_FROM_PROPERTY);
		String filterTo = session.getProperty(ExchangeConstants.FILTER_TO_PROPERTY);
		boolean delete = Boolean.parseBoolean(session.getProperty(DELETE_PROPERTY));
		boolean secure = Boolean.parseBoolean(session.getProperty(prefix + SSL_PROPERTY));
		int limit = -1;
		String limitString = session.getProperty(LIMIT_PROPERTY);
		if (limitString != null) {
			try {
				limit = Integer.parseInt(limitString);
			} catch (NumberFormatException ex) {
				throw new NumberFormatException("Invalid limit specified: " + limitString);
			}
		}
		try {
			URL url = new URL(host);
			// if parsing succeeded, then strip out the components and use
			secure = "https".equalsIgnoreCase(url.getProtocol());
			host = url.getHost();
			int specifiedPort = url.getPort();
			if (specifiedPort != -1)
				port = specifiedPort;
		} catch (MalformedURLException ex) {
			if (session.getDebug()) {
				session.getDebugOut().println("Not parsing " + host + " as a URL; using explicit options for " + "secure, host, and port.");
			}
		}
		if (port == -1) {
			try {
				port = Integer.parseInt(session.getProperty(prefix + PORT_PROPERTY));
			} catch (Exception ignore) {
			}
			if (port == -1)
				port = secure ? HTTPS_PORT : HTTP_PORT;
		}
		String server = (secure ? "https://" : "http://") + host;
		if (secure ? (port != HTTPS_PORT) : (port != HTTP_PORT)) {
			server += ":" + port;
		}
		String mailbox = session.getProperty(MAILBOX_PROPERTY);
		if (mailbox == null) {
			mailbox = session.getProperty(prefix + FROM_PROPERTY);
			if (mailbox == null) {
				mailbox = InternetAddress.getLocalAddress(session).getAddress();
			}
		}

		int index = username.indexOf(':');
		if (index != -1) {
			mailbox = username.substring(index + 1);
			username = username.substring(0, index);
			String mailboxOptions = null;
			index = mailbox.indexOf('[');
			if (index != -1) {
				mailboxOptions = mailbox.substring(index + 1);
				mailboxOptions = mailboxOptions.substring(0, mailboxOptions.indexOf(']'));
				mailbox = mailbox.substring(0, index);
			}
			if (mailboxOptions != null) {
				Properties props = null;
				try {
					props = parseOptions(mailboxOptions);
				} catch (Exception ex) {
					throw new IllegalArgumentException("Unable to parse mailbox options: " + ex.getMessage(), ex);
				}
				String value = props.getProperty("unfiltered");
				if (value != null)
					unfiltered = Boolean.parseBoolean(value);

				/* Mirco */
				value = props.getProperty("filterLastCheck");
				if (value != null)
					filterLastCheck = value;
				value = props.getProperty("filterTo");
				if (value != null)
					filterTo = value;
				value = props.getProperty("filterFrom");
				if (value != null)
					filterFrom = value;
				value = props.getProperty("filterNotFrom");
				if (value != null)
					filterNotFrom = value;

				value = props.getProperty("delete");
				if (value != null)
					delete = Boolean.parseBoolean(value);
				value = props.getProperty("limit");
				if (value != null) {
					try {
						limit = Integer.parseInt(value);
					} catch (NumberFormatException ex) {
						throw new NumberFormatException("Invalid limit specified: " + value);
					}
				}
			} else if (session.getDebug()) {
				session.getDebugOut().println("No mailbox options specified; " + "using explicit limit, unfiltered, and delete.");
			}
		} else if (session.getDebug()) {
			session.getDebugOut().println("No mailbox specified in username; " + "using explicit mailbox, limit, unfiltered, and delete.");
		}
		int timeout = -1;
		String timeoutString = session.getProperty(prefix + TIMEOUT_PROPERTY);
		if (timeoutString != null) {
			try {
				timeout = Integer.parseInt(timeoutString);
			} catch (NumberFormatException ex) {
				throw new NumberFormatException("Invalid timeout value: " + timeoutString);
			}
		}
		int connectionTimeout = -1;
		timeoutString = session.getProperty(prefix + CONNECTION_TIMEOUT_PROPERTY);
		if (timeoutString != null) {
			try {
				connectionTimeout = Integer.parseInt(timeoutString);
			} catch (NumberFormatException ex) {
				throw new NumberFormatException("Invalid connection timeout value: " + timeoutString);
			}
		}
		InetAddress localAddress = null;
		String localAddressString = session.getProperty(prefix + LOCAL_ADDRESS_PROPERTY);
		if (localAddressString != null) {
			try {
				localAddress = InetAddress.getByName(localAddressString);
			} catch (Exception ex) {
				throw new UnknownHostException("Invalid local address specified: " + localAddressString);
			}
		}
		if (mailbox == null) {
			throw new IllegalStateException("No mailbox specified.");
		}
		if (session.getDebug()) {
			PrintStream debugStream = session.getDebugOut();
			debugStream.println("Server:\t" + server);
			debugStream.println("Username:\t" + username);
			debugStream.println("Password:\t" + pwd);
			debugStream.println("Mailbox:\t" + mailbox);
			debugStream.print("Options:\t");
			debugStream.print((limit > 0) ? "Message Limit = " + limit : "Unlimited Messages");
			debugStream.print(unfiltered ? "; Unfiltered" : "; Filtered to Unread");
			debugStream.print(filterLastCheck == null || "".equals(filterLastCheck) ? "; NO filterLastCheck" : "; Filtered after " + filterLastCheck);
			debugStream.print(filterFrom == null || "".equals(filterFrom) ? "; NO filterFromDomain" : "; Filtered from " + filterFrom);
			debugStream.print(filterNotFrom == null || "".equals(filterNotFrom) ? "; NO filterNotFrom" : "; Filtered not from " + filterNotFrom);
			debugStream.print(filterTo == null || "".equals(filterTo) ? "; NO filterToEmail" : "; Filtered to " + filterTo);
			debugStream.println(delete ? "; Delete Messages on Delete" : "; Mark as Read on Delete");
			if (timeout > 0) {
				debugStream.println("Read timeout:\t" + timeout + " ms");
			}
			if (connectionTimeout > 0) {
				debugStream.println("Connection timeout:\t" + connectionTimeout + " ms");
			}
		}
		return new ExchangeConnection(session, server, mailbox, username, password, timeout, connectionTimeout, localAddress, unfiltered, delete, limit, filterLastCheck, filterFrom, filterNotFrom, filterTo);
	}

	private ExchangeConnection(Session session, String server, String mailbox, String username, String password, int timeout, int connectionTimeout, InetAddress localAddress, boolean unfiltered, boolean delete, int limit, String filterLastCheck, String filterFrom, String filterNotFrom, String filterTo) {
		this.session = session;
		this.server = server;
		this.mailbox = mailbox;
		this.username = username;
		this.password = password;
		this.timeout = timeout;
		this.connectionTimeout = connectionTimeout;
		this.localAddress = localAddress;
		this.unfiltered = unfiltered;
		this.delete = delete;
		this.limit = limit;
		/* Mirco */
		this.filterLastCheck = filterLastCheck;
		this.filterFrom = filterFrom;
		this.filterNotFrom = filterNotFrom;
		this.filterTo = filterTo;
	}

	public void connect() throws Exception {
		synchronized (this) {
			inbox = null;
			drafts = null;
			submissionUri = null;
			sentitems = null;
			outbox = null;
			try {
				signOn();
			} catch (Exception ex) {
				inbox = null;
				drafts = null;
				submissionUri = null;
				sentitems = null;
				outbox = null;
				throw ex;
			}
		}
	}

	public List<String> getMessages(String name) throws Exception {
		final List<String> messages = new ArrayList<String>();

		/* by default we list inbox */
		String currentFolder = inbox;
		if (name.equalsIgnoreCase(ExchangeFolder.INBOX)) {
			currentFolder = inbox;
		} else if (name.equalsIgnoreCase(ExchangeFolder.SENTITEMS)) {
			currentFolder = sentitems;
		} else if (name.equalsIgnoreCase(ExchangeFolder.OUTBOX)) {
			currentFolder = outbox;
		} else if (name.equalsIgnoreCase(ExchangeFolder.DRAFT)) {
			currentFolder = drafts;
		}

		listFolder(new DefaultHandler() {
			private final StringBuilder content = new StringBuilder();

			public void characters(char[] ch, int start, int length) throws SAXException {
				content.append(ch, start, length);
			}

			public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
				content.setLength(0);
			}

			public void endElement(String uri, String localName, String qName) throws SAXException {
				if (!DAV_NAMESPACE.equals(uri))
					return;
				if (!"href".equals(localName))
					return;
				messages.add(content.toString());
			}
		}, currentFolder);
		return Collections.unmodifiableList(messages);
	}

	public void send(MimeMessage message) throws Exception {
		Address[] bccRecipients = message.getRecipients(Message.RecipientType.BCC);
		if (bccRecipients == null || bccRecipients.length == 0) {
			bccRecipients = null;
		}
		message.setRecipients(Message.RecipientType.BCC, (Address[]) null);
		synchronized (this) {
			if (!isConnected()) {
				throw new IllegalStateException("Not connected.");
			}
			if (!canSend()) {
				throw new IllegalStateException("Unable to access outbox.");
			}
			HttpClient client = getClient();
			String path = drafts;
			if (!path.endsWith("/"))
				path += "/";
			String messageName = generateMessageName();
			path += escape(messageName + ".eml");
			PutMethod op = new PutMethod(path);
			op.setRequestHeader("Content-Type", MESSAGE_CONTENT_TYPE);
			op.setRequestEntity(createMessageEntity(message));
			InputStream stream = null;
			try {
				int status = client.executeMethod(op);
				stream = op.getResponseBodyAsStream();
				if (status >= 300) {
					throw new IllegalStateException("Unable to post message to draft folder.");
				}
			} finally {
				try {
					if (stream != null) {
						byte[] buf = new byte[65536];
						try {
							if (session.getDebug()) {
								PrintStream log = session.getDebugOut();
								log.println("Response Body:");
								int count;
								while ((count = stream.read(buf, 0, 65536)) != -1) {
									log.write(buf, 0, count);
								}
								log.flush();
								log.println();
							} else {
								while (stream.read(buf, 0, 65536) != -1)
									;
							}
						} catch (Exception ignore) {
						} finally {
							try {
								stream.close();
							} catch (Exception ignore2) {
							}
						}
					}
				} finally {
					op.releaseConnection();
				}
			}
			if (bccRecipients != null) {
				ExchangeMethod patch = new ExchangeMethod(PROPPATCH_METHOD, path);
				patch.setHeader("Content-Type", XML_CONTENT_TYPE);
				patch.addHeader("Depth", "0");
				patch.addHeader("Translate", "f");
				patch.addHeader("Brief", "t");
				patch.setRequestEntity(createAddBccEntity(bccRecipients));
				stream = null;
				try {
					int status = client.executeMethod(patch);
					stream = patch.getResponseBodyAsStream();
					if (status >= 300) {
						throw new IllegalStateException("Unable to add BCC recipients. Status: " + status);
					}
				} finally {
					try {
						if (stream != null) {
							byte[] buf = new byte[65536];
							try {
								if (session.getDebug()) {
									PrintStream log = session.getDebugOut();
									log.println("Response Body:");
									int count;
									while ((count = stream.read(buf, 0, 65536)) != -1) {
										log.write(buf, 0, count);
									}
									log.flush();
									log.println();
								} else {
									while (stream.read(buf, 0, 65536) != -1)
										;
								}
							} catch (Exception ignore) {
							} finally {
								try {
									stream.close();
								} catch (Exception ignore2) {
								}
							}
						}
					} finally {
						patch.releaseConnection();
					}
				}
			}
			ExchangeMethod move = new ExchangeMethod(MOVE_METHOD, path);
			String destination = submissionUri;
			if (!destination.endsWith("/"))
				destination += "/";
			move.setHeader("Destination", destination);
			stream = null;
			try {
				int status = client.executeMethod(move);
				stream = move.getResponseBodyAsStream();
				if (status >= 300) {
					throw new IllegalStateException("Unable to move message to outbox: Status " + status);
				}
			} finally {
				try {
					if (stream != null) {
						byte[] buf = new byte[65536];
						try {
							if (session.getDebug()) {
								PrintStream log = session.getDebugOut();
								log.println("Response Body:");
								int count;
								while ((count = stream.read(buf, 0, 65536)) != -1) {
									log.write(buf, 0, count);
								}
								log.flush();
								log.println();
							} else {
								while (stream.read(buf, 0, 65536) != -1)
									;
							}
						} catch (Exception ignore) {
						} finally {
							try {
								stream.close();
							} catch (Exception ignore2) {
							}
						}
					}
				} finally {
					move.releaseConnection();
				}
			}
			if (session.getDebug()) {
				session.getDebugOut().println("Sent successfully.");
			}
		}
	}

	public void delete(List<ExchangeMessage> messages) throws Exception {
		if (delete) {
			doDelete(messages);
		} else {
			doMarkRead(messages);
		}
	}

	public InputStream getInputStream(ExchangeMessage message) throws Exception {
		synchronized (this) {
			if (!isConnected()) {
				throw new IllegalStateException("Not connected.");
			}
			HttpClient client = getClient();
			GetMethod op = new GetMethod(escape(message.getUrl()));
			op.setRequestHeader("Translate", "F");
			InputStream stream = null;
			try {
				int status = client.executeMethod(op);
				stream = op.getResponseBodyAsStream();
				if (status >= 300) {
					throw new IllegalStateException("Unable to obtain inbox: " + status);
				}
				final File tempFile = File.createTempFile("exmail", null, null);
				tempFile.deleteOnExit();
				OutputStream output = new FileOutputStream(tempFile);
				byte[] buf = new byte[65536];
				int count;
				while ((count = stream.read(buf, 0, 65536)) != -1) {
					output.write(buf, 0, count);
				}
				output.flush();
				output.close();
				stream.close();
				stream = null;
				return new CachedMessageStream(tempFile, (ExchangeFolder) message.getFolder());
			} finally {
				try {
					if (stream != null) {
						byte[] buf = new byte[65536];
						try {
							if (session.getDebug()) {
								PrintStream log = session.getDebugOut();
								log.println("Response Body:");
								int count;
								while ((count = stream.read(buf, 0, 65536)) != -1) {
									log.write(buf, 0, count);
								}
								log.flush();
								log.println();
							} else {
								while (stream.read(buf, 0, 65536) != -1)
									;
							}
						} catch (Exception ignore) {
						} finally {
							try {
								stream.close();
							} catch (Exception ignore2) {
							}
						}
					}
				} finally {
					op.releaseConnection();
				}
			}
		}
	}

	private void doDelete(List<ExchangeMessage> messages) throws Exception {
		synchronized (this) {
			if (!isConnected()) {
				throw new IllegalStateException("Not connected.");
			}
			HttpClient client = getClient();
			String path = inbox;
			if (!path.endsWith("/"))
				path += "/";
			ExchangeMethod op = new ExchangeMethod(BDELETE_METHOD, path);
			op.setHeader("Content-Type", XML_CONTENT_TYPE);
			op.addHeader("If-Match", "*");
			op.addHeader("Brief", "t");
			op.setRequestEntity(createDeleteEntity(messages));
			InputStream stream = null;
			try {
				int status = client.executeMethod(op);
				stream = op.getResponseBodyAsStream();
				if (status >= 300) {
					throw new IllegalStateException("Unable to delete messages.");
				}
			} finally {
				try {
					if (stream != null) {
						byte[] buf = new byte[65536];
						try {
							if (session.getDebug()) {
								PrintStream log = session.getDebugOut();
								log.println("Response Body:");
								int count;
								while ((count = stream.read(buf, 0, 65536)) != -1) {
									log.write(buf, 0, count);
								}
								log.flush();
								log.println();
							} else {
								while (stream.read(buf, 0, 65536) != -1)
									;
							}
						} catch (Exception ignore) {
						} finally {
							try {
								stream.close();
							} catch (Exception ignore2) {
							}
						}
					}
				} finally {
					op.releaseConnection();
				}
			}
		}
	}

	private void doMarkRead(List<ExchangeMessage> messages) throws Exception {
		synchronized (this) {
			if (!isConnected()) {
				throw new IllegalStateException("Not connected.");
			}
			HttpClient client = getClient();
			String path = inbox;
			if (!path.endsWith("/"))
				path += "/";
			ExchangeMethod op = new ExchangeMethod(BPROPPATCH_METHOD, path);
			op.setHeader("Content-Type", XML_CONTENT_TYPE);
			op.addHeader("If-Match", "*");
			op.addHeader("Brief", "t");
			op.setRequestEntity(createMarkReadEntity(messages));
			InputStream stream = null;
			try {
				int status = client.executeMethod(op);
				stream = op.getResponseBodyAsStream();
				if (status >= 300) {
					throw new IllegalStateException("Unable to mark messages read.");
				}
			} finally {
				try {
					if (stream != null) {
						byte[] buf = new byte[65536];
						try {
							if (session.getDebug()) {
								PrintStream log = session.getDebugOut();
								log.println("Response Body:");
								int count;
								while ((count = stream.read(buf, 0, 65536)) != -1) {
									log.write(buf, 0, count);
								}
								log.flush();
								log.println();
							} else {
								while (stream.read(buf, 0, 65536) != -1)
									;
							}
						} catch (Exception ignore) {
						} finally {
							try {
								stream.close();
							} catch (Exception ignore2) {
							}
						}
					}
				} finally {
					op.releaseConnection();
				}
			}
		}
	}

	private boolean isConnected() {
		return (inbox != null);
	}

	private boolean canSend() {
		return (drafts != null && submissionUri != null);
	}

	private void listFolder(DefaultHandler handler, String folder) throws Exception {
		synchronized (this) {
			if (!isConnected()) {
				throw new IllegalStateException("Not connected.");
			}
			HttpClient client = getClient();
			ExchangeMethod op = new ExchangeMethod(SEARCH_METHOD, folder);
			op.setHeader("Content-Type", XML_CONTENT_TYPE);
			if (limit > 0)
				op.setHeader("Range", "rows=0-" + limit);
			op.setHeader("Brief", "t");

			/* Mirco: Manage of custom query */
			if ((filterLastCheck == null || "".equals(filterLastCheck)) && (filterFrom == null || "".equals(filterFrom)) && (filterNotFrom == null || "".equals(filterNotFrom)) && (filterTo == null || "".equals(filterTo))) {
				op.setRequestEntity(unfiltered ? createAllInboxEntity() : createUnreadInboxEntity());
			} else {
				op.setRequestEntity(createCustomInboxEntity(unfiltered, filterLastCheck, filterFrom, filterNotFrom, filterTo));
			}
			InputStream stream = null;
			try {
				int status = client.executeMethod(op);
				stream = op.getResponseBodyAsStream();
				if (status >= 300) {
					throw new IllegalStateException("Unable to obtain " + folder + ".");
				}
				SAXParserFactory spf = SAXParserFactory.newInstance();
				spf.setNamespaceAware(true);
				SAXParser parser = spf.newSAXParser();
				parser.parse(stream, handler);
				stream.close();
				stream = null;
			} finally {
				try {
					if (stream != null) {
						byte[] buf = new byte[65536];
						try {
							if (session.getDebug()) {
								PrintStream log = session.getDebugOut();
								log.println("Response Body:");
								int count;
								while ((count = stream.read(buf, 0, 65536)) != -1) {
									log.write(buf, 0, count);
								}
								log.flush();
								log.println();
							} else {
								while (stream.read(buf, 0, 65536) != -1)
									;
							}
						} catch (Exception ignore) {
						} finally {
							try {
								stream.close();
							} catch (Exception ignore2) {
							}
						}
					}
				} finally {
					op.releaseConnection();
				}
			}
		}
	}

	private void findInbox() throws Exception {
		inbox = null;
		drafts = null;
		submissionUri = null;
		sentitems = null;
		outbox = null;
		HttpClient client = getClient();
		ExchangeMethod op = new ExchangeMethod(PROPFIND_METHOD, server + "/exchange/" + mailbox);
		op.setHeader("Content-Type", XML_CONTENT_TYPE);
		op.setHeader("Depth", "0");
		op.setHeader("Brief", "t");
		op.setRequestEntity(createFindInboxEntity());
		InputStream stream = null;
		try {
			int status = client.executeMethod(op);
			stream = op.getResponseBodyAsStream();
			if (status >= 300) {
				throw new IllegalStateException("Unable to obtain inbox.");
			}
			SAXParserFactory spf = SAXParserFactory.newInstance();
			spf.setNamespaceAware(true);
			SAXParser parser = spf.newSAXParser();
			parser.parse(stream, new DefaultHandler() {
				private final StringBuilder content = new StringBuilder();

				public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
					content.setLength(0);
				}

				public void characters(char[] ch, int start, int length) throws SAXException {
					content.append(ch, start, length);
				}

				public void endElement(String uri, String localName, String qName) throws SAXException {
					if (!HTTPMAIL_NAMESPACE.equals(uri))
						return;
					if ("inbox".equals(localName)) {
						ExchangeConnection.this.inbox = content.toString();
					} else if ("drafts".equals(localName)) {
						ExchangeConnection.this.drafts = content.toString();
					} else if ("sentitems".equals(localName)) {
						ExchangeConnection.this.sentitems = content.toString();
					} else if ("outbox".equals(localName)) {
						ExchangeConnection.this.outbox = content.toString();
					} else if ("sendmsg".equals(localName)) {
						ExchangeConnection.this.submissionUri = content.toString();
					}
				}
			});
			stream.close();
			stream = null;
		} finally {
			try {
				if (stream != null) {
					byte[] buf = new byte[65536];
					try {
						if (session.getDebug()) {
							PrintStream log = session.getDebugOut();
							log.println("Response Body:");
							int count;
							while ((count = stream.read(buf, 0, 65536)) != -1) {
								log.write(buf, 0, count);
							}
							log.flush();
							log.println();
						} else {
							while (stream.read(buf, 0, 65536) != -1)
								;
						}
					} catch (Exception ignore) {
					} finally {
						try {
							stream.close();
						} catch (Exception ignore2) {
						}
					}
				}
			} finally {
				op.releaseConnection();
			}
		}
	}

	private HttpClient getClient() {
		synchronized (this) {
			if (client == null) {
				client = new HttpClient();
				if (timeout > 0)
					client.getParams().setSoTimeout(timeout);
				if (connectionTimeout > 0) {
					client.getHttpConnectionManager().getParams().setConnectionTimeout(connectionTimeout);
				}
				if (localAddress != null) {
					client.getHostConfiguration().setLocalAddress(localAddress);
				}
			}
			return client;
		}
	}

	private void signOn() throws Exception {
		HttpClient client = getClient();
		URL serverUrl = new URL(server);
		String host = serverUrl.getHost();
		int port = serverUrl.getPort();
		if (port == -1)
			port = serverUrl.getDefaultPort();
		AuthScope authScope = new AuthScope(host, port);

		if (username.indexOf("\\") < 0) {
			client.getState().setCredentials(authScope, new UsernamePasswordCredentials(username, password));
		} else {
			// Try to connect with NTLM authentication
			String domainUser = username.substring(username.indexOf("\\") + 1, username.length());
			String domain = username.substring(0, username.indexOf("\\"));
			client.getState().setCredentials(authScope, new NTCredentials(domainUser, password, host, domain));
		}

		boolean authenticated = false;
		OptionsMethod authTest = new OptionsMethod(server + "/exchange");
		try {
			authenticated = (client.executeMethod(authTest) < 400);
		} finally {
			try {
				InputStream stream = authTest.getResponseBodyAsStream();
				byte[] buf = new byte[65536];
				try {
					if (session.getDebug()) {
						PrintStream log = session.getDebugOut();
						log.println("Response Body:");
						int count;
						while ((count = stream.read(buf, 0, 65536)) != -1) {
							log.write(buf, 0, count);
						}
						log.flush();
						log.println();
					} else {
						while (stream.read(buf, 0, 65536) != -1)
							;
					}
				} catch (Exception ignore) {
				} finally {
					try {
						stream.close();
					} catch (Exception ignore2) {
					}
				}
			} finally {
				authTest.releaseConnection();
			}
		}
		if (!authenticated) {
			PostMethod op = new PostMethod(server + SIGN_ON_URI);
			op.setRequestHeader("Content-Type", FORM_URLENCODED_CONTENT_TYPE);
			op.addParameter("destination", server + "/exchange");
			op.addParameter("flags", "0");
			op.addParameter("username", username);
			op.addParameter("password", password);
			try {
				int status = client.executeMethod(op);
				if (status >= 400) {
					throw new IllegalStateException("Sign-on failed: " + status);
				}
			} finally {
				try {
					InputStream stream = op.getResponseBodyAsStream();
					byte[] buf = new byte[65536];
					try {
						if (session.getDebug()) {
							PrintStream log = session.getDebugOut();
							log.println("Response Body:");
							int count;
							while ((count = stream.read(buf, 0, 65536)) != -1) {
								log.write(buf, 0, count);
							}
							log.flush();
							log.println();
						} else {
							while (stream.read(buf, 0, 65536) != -1)
								;
						}
					} catch (Exception ignore) {
					} finally {
						try {
							stream.close();
						} catch (Exception ignore2) {
						}
					}
				} finally {
					op.releaseConnection();
				}
			}
		}
		findInbox();
	}

	private RequestEntity createMessageEntity(MimeMessage message) throws Exception {
		final File tempFile = File.createTempFile("exmail", null, null);
		tempFile.deleteOnExit();
		OutputStream output = new BufferedOutputStream(new FileOutputStream(tempFile));
		message.writeTo(output);
		if (session.getDebug()) {
			PrintStream log = session.getDebugOut();
			log.println("Message Content:");
			message.writeTo(log);
			log.println();
			log.flush();
		}
		output.flush();
		output.close();
		InputStream stream = new FileInputStream(tempFile) {
			public void close() throws IOException {
				try {
					super.close();
				} finally {
					try {
						if (!tempFile.delete())
							tempFile.deleteOnExit();
					} catch (Exception ignore) {
					}
				}
			}
		};
		return new InputStreamRequestEntity(stream, tempFile.length(), MESSAGE_CONTENT_TYPE);
	}

	private static RequestEntity createFindInboxEntity() throws Exception {
		synchronized (ExchangeConnection.class) {
			if (findInboxEntity == null) {
				DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
				dbf.setNamespaceAware(true);
				Document doc = dbf.newDocumentBuilder().newDocument();
				Element propfind = doc.createElementNS(DAV_NAMESPACE, "propfind");
				doc.appendChild(propfind);
				Element prop = doc.createElementNS(DAV_NAMESPACE, "prop");
				propfind.appendChild(prop);
				Element inbox = doc.createElementNS(HTTPMAIL_NAMESPACE, "inbox");
				prop.appendChild(inbox);
				Element drafts = doc.createElementNS(HTTPMAIL_NAMESPACE, "drafts");
				prop.appendChild(drafts);
				Element sendmsg = doc.createElementNS(HTTPMAIL_NAMESPACE, "sendmsg");
				prop.appendChild(sendmsg);
				Element outbox = doc.createElementNS(HTTPMAIL_NAMESPACE, "outbox");
				prop.appendChild(outbox);
				Element sentitems = doc.createElementNS(HTTPMAIL_NAMESPACE, "sentitems");
				prop.appendChild(sentitems);

				// http://msdn.microsoft.com/en-us/library/ms992623(EXCHG.65).aspx

				ByteArrayOutputStream collector = new ByteArrayOutputStream();
				Transformer transformer = TransformerFactory.newInstance().newTransformer();
				transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
				transformer.transform(new DOMSource(doc), new StreamResult(collector));
				findInboxEntity = collector.toByteArray();
			}
			return new ByteArrayRequestEntity(findInboxEntity, XML_CONTENT_TYPE);
		}
	}

	private static RequestEntity createUnreadInboxEntity() throws Exception {
		synchronized (ExchangeConnection.class) {
			if (unreadInboxEntity == null) {
				unreadInboxEntity = createSearchEntity(new String(getResource(GET_UNREAD_MESSAGES_SQL_RESOURCE), "UTF-8"));
			}
			return new ByteArrayRequestEntity(unreadInboxEntity, XML_CONTENT_TYPE);
		}
	}

	private static RequestEntity createAllInboxEntity() throws Exception {
		synchronized (ExchangeConnection.class) {
			if (allInboxEntity == null) {
				allInboxEntity = createSearchEntity(new String(getResource(GET_ALL_MESSAGES_SQL_RESOURCE), "UTF-8"));
			}
			return new ByteArrayRequestEntity(allInboxEntity, XML_CONTENT_TYPE);
		}
	}

	private static RequestEntity createCustomInboxEntity(boolean unfiltered, String filterLastCheck, String filterFrom, String filterNotFrom, String filterTo) throws Exception {
		synchronized (ExchangeConnection.class) {
			/*
			 * If user has to use filter base on a adte we have to build a new
			 * filter
			 */
			// if (customInboxEntity == null) {
			customInboxEntity = createSearchEntity(new String(getResource(GET_FILTERED_MESSAGES_SQL_RESOURCE), "UTF-8"));
			// }

			/* Mirco: Replace Filtes */
			String filter = new String(customInboxEntity);
			if (!unfiltered) {
				filter = filter.replace(BOOKMARK_FILTER_UNREADED, "AND \"urn:schemas:httpmail:read\" = False");
			} else {
				filter = filter.replace(BOOKMARK_FILTER_UNREADED, "");
			}
			if (filterLastCheck != null && !"".equals(filterLastCheck)) {
				// Es. AND "urn:schemas:httpmail:datereceived" >
				// CAST("2010-08-04T00:00:00Z" as 'dateTime')
				filter = filter.replace(BOOKMARK_FILTER_LAST_CHECK, "AND \"urn:schemas:httpmail:datereceived\" > CAST(\"" + filterLastCheck + "\" as 'dateTime')");
			} else {
				filter = filter.replace(BOOKMARK_FILTER_LAST_CHECK, "");
			}
			if (filterFrom != null && !"".equals(filterFrom)) {
				// Es. AND "urn:schemas:httpmail:fromemail" LIKE '@domain.com%'
				filter = filter.replace(BOOKMARK_FILTER_FROM, "AND \"urn:schemas:httpmail:fromemail\" LIKE '%" + filterFrom + "%'");
			} else {
				filter = filter.replace(BOOKMARK_FILTER_FROM, "");
			}
			if (filterNotFrom != null && !"".equals(filterNotFrom)) {
				if (filterNotFrom.indexOf(";") > 0) {
					StringBuilder sb = new StringBuilder();
					// sb.append("AND (");
					for (String aFilter : filterNotFrom.split(";")) {
						sb.append("AND \"urn:schemas:httpmail:fromemail\" NOT LIKE '%" + aFilter + "%'");
					}
					// sb.append(")");

					filter = filter.replace(BOOKMARK_FILTER_NOT_FROM, sb.toString());
				} else {
					// Es. AND "urn:schemas:httpmail:fromemail" LIKE
					// '@domain.com%'
					filter = filter.replace(BOOKMARK_FILTER_NOT_FROM, "AND \"urn:schemas:httpmail:fromemail\" NOT LIKE '%" + filterNotFrom + "%'");
				}
			} else {
				filter = filter.replace(BOOKMARK_FILTER_NOT_FROM, "");
			}
			if (filterTo != null && !"".equals(filterTo)) {
				// Es. AND "urn:schemas:httpmail:to" LIKE '%test@domain.com%'
				filter = filter.replace(BOOKMARK_FILTER_TO, "AND \"urn:schemas:httpmail:to\" LIKE '%" + filterTo + "%'");
			} else {
				filter = filter.replace(BOOKMARK_FILTER_TO, "");
			}

			customInboxEntity = filter.getBytes();
			return new ByteArrayRequestEntity(customInboxEntity, XML_CONTENT_TYPE);
		}
	}

	private static byte[] createSearchEntity(String sqlString) throws Exception {
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		dbf.setNamespaceAware(true);
		Document doc = dbf.newDocumentBuilder().newDocument();
		Element searchRequest = doc.createElementNS(DAV_NAMESPACE, "searchrequest");
		doc.appendChild(searchRequest);
		Element sql = doc.createElementNS(DAV_NAMESPACE, "sql");
		searchRequest.appendChild(sql);
		sql.appendChild(doc.createTextNode(sqlString));
		ByteArrayOutputStream collector = new ByteArrayOutputStream();
		Transformer transformer = TransformerFactory.newInstance().newTransformer();
		transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
		transformer.transform(new DOMSource(doc), new StreamResult(collector));
		return collector.toByteArray();
	}

	private static RequestEntity createDeleteEntity(List<ExchangeMessage> messages) throws Exception {
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		dbf.setNamespaceAware(true);
		Document doc = dbf.newDocumentBuilder().newDocument();
		Element delete = doc.createElementNS(DAV_NAMESPACE, "delete");
		doc.appendChild(delete);
		Element target = doc.createElementNS(DAV_NAMESPACE, "target");
		delete.appendChild(target);
		for (ExchangeMessage message : messages) {
			String url = message.getUrl();
			Element href = doc.createElementNS(DAV_NAMESPACE, "href");
			target.appendChild(href);
			String file = url.substring(url.lastIndexOf("/") + 1);
			href.appendChild(doc.createTextNode(file));
		}
		ByteArrayOutputStream collector = new ByteArrayOutputStream();
		Transformer transformer = TransformerFactory.newInstance().newTransformer();
		transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
		transformer.transform(new DOMSource(doc), new StreamResult(collector));
		return new ByteArrayRequestEntity(collector.toByteArray(), XML_CONTENT_TYPE);
	}

	private RequestEntity createAddBccEntity(Address[] addresses) throws Exception {
		StringBuilder recipientList = new StringBuilder();
		for (Address address : addresses) {
			if (recipientList.length() != 0)
				recipientList.append(';');
			recipientList.append(((InternetAddress) address).getAddress());
		}
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		dbf.setNamespaceAware(true);
		Document doc = dbf.newDocumentBuilder().newDocument();
		Element propertyUpdate = doc.createElementNS(DAV_NAMESPACE, "propertyupdate");
		doc.appendChild(propertyUpdate);
		Element set = doc.createElementNS(DAV_NAMESPACE, "set");
		propertyUpdate.appendChild(set);
		Element prop = doc.createElementNS(DAV_NAMESPACE, "prop");
		set.appendChild(prop);
		Element bcc = doc.createElementNS(MAILHEADER_NAMESPACE, "bcc");
		prop.appendChild(bcc);
		bcc.appendChild(doc.createTextNode(recipientList.toString()));
		ByteArrayOutputStream collector = new ByteArrayOutputStream();
		Transformer transformer = TransformerFactory.newInstance().newTransformer();
		transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
		if (session.getDebug()) {
			transformer.transform(new DOMSource(doc), new StreamResult(session.getDebugOut()));
			session.getDebugOut().println();
		}
		transformer.transform(new DOMSource(doc), new StreamResult(collector));
		return new ByteArrayRequestEntity(collector.toByteArray(), XML_CONTENT_TYPE);
	}

	private static RequestEntity createMarkReadEntity(List<ExchangeMessage> messages) throws Exception {
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		dbf.setNamespaceAware(true);
		Document doc = dbf.newDocumentBuilder().newDocument();
		Element propertyUpdate = doc.createElementNS(DAV_NAMESPACE, "propertyupdate");
		doc.appendChild(propertyUpdate);
		Element target = doc.createElementNS(DAV_NAMESPACE, "target");
		propertyUpdate.appendChild(target);
		for (ExchangeMessage message : messages) {
			String url = message.getUrl();
			Element href = doc.createElementNS(DAV_NAMESPACE, "href");
			target.appendChild(href);
			String file = url.substring(url.lastIndexOf("/") + 1);
			href.appendChild(doc.createTextNode(file));
		}
		Element set = doc.createElementNS(DAV_NAMESPACE, "set");
		propertyUpdate.appendChild(set);
		Element prop = doc.createElementNS(DAV_NAMESPACE, "prop");
		set.appendChild(prop);
		Element read = doc.createElementNS(HTTPMAIL_NAMESPACE, "read");
		prop.appendChild(read);
		read.appendChild(doc.createTextNode("1"));
		ByteArrayOutputStream collector = new ByteArrayOutputStream();
		Transformer transformer = TransformerFactory.newInstance().newTransformer();
		transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
		transformer.transform(new DOMSource(doc), new StreamResult(collector));
		return new ByteArrayRequestEntity(collector.toByteArray(), XML_CONTENT_TYPE);
	}

	private static byte[] getResource(String resource) {
		if (resource == null)
			return null;
		synchronized (RESOURCES) {
			byte[] content = RESOURCES.get(resource);
			if (content != null)
				return content;
			try {
				InputStream input = ExchangeConnection.class.getResourceAsStream(resource);
				ByteArrayOutputStream collector = new ByteArrayOutputStream();
				byte[] buf = new byte[65536];
				int count;
				while ((count = input.read(buf, 0, 65536)) != -1) {
					collector.write(buf, 0, count);
				}
				input.close();
				collector.flush();
				content = collector.toByteArray();
				RESOURCES.put(resource, content);
				return content;
			} catch (Exception ex) {
				throw new IllegalStateException(ex.getMessage(), ex);
			}
		}
	}

	private static String escape(String url) {
		StringBuilder collector = new StringBuilder(url);
		for (int i = collector.length() - 1; i >= 0; i--) {
			int value = (int) collector.charAt(i);
			if (value > 127 || !ALLOWED_CHARS[value]) {
				collector.deleteCharAt(i);
				collector.insert(i, HEXABET[value & 0x0f]);
				value >>>= 4;
				collector.insert(i, HEXABET[value & 0x0f]);
				value >>>= 4;
				collector.insert(i, '%');
				if (value > 0) {
					collector.insert(i, HEXABET[value & 0x0f]);
					value >>>= 4;
					collector.insert(i, HEXABET[value & 0x0f]);
					collector.insert(i, '%');
				}
			}
		}
		return collector.toString();
	}

	private static String generateMessageName() {
		synchronized (RANDOM) {
			return new BigInteger(200, RANDOM).toString(Character.MAX_RADIX);
		}
	}

	private static Properties parseOptions(String options) throws Exception {
		StringBuilder collector = new StringBuilder();
		String[] nvPairs = options.split("[,;]");
		for (String nvPair : nvPairs)
			collector.append(nvPair).append('\n');
		Properties properties = new Properties();
		properties.load(new ByteArrayInputStream(collector.toString().getBytes("ISO-8859-1")));
		return properties;
	}

}
