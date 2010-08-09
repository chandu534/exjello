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

import java.io.InputStream;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.Vector;

import javax.mail.Flags;
import javax.mail.Folder;
import javax.mail.FolderNotFoundException;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.MethodNotSupportedException;
import javax.mail.Store;

import javax.mail.event.ConnectionEvent;

class ExchangeFolder extends Folder {

    public static final String INBOX = "INBOX";
    public static final String SENTITEMS = "SENT ITEMS";
    public static final String OUTBOX = "OUTBOX";
    public static final String DRAFT = "DRAFT";

    private static final String ROOT = "";

    private final String name;

    private final ExchangeConnection connection;

    private List<String> messages;

    private Vector<ExchangeMessage> cache;

    private Set<InputStream> openStreams;

    private boolean open = false;

	public ExchangeFolder(ExchangeStore store, String name,
            ExchangeConnection connection) throws MessagingException {
		super(store);
        this.name = name;
        this.connection = connection;
	}

	public void appendMessages(Message[] messages) throws MessagingException {
        throw new MethodNotSupportedException("appendMessages");
	}

	public void close(boolean expunge) throws MessagingException {
        synchronized (this) {
            if (!isOpen()) throw new IllegalStateException("Already closed.");
            if (ROOT.equals(getName())) {
                throw new MethodNotSupportedException("close");
            }
            try {
                if (expunge && mode == READ_WRITE) {
                    List<ExchangeMessage> deletedMessages =
                            new ArrayList<ExchangeMessage>();
                    for (int i = cache.size() - 1; i >= 0; i--) {
                        ExchangeMessage message = cache.get(i);
                        if (message != null &&
                                message.isSet(Flags.Flag.DELETED)) {
                            deletedMessages.add(message);
                        }
                    }
                    if (!deletedMessages.isEmpty()) {
                        try {
                            connection.delete(deletedMessages);
                        } catch (Exception ex) {
                            throw new MessagingException(ex.getMessage(), ex);
                        }
                    }
                }
            } finally {
                cache.clear();
                if (openStreams != null) {
                    for (InputStream stream : openStreams) {
                        try {
                            stream.close();
                        } catch (Exception ignore) { }
                    }
                    openStreams.clear();
                }
                openStreams = null;
                cache = null;
                open = false;
            }
        }
        notifyConnectionListeners(ConnectionEvent.CLOSED);
	}

	public boolean create(int type) throws MessagingException {
        return false;
	}

	public boolean delete(boolean recurse) throws MessagingException {
        throw new MethodNotSupportedException("delete");
	}

	public boolean exists() throws MessagingException {
        String name = getName();
        return INBOX.equalsIgnoreCase(name) ||
                SENTITEMS.equalsIgnoreCase(name) ||
                OUTBOX.equalsIgnoreCase(name) ||
                DRAFT.equalsIgnoreCase(name) ||
                ROOT.equals(name);
	}

	public Message[] expunge() throws MessagingException {
        throw new MethodNotSupportedException("expunge");
	}

	public Folder getFolder(String name) throws MessagingException {
        if (!ROOT.equals(getName())) {
            throw new MessagingException("Hierarchy not supported.");
        }
        
        if (INBOX.equalsIgnoreCase(name)) return getStore().getFolder(INBOX);
        if (SENTITEMS.equalsIgnoreCase(name)) {
            return getStore().getFolder(SENTITEMS);
        }
        if (DRAFT.equalsIgnoreCase(name)) return getStore().getFolder(DRAFT);
        if (OUTBOX.equalsIgnoreCase(name)) return getStore().getFolder(OUTBOX);
        throw new MessagingException("Folder not supported.");
	}

	public String getFullName() {
        return getName();
	}

	public Message getMessage(int messageNumber) throws MessagingException {
        if (!isOpen()) throw new IllegalStateException("Folder is closed.");
        if (!exists()) throw new FolderNotFoundException(this);
        if (ROOT.equals(getName())) {
            throw new MethodNotSupportedException("getMessage");
        }
        synchronized (this) {
            int index = messageNumber - 1;
            ExchangeMessage message = cache.get(index);
            if (message == null) {
                cache.set(index, (message = new ExchangeMessage(this,
                        messageNumber, messages.get(index), connection)));
            }
            return message;
        }
	}

	public int getMessageCount() throws MessagingException {
        if (!isOpen()) return -1;
        if (!exists()) throw new FolderNotFoundException(this);
        if (ROOT.equals(getName())) return 0;
        synchronized (this) {
            return messages.size();
        }
	}

	public String getName() {
		return name;
	}

	public Folder getParent() throws MessagingException {
        String name = getName();
        if (name == null) name = ROOT;
        return ROOT.equals(name) ? null : new ExchangeFolder((ExchangeStore)
                getStore(), ROOT, connection);
	}

	public Flags getPermanentFlags() {
		return new Flags();
	}

	public char getSeparator() throws MessagingException {
        return ROOT.equals(getName()) ? '/' : '\u0000';
	}

	public int getType() throws MessagingException {
        return ROOT.equals(getName()) ? HOLDS_FOLDERS : HOLDS_MESSAGES;
	}

	public boolean hasNewMessages() throws MessagingException {
		return false;
	}

	public boolean isOpen() {
        if (ROOT.equals(getName())) return false;
        synchronized (this) {
            return open;
        }
	}

	public Folder[] list(String pattern) throws MessagingException {
        String name = getName();
        if (!ROOT.equals(name)) {
            throw new MessagingException("Hierarchy not supported.");
        }
        /* return only supported folders */
        return new Folder[] {
            getStore().getFolder(INBOX),
            getStore().getFolder(SENTITEMS),
            getStore().getFolder(OUTBOX),
            getStore().getFolder(DRAFT)
        };
	}

	public void open(int mode) throws MessagingException {
        if (isOpen()) throw new IllegalStateException("Folder is open.");
        if (!exists()) throw new FolderNotFoundException(this);
        if (ROOT.equals(getName())) {
            throw new MethodNotSupportedException("open");
        }
        ExchangeStore store = (ExchangeStore) getStore();
        synchronized (this) {
            this.mode = mode;
            try {
                messages = connection.getMessages(name);
            } catch (MessagingException ex) {
                throw ex;
            } catch (Exception ex) {
                throw new MessagingException(ex.getMessage(), ex);
            }
            cache = new Vector<ExchangeMessage>(messages.size());
            cache.setSize(cache.capacity());
            open = true;
        }
        notifyConnectionListeners(ConnectionEvent.OPENED);
	}

	public boolean renameTo(Folder target) throws MessagingException {
		throw new MethodNotSupportedException("renameTo");
	}

    public void register(InputStream stream) throws MessagingException {
        if (!isOpen()) throw new IllegalStateException("Folder is closed.");
        if (!exists()) throw new FolderNotFoundException(this);
        synchronized (this) {
            if (openStreams == null) openStreams = new HashSet<InputStream>();
            openStreams.add(stream);
        }
    }

}

