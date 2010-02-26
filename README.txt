                exJello - JavaMail Provider for Exchange

exJello is a JavaMail provider that connects to a Microsoft Exchange server
(actually, it uses the WebDAV interface exposed by Outlook Web Access).  It is
designed as a drop-in replacement for the standard POP3 and SMTP providers.
This allows you to send and receive messages through your Exchange server in
situations where a POP3/SMTP interface is not available (through a restrictive
firewall, for example, or if your administrator simply does not provide a POP3
or SMTP gateway).

Installation and configuration documentation is provided in the "doc"
subdirectory.

exJello is built using the Apache Maven2 framework.  To compile from source
(assuming you have installed Maven), simply open a command line in the
directory where you expanded the exJello archive and type:

    mvn assembly:assembly

The binary and source distributions will be built under the "target"
subdirectory.
