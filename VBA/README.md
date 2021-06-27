The VBA source code is used in a standard UserForm module, in a Microsoft Word application.
I tested it using Office 64 bit, and it performs the same functions, without p-code decompiling, of the Delphi application,
obviously much more slowly.

The VBA projects doesn't require trusted access to the projects, and doesn't load the Office documents using the
standard applications, so it's safe to use on any Office document.
