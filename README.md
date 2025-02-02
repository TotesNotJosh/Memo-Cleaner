# Memo-Cleaner
A simple macro for MS outlook that will scrub a received email's body of special space characters and a couple invisible characters to help hide the identity of whistle blowers. Companies like Tesla, SpaceX, X.com, and government agencies like DOGE have been known to set up special gotcha emails to fire whistleblowers. Be safe by cleaning the emails before leaking them so that people with strong ethics can continue to work for companies and agencies that take retaliatory measures against ethical employees. This is not a comprehensive list of hidden characters that these organizations might use you can help contribute by raising issues with more characters listed, or suggestions.
# Installation
1. Open Outlook
2. Press Alt-F11
3. Right click under Project on the left hand side and insert a module.
4. Paste the memo cleaner.bas text into the module.
5. Ctrl-S to save
# Using
To use this, select the email you wish to clean of special whitespace characters. click Alt-F8 and select clean_email_body and hit Run. This will edit the body of the email itself to remove whitespace characters. The whitespace characters came from these two websites.
<br>https://www.compart.com/en/unicode/bidiclass/WS
<br>https://www.compart.com/en/unicode/category/Zs
