# cookie-management-tool
Cookie Management Tool


This tool has been written in C# programming language what makes it runnable only on relatively modern distributions of Windows operating system. Currently it has been tested only on three instances of Windows 10 operating system. Immediately after application start-up, detection of user's currently installed web-browser is being performed (auto-detection is being done with checking web-browsers' registry keys). At the moment when this description was written, following web-browsers were supported: 
* Mozilla Firefox
* Google Chrome
* Internet Explorer
* Microsoft Edge

By selecting one of the options from the drop-down list with web-browsers, in tab-control tabs show up whose names are equal to present user profiles of the selected web-browser. By opening one of these tabs, in main table is being displayed list of all cookies of the selected web-browser and its user profile. As Firefox and Chrome are multi-profile web-browsers, their belonging profiles are being displayed in tab-control when any of them is selected. On the other hand, Internet Explorer has two default profiles – one of them is active when IE is used in Protected (high security-level) mode, and the another one otherwise (Protected mode is activated/deactivated in IE's settings).
![Firefox cookies](/images/firefox_cookies.png?raw=true "Displayed all cookies of default profile of Mozilla Firefox.")<br/>
It is also possible in drop-down list to choose option that displays cookies of all installed web-browsers. When using that option, cookies of all available profiles are being shown.
Searching and filtering cookies is also supported – to do so, it is required to enter arbitrary consecutive array of characters in search-bar (which is located on the top of the window) and the cookie list is immediately updated according to specified criterias. From the drop-down list it is possible to select attribute name which value we want to use for filtering. It is also possible to select option that performs search of cookies whose any attribute value (like cookie's web-domain, name/key or value) contains specified term. There are also other filtering options (like case sensitivity and exact word matching) that can be set by checking/unchecking checkboxes.
![All cookies](/images/all_browsers_filtered_cookies.png?raw=true "Displayed all cookies all web-browsers and their all profiles.")<br/>
After selecting one or more cookies from the table, it is possible to perform deletion with click on _Delete_ button. Deletion can even be performed when cookies of all profiles of all web-browsers are displayed.
Operations like addition and modification of cookies are also supported. Unfortunately, in current version of this application it is not possible to add new cookies or update current ones for Internet Edge web-browser. In order to update cookie, user first has to select cookie in main window that they want to update and after that _Edit_ button has to be clicked.
![Add/Edit cookie](/images/add_edit_cookie.png?raw=true "Displayed window for adding/editing cookie.")<br/>
This tools also allows commiting insertion and modification changes of cookies into multiple web-browsers – to do that, in _Add/Edit cookie_ window user has to select destination web-browsers in which they want to import changes. As mentioned before, some web-browsers can have multiple profiles, therefore after selection of destination web-browsers it is also possible to select destination profiles where changes should be stored.

![Selecting destination web-browser profiles](/images/destination_browser_profiles.png?raw=true "Displayed window for selecting cookie storage-destination web-browser profiles.")<br/>
Google Chrome allows storing encrypted cookie values instead of regular unsafe ones, therefore it is possible on each cookie insertion or modification to select if we want to store its value encrypted or not. If cookie value was until now stored as encrypted, it is possible to store it normally and vice versa.
![Selecting if cookie value should be stored encrypted or not](/images/chrome_cookie_value_encryption_option.png?raw=true "Selecting if cookie value should be stored encrypted or not")<br/>
When storing cookies into Internet Explorer, that need to be available in its Protected mode, application needs to run subprogram (that is also included in this solution) that performs same task as main program when it normally stores cookies, but executes that action in low integrity-level process because those cookies are stored in directory with higher level of access control.
