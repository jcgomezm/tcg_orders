TCG Orders
==============================================

**TCG Orders** is a script that helps you create card orders faster and easier to send to your favourite supplier.

Installation
------------

Install the required libraries using pipenv:

    $ pipenv install

Usage
------------

Copy your list on the folder:

    imports/<filename>.txt

The list should look like this:

    <your name>
    <quantity> <item URL>
    <quantity> <item URL>
    <quantity> <item URL>
    ...

Run the script:

    $ python main.py <filename>

Find your order on the exports folder:

    exports/<filename>.xlsx
    
Possible future features
----------

-   Support for custom Excel headers.
-   Support for other retail websites.
-   Support for other card conditions (current support only for near mint).
-   Support for Google Sheets.
-   Front-end development for web app.
