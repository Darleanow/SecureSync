SecureSync
==========

What is it about ?
------------------

SecureSync is a personal project that I carried out to streamline our work as CyberSecurity Consultants.
Its purpose is to merge the safety reports into an Excel file containing all the data, similar to a data scrapper.
It was designed to work with both `PingCastle <https://www.pingcastle.com/>`_ and `Purple Knight <https://www.purple-knight.com/>`_.

Installation
------------

SecureSync is very simple to install and operate.

``bash
$ git clone "https://github.com/Darleanow/SecureSync"
$ cd SecureSync
$ pip install -r requirements.txt
$ python3 main.py
``

Use
---

To use SecureSync, just open it.
You will be asked to enter the paths of the reports.
As Purple Knight reports comes out as .pdf files, files, you must convert it to a Word document, please use `Adobe <https://www.adobe.com/fr/acrobat/online/pdf-to-word.html>`_  
as it's the most accurate one, you might lose data if you use another one.
PLEASE NOTE that providing wrong paths, wrong file extensions may lead to unexpected behavior.
