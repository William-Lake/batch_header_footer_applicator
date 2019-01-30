==============================
batch_header_footer_applicator
==============================


.. image:: https://img.shields.io/pypi/v/batch_header_footer_applicator.svg
        :target: https://pypi.python.org/pypi/batch_header_footer_applicator

.. image:: https://img.shields.io/travis/William-Lake/batch_header_footer_applicator.svg
        :target: https://travis-ci.org/William-Lake/batch_header_footer_applicator

.. image:: https://readthedocs.org/projects/batch-header-footer-applicator/badge/?version=latest
        :target: https://batch-header-footer-applicator.readthedocs.io/en/latest/?badge=latest
        :alt: Documentation Status




Applies the header and footer from a template word document to a batch of other word documents.


* Free software: MIT license
* Documentation: https://batch-header-footer-applicator.readthedocs.io.

Dependencies
--------

If you'd only like to use this module, you'll only need the following dependencies:

- pywin32
- PySimpleGui

Which can be installed via the requirements.txt file: `pip install -r requirements.txt`

If you'd like to develop the module as well, you'll need the requirements_dev.txt file: `pip install -r requirements_dev.txt`

Usage
--------

#. Ensure you have python3 and the dependencies installed.
#. Open a terminal in the same directory as search_word_docs.py
#. Execute `python search_word_docs.py`
#. Select a directory to search for word files.
#. Provide a search term.
#. Select/UnSelect the Recursive Searching Textbox.
#. Click Search.
#. If desired, click the 'Save' button and determine where you'd like to save the results.

Updates are provided throughout the process, when finished the results will be provided in the bottom-most text box.

Credits
-------

This package was created with Cookiecutter_ and the `audreyr/cookiecutter-pypackage`_ project template.

.. _Cookiecutter: https://github.com/audreyr/cookiecutter
.. _`audreyr/cookiecutter-pypackage`: https://github.com/audreyr/cookiecutter-pypackage
