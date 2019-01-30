==============================
batch_header_footer_applicator
==============================

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
#. Open a terminal in the root directory of batch_header_footer_applicator
#. Execute `python batch_header_footer_applicator`
#. Use the first 'Browse' button to navigate to and select a text file containing the paths to the documents you'd like the header/footer applied to.
#. Use the second 'Browse' button to navigate to and select the Word file containing the new header/footer to apply.
#. Click 'Apply' to start the process.

Updates are provided throughout the process, when finished the results will be provided in the bottom-most text box.

Credits
-------

This package was created with Cookiecutter_ and the `audreyr/cookiecutter-pypackage`_ project template.

.. _Cookiecutter: https://github.com/audreyr/cookiecutter
.. _`audreyr/cookiecutter-pypackage`: https://github.com/audreyr/cookiecutter-pypackage
