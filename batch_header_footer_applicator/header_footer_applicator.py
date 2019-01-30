import logging
from pathlib import Path
import win32com.client as win32

class HeaderFooterApplicator(object):

    def apply_header_footer(self, file_path_doc, template_doc, update_text_callback):

        self.__update_text_callback = update_text_callback

        self.__update_text_callback(f'Applying template in {template_doc} to the files listed in {file_path_doc}.')

        # Create the Microsoft Word Instance
        self.__msword = win32.gencache.EnsureDispatch('Word.Application')

        # Open the template in Word
        # https://docs.microsoft.com/en-us/office/vba/api/word.documents.open
        self.__template_word_doc = self.__msword.Documents.Open(Path(template_doc).__str__(), Visible = False) 

        self.__doc_paths = [Path(doc_path.strip()).__str__() for doc_path in open(file_path_doc).readlines()]

        self.__apply_template_header_footer_to_target_docs()

        self.__update_text_callback('Process Complete', do_replace = True)

        self.__template_word_doc.Close()

        self.__msword.Quit()

    def __apply_template_header_footer_to_target_docs(self):
        '''Applys the Header/Footer in the template Word Document to the target Word Documents.'''

        logging.info('Applying Template Header/Footer to target word docs.')

        self.__update_text_callback(f'Applying template to {len(self.__doc_paths)} docs.')

        for doc_path_index, doc_path in enumerate(self.__doc_paths):

            logging.debug(f'Working on {doc_path_index + 1} out of {len(self.__doc_paths)}: {doc_path}')

            self.__update_text_callback(f'Working on {doc_path_index + 1} out of {len(self.__doc_paths)}: {doc_path}',do_replace=True)

            try:

                target_word_doc = self.__msword.Documents.Open(doc_path, Visible = False) 

            except Exception as e:

                self.__update_text_callback(f'Error while opening {doc_path}!\n{str(e)}')

                logging.warn(f'Error while opening {doc_path}!\n{str(e)}')

                continue

            # Gather number of pages so you can determine if additional pages were added.
            num_pages_pre_additions = target_word_doc.ActiveWindow.Selection.Information(4)

            # # Setting the target word doc's header footer settings
            self.__copy_header_footer_properties(target_word_doc)

            # Apply the header/footer info in the template to the target.
            for template_section in self.__template_word_doc.Sections:

                # Gather the Target's sections that match the Template's sections. 
                target_sections = [target_section for target_section in target_word_doc.Sections if target_section.Index == template_section.Index]

                target_section = target_sections[0] if len(target_sections) > 0 else None

                # Only apply the section changes if the target contains the sections to be changed.
                if target_section is None: continue

                # Copy the headers and footers from the template to the target.
                self.__copy_section_items(template_section.Headers, target_section.Headers)

                self.__copy_section_items(template_section.Footers, target_section.Footers)

            # Review the pages to see if there's been uncessary additional pages added.
            self.__delete_trailing_empty_lines(target_word_doc, num_pages_pre_additions)

            # Save the changes.
            target_word_doc.Save()

            target_word_doc.Close()

    def __copy_header_footer_properties(self, target_word_doc):
        '''Copies the header/footer properties from the template doc to the target doc. E.g. Header distance.
        
        Arguments:
            target_doc {Document} -- The target Word Document to make the changes to.
        '''

        logging.debug('\tCopying Header Footer properties from Template Word Doc to Target Word Doc.')

        self.__update_text_callback('\tCopying Header Footer properties from Template Word Doc to Target Word Doc.')

        target_word_doc.PageSetup.DifferentFirstPageHeaderFooter = self.__template_word_doc.PageSetup.DifferentFirstPageHeaderFooter

        target_word_doc.PageSetup.HeaderDistance = self.__template_word_doc.PageSetup.HeaderDistance

        target_word_doc.PageSetup.FooterDistance = self.__template_word_doc.PageSetup.FooterDistance

        target_word_doc.PageSetup.LineNumbering = self.__template_word_doc.PageSetup.LineNumbering

        target_word_doc.PageSetup.OddAndEvenPagesHeaderFooter = self.__template_word_doc.PageSetup.OddAndEvenPagesHeaderFooter

    def __copy_section_items(self, template_items, target_items):
        '''Replaces the given target items with the given template items.
        
        Arguments:
            template_items {list} -- The Template items to use when making replacements.
            target_items {list} -- The Target items being replaced.
        '''

        logging.debug('\tCopying Section Items')

        self.__update_text_callback('\tCopying Section Items')

        for template_item in template_items:

            if template_item.Exists == False: continue

            target_item = [target_item for target_item in target_items if target_item.Index == template_item.Index][0]

            if target_item.Exists == False:

                # Create the item?

                pass

            template_range = template_item.Range

            target_range = target_item.Range

            template_range.Select()

            template_range.Copy()

            target_range.Select()

            target_range.Paste()

    def __delete_trailing_empty_lines(self, target_word_doc, num_pages_pre_additions):
        '''Deletes the trailing empty lines in a given document.

        This is important after applying a header/footer to a document since that process may push the trailing lines onto a document's second page, which isn't desired.

        Arguments:
            target_doc {Document} -- The target Word Doc to review.
        '''

        logging.debug('Reviewing trailing empty lines.')

        self.__update_text_callback('Reviewing trailing empty lines.')

        # Unselect whatever is selected,
        target_word_doc.ActiveWindow.Selection.EscapeKey()

        # Determine the number of pages in the document.
        current_num_pages = target_word_doc.ActiveWindow.Selection.Information(4)
        
        # If there's a discrepancy,
        if num_pages_pre_additions < current_num_pages:

            logging.debug('Deleting trailing empty lines.')

            self.__update_text_callback('Deleting trailing empty lines.')

            paragraphs = target_word_doc.Paragraphs

            last_non_empty_paragraph_index = -1

            for paragraph_index, paragraph in enumerate(paragraphs):

                if len(paragraph.Range.Text.strip()) > 0: last_non_empty_paragraph_index = paragraph_index

            for paragraph_index, paragraph in enumerate(paragraphs):

                if paragraph_index < last_non_empty_paragraph_index + 1: continue

                paragraph.Range.Delete()