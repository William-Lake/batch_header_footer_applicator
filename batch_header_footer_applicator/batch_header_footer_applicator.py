import logging
from header_footer_applicator import HeaderFooterApplicator
from primary_ui import PrimaryUI


def main():
    '''
    Main Method
    '''

    logging.basicConfig(level=logging.DEBUG)

    primary_ui = PrimaryUI()

    header_footer_applicator = HeaderFooterApplicator()

    primary_ui.start(header_footer_applicator.apply_header_footer)
